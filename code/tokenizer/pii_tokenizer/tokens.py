"""
Tokens, the token vault and free-text scrubbing.

A token replaces one personal value with a short stand-in such as ``EMP-7f3a91c2d0``:

    token = PREFIX + "-" + first 10 hex digits of HMAC-SHA256(key, kind + normalised value)

* **Keyed, not a plain hash.** Without the key nobody can rebuild the token for a known name and so
  re-identify a person (a plain XHA-256 of "John Smith" can be looked up by anyone).
* **Deterministic.** The same person gets the same token in every file and every week, so the checks
  that group lines by employee still work on the tokenized copy.
* **Reversible only on the Google side.** HMAC is one-way, so the vault keeps ``token -> original value``.
  The vault is encrypted (AES-256-GCM) with a key derived from the same secret and is stored in a private
  Google Cloud Storage bucket that AWS has no access to.

The key must be kept for as long as tokenized results are kept (the AWS archive keeps them 10 years):
without it (and the vault) archived results can no longer be turned back into names.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import unicodedata

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.hkdf import HKDF

TOKEN_HEX = 10                    # 40 bits: collisions are checked for and refused, see Vault.add
# Any token produced by this module, whatever its prefix (prefixes are 2-5 capital letters).
TEXT_KIND, TEXT_PREFIX = "text", "TXT"   # exact text of cells the agent doesn't read (Keyring.text_token)
TOKEN_RE = re.compile(r"\b([A-Z]{2,5})-([0-9a-f]{%d})\b" % TOKEN_HEX)


class TokenError(RuntimeError):
    """Raised for anything that must stop a file from being copied (fail closed)."""


def normalise(value) -> str:
    """The form a value is hashed in: so 'SMITH,  John ' and 'smith, john' get the same token."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)                       # 12345.0 read from a spreadsheet -> "12345"
    s = unicodedata.normalize("NFKC", str(value)).strip()
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def load_key(raw: bytes | str) -> bytes:
    """Accept the secret as raw bytes, or as base64 / hex text (as stored in Secret Manager)."""
    if isinstance(raw, bytes):
        try:
            raw = raw.decode("ascii")
        except UnicodeDecodeError:
            key = raw
            raw = None
    if raw is not None:
        text = raw.strip()
        try:
            key = bytes.fromhex(text)
        except ValueError:
            key = base64.b64decode(text, validate=True)
    if len(key) < 32:
        raise TokenError("The tokenization key must be at least 32 bytes (256 bits).")
    return key


class Keyring:
    """Makes tokens from values. Holds the HMAC key and the separate vault-encryption key."""

    def __init__(self, key: bytes):
        self._mac_key = _derive(key, b"pii-tokenizer/hmac/v1")
        self.vault_key = _derive(key, b"pii-tokenizer/vault/v1")

    def token(self, prefix: str, kind: str, value) -> str:
        msg = f"{kind}\x1f{normalise(value)}".encode("utf-8")
        digest = hmac.new(self._mac_key, msg, hashlib.sha256).hexdigest()
        return f"{prefix}-{digest[:TOKEN_HEX]}"

    def text_token(self, text: str) -> str:
        """A token for a piece of text kept exactly as written (spacing and line breaks included): the text of a
        cell the agent doesn't read, replaced because it might hold personal data (see xlsx.py, "if in doubt")."""
        msg = f"{TEXT_KIND}\x1f{text}".encode("utf-8")
        digest = hmac.new(self._mac_key, msg, hashlib.sha256).hexdigest()
        return f"{TEXT_PREFIX}-{digest[:TOKEN_HEX]}"


    def exact_token(self, prefix: str, kind: str, text: str) -> str:
        """A token for text kept exactly as written, of its own kind (#53: words of a file or folder name, FIL-)."""
        msg = f"{kind}\x1f{text}".encode("utf-8")
        digest = hmac.new(self._mac_key, msg, hashlib.sha256).hexdigest()
        return f"{prefix}-{digest[:TOKEN_HEX]}"


def _derive(key: bytes, info: bytes) -> bytes:
    return HKDF(algorithm=hashes.SHA256(), length=32, salt=None, info=info).derive(key)


class Vault:
    """token -> (kind, normalised value, original value, original type).

    The original value is the first spelling seen; detokenizing gives that spelling back.
    ``dirty`` tells the caller the vault must be saved."""

    def __init__(self, entries: dict | None = None):
        self.entries: dict[str, list] = entries or {}
        self.dirty = False

    def add_text(self, token: str, text: str) -> None:
        """Keep a text token's text exactly as written (no normalising, no trimming), so it comes back unchanged."""
        known = self.entries.get(token)
        if known is None:
            self.entries[token] = [TEXT_KIND, text, text, "str"]
            self.dirty = True
        elif known[0] != TEXT_KIND or known[1] != text:
            raise TokenError(f"Token collision on {token}; increase TOKEN_HEX before continuing.")

    def add_exact(self, token: str, kind: str, text: str) -> None:
        """As add_text, for an exact-text token of another kind (Keyring.exact_token)."""
        known = self.entries.get(token)
        if known is None:
            self.entries[token] = [kind, text, text, "str"]
            self.dirty = True
        elif known[0] != kind or known[1] != text:
            raise TokenError(f"Token collision on {token}; increase TOKEN_HEX before continuing.")

    def add(self, token: str, kind: str, value) -> None:
        norm = normalise(value)
        known = self.entries.get(token)
        if known is None:
            vtype = "int" if isinstance(value, int) or (isinstance(value, float) and value.is_integer()) else \
                    "float" if isinstance(value, float) else "str"
            shown = str(int(value)) if vtype == "int" else str(value).strip()
            self.entries[token] = [kind, norm, shown, vtype]
            self.dirty = True
        elif known[0] != kind or known[1] != norm:
            # Two different values hashed to the same 40-bit token. Astronomically unlikely for a payroll
            # roster, but refusing is the only safe answer: a wrong name in a result would be worse.
            raise TokenError(f"Token collision on {token}; increase TOKEN_HEX before continuing.")

    def merge(self, other: "Vault") -> None:
        """Take in tokens another run saved meanwhile (refusing a collision, as add does)."""
        for tok, e in other.entries.items():
            mine = self.entries.get(tok)
            if mine is None:
                self.entries[tok] = e
            elif mine[0] != e[0] or mine[1] != e[1]:
                raise TokenError(f"Token collision on {tok}; increase TOKEN_HEX before continuing.")
        self.dirty = True

    def original(self, token: str):
        e = self.entries.get(token)
        if e is None:
            return None
        if e[3] == "int":
            return int(e[2])
        if e[3] == "float":
            return float(e[2])
        return e[2]

    def names(self, kinds: set[str]) -> list[tuple[str, str, str]]:
        """(original value, token, kind) for the kinds whose values are searched for in free text."""
        return [(e[2], t, e[0]) for t, e in self.entries.items() if e[0] in kinds]

    # --- encrypted storage ------------------------------------------------------------------------
    def to_bytes(self, keyring: Keyring) -> bytes:
        nonce = os.urandom(12)
        plain = json.dumps({"version": 1, "entries": self.entries}, ensure_ascii=False,
                           sort_keys=True).encode("utf-8")
        ct = AESGCM(keyring.vault_key).encrypt(nonce, plain, b"pii-tokenizer-vault")
        return json.dumps({"format": "pii-tokenizer-vault/1", "nonce": base64.b64encode(nonce).decode(),
                           "data": base64.b64encode(ct).decode()}).encode()

    @classmethod
    def from_bytes(cls, blob: bytes | None, keyring: Keyring) -> "Vault":
        if not blob:
            return cls()
        env = json.loads(blob)
        try:
            plain = AESGCM(keyring.vault_key).decrypt(base64.b64decode(env["nonce"]),
                                                      base64.b64decode(env["data"]), b"pii-tokenizer-vault")
        except Exception as e:  # wrong key or damaged file: never start a fresh vault over it
            raise TokenError("Cannot decrypt the token vault (wrong key, or the file is damaged).") from e
        return cls(json.loads(plain)["entries"])


# ------------------------------------------------------------------------------------------------
# Free-text scrubbing
# ------------------------------------------------------------------------------------------------
MIN_DIGITS_IN_FREE_TEXT = 5
# An SSN written without dashes is only recognised next to a label ("SSN 123456789", "Social Security No: ..."):
# other 9-digit numbers (project, permit and decision numbers) are common. The label and number must be on one line.
SSN_LABELLED = re.compile(r"(?i)(?:\bssn\b|\bss\s?#|\bsoc(?:ial)?\.?\s*sec(?:urity)?\b(?:\s*(?:no\b|num\b|number\b|#))?)"
                          r"[^\d\n\x00]{0,15}(?<!\d)\d{9}(?!\d)")
_WORD_RE = re.compile(r"\w+")
_VALUE_WORD_RE = re.compile(r"[\w-]+")
_BARE_NUMBER_RE = re.compile(r"[\d.,\s-]*")


class Scrubber:
    """Replaces personal data inside free text (notes, comments, cached formula results).

    * Known names (from the name columns of this file and from the vault) are replaced wherever they
      appear as whole words. Only names of at least ``min_name_words`` words are scanned for, so a
      one-word value such as a last name that is also a trade ("Mason") is not replaced everywhere.
    * Patterns (SSN, e-mail, phone) are replaced with tokens of their own kind.
    """

    def __init__(self, keyring: Keyring, vault: Vault, patterns: dict, name_kinds: set[str],
                 min_name_words: int = 2, value_kinds: set[str] = frozenset()):
        self.k, self.vault = keyring, vault
        # Header labels seen in a header row this run (saw_label): never learnt for free text (see saw_label)
        self._labels: set[str] = set()
        self.patterns = [(p["prefix"], kind, re.compile(p["regex"])) for kind, p in patterns.items()]
        self.name_kinds, self.min_words = name_kinds, min_name_words
        self._names: dict[str, str] = {}
        self._names_order: list[tuple[str, str]] = []       # the same, in the order learnt (append-only)
        self._name_re = None
        self._values: dict[str, str] = {}
        self._value_re = None
        self._first_words = self._value_keys = None
        self._memo: dict[str, str] = {}
        self._always_names = self._always_values = False
        for value, tok, kind in vault.names(name_kinds):
            self.add_name(value, tok, kind)
        for value, tok, kind in vault.names(set(value_kinds)):
            self.add_value(value, tok)

    @staticmethod
    def _label_key(text) -> str:
        return re.sub(r"\s+", " ", str(text)).strip().casefold()

    def _is_label(self, value) -> bool:
        return isinstance(value, str) and bool(self._labels) and self._label_key(value) in self._labels

    def saw_label(self, text) -> None:
        """``text`` is a header label of a header row (HeaderTracker found it there). The same text in a data cell
        (a payroll with its header row pasted twice holds "Employee Name" in the Employee Name column) is still
        tokenized where it stands, but never learnt for free text: learnt, it replaced that label in every later
        file of the run, so no payroll was recognized (found by the evaluation harness, Nov 13, 2026). A label
        learnt before (earlier in the run, or from the vault) is forgotten. Only labels actually seen in a header
        row count (HeaderTracker: two or more labels in the row, none with a digit), and never one with a digit: a
        header pattern alone would also match real values such as "EMP#10432"."""
        key = self._label_key(text)
        if not key or key in self._labels or re.search(r"\d", key):
            return                                  # a label with a digit could be an ID: never skipped (verification)
        self._labels.add(key)
        if key in self._names:
            del self._names[key]
            self._names_order = [(n, t) for n, t in self._names_order if n != key]
            self._name_re = self._first_words = None
            for cache in ("_tuples_n", "_ni_keys"):         # name_spans' and _name_index's: rebuilt when next used
                self.__dict__.pop(cache, None)
            self._memo.clear()
        if key in self._values:
            del self._values[key]
            self._value_re = self._value_keys = None
            self._memo.clear()

    def add_value(self, value, token: str) -> None:
        """Search free text for an ID or SSN (4+ characters, as a whole word). An SSN is also searched for
        without its dashes."""
        if self._is_label(value):
            return                                  # a header label (see saw_label)
        v = str(int(value)) if isinstance(value, float) and value.is_integer() else str(value).strip()
        forms = {v}
        if re.fullmatch(r"\d{3}-\d{2}-\d{4}", v):
            forms.add(v.replace("-", ""))
        for f in forms:
            # A short all-digit value (an employee ID such as 2025, the last 4 of an SSN) would also match years,
            # dates, amounts and wage decision numbers everywhere ("WA L&I 2025", "11/10/2025"). It is still
            # tokenized in its own column; in free text only all-digit values of 5+ digits are searched for.
            if f.isdigit() and len(f) < MIN_DIGITS_IN_FREE_TEXT:
                continue
            if len(f) >= 4 and f.casefold() not in self._values:
                self._values[f.casefold()] = token
                self._value_re = self._value_keys = None
                self._memo.clear()

    def add_name(self, value: str, token: str, kind: str = "employee_name") -> None:
        """Search free text for ``value`` (replaced by ``token``). "Smith, John" is also searched for as
        "John Smith", which gets a token of its own so that detokenizing gives back the exact wording."""
        if self._is_label(value):
            return                                  # a header label, not a name (see saw_label)
        value = re.sub(r"\s+", " ", str(value)).strip()
        variants = {value: token}
        if "," in value:
            last, _, first = value.partition(",")
            flipped = f"{first.strip()} {last.strip()}"
            vtok = self.k.token(token.split("-")[0], kind, flipped)
            self.vault.add(vtok, kind, flipped)
            variants[flipped] = vtok
        for v, tok in list(variants.items()):
            words = [w for w in re.split(r"[\s,]+", v) if w]
            if any(len(w.strip(".")) == 1 for w in words):        # "Lee J Mason": also look for "Lee Mason"
                short = " ".join(w for w in re.split(r"\s+", v) if len(w.strip(".,")) > 1)
                if short and short not in variants:
                    stok = self.k.token(token.split("-")[0], kind, short)
                    self.vault.add(stok, kind, short)
                    variants[short] = stok
        for v, tok in variants.items():
            words = [w for w in re.split(r"[\s,]+", v) if w]
            if sum(len(w.strip(".")) >= 2 for w in words) >= self.min_words:
                if v.casefold() not in self._names:
                    self._names[v.casefold()] = tok
                    self._names_order.append((v.casefold(), tok))
                    self._name_re = self._first_words = None
                    self._memo.clear()

    def _compiled(self):
        if self._name_re is None and self._names:
            alts = sorted(self._names, key=len, reverse=True)   # longest first: "John A Smith" before "John"
            # Whitespace inside a name may be any run of spaces in the text.
            # Any run of spaces between words, and an optional middle initial ("Ana B. Rivera").
            sep = r"\s+(?:[A-Za-z]\.?\s+)?"
            body = "|".join(sep.join(map(re.escape, a.split(" "))) for a in alts)
            self._name_re = re.compile(rf"(?<!\w)(?:{body})(?!\w)", re.IGNORECASE)
        return self._name_re

    def _name_token(self, m) -> str:
        words = m.group(0).split()
        key = " ".join(words).casefold()
        if key not in self._names:                     # matched with a middle initial: drop it to look up
            key = " ".join(w for w in words if len(w.strip(".")) > 1).casefold()
        return self._names.get(key, m.group(0))

    def _may_hold_name(self, text: str) -> bool:
        """Cheap pre-check: a name can only match where the text has, as a whole word, the first word of some
        known name. Most cells (trades, dates, amounts) have none, so the big search is skipped for them."""
        if self._first_words is None:
            firsts, always = set(), False
            for n in self._names:
                m = _WORD_RE.match(n)
                if not m or not n.isascii():
                    always = True               # unusual first character or non-ASCII: no shortcut
                    break
                firsts.add(m.group(0))
            self._first_words, self._always_names = firsts, always
        if self._always_names or not text.isascii():
            return True
        return any(w in self._first_words for w in _WORD_RE.findall(text.lower()))

    def name_places(self, text: str) -> list[tuple[int, int, str]]:
        """Where learnt names are in a file or folder name: (start, end, token), longest first, left to right.
        Words are compared after Unicode normalisation (accents written apart, full-width letters) and case folding,
        whatever separates them (spaces, "_", ".", "-", ",", capitals: "RiveraAna", "McDonald_John"), in any order
        of the name's parts, with a word written in two ("Mc Donald", "O'Brien") and up to two middle initials
        skipped between two of the name's words. Whole words only: "Erica Lee" is not "Eric Lee"."""
        keys, prefixes, bounds = self._name_index()
        words = name_words(text)
        out, i = [], 0
        while i < len(words):
            best, stack = None, [(i, "", 0)]
            while stack:
                j, acc, skipped = stack.pop()
                if j >= len(words):
                    continue
                w = words[j][2]
                if acc + w in prefixes:
                    if acc + w in keys and (best is None or j > best[0]):
                        best = (j, keys[acc + w])
                    stack.append((j + 1, acc + w, skipped))
                if len(w) == 1 and acc in bounds and skipped < 2:       # "Ana B. C. Rivera": initials skipped
                    stack.append((j + 1, acc, skipped + 1))
            if best:
                out.append((words[i][0], words[best[0]][1], best[1]))
                i = best[0] + 1
            else:
                i += 1
        return out

    def name_letters_in(self, text: str) -> bool:
        """True if the letters of any learnt name, in any order it may be written, appear anywhere in ``text``,
        word breaks ignored ("TSAnaRivera", "xanariveray"). The strict check behind tokenize_name: a file whose name
        still matches is held back (rarely a false alarm, such as "Banana Rivera"; never a name sent)."""
        keys, prefixes, _ = self._name_index()
        letters = "".join(w for _, _, w in name_words(text))
        for s in range(len(letters)):
            e = s + 1
            while e <= len(letters) and letters[s:e] in prefixes:
                if letters[s:e] in keys and e - s >= 6:
                    return True
                e += 1
        return False

    def _name_index(self):
        """Every learnt name in every word order, as letters only: the order as learnt first (its own token), then
        the others. Grows with the names (they are only ever added), never rebuilt."""
        if "_ni_keys" not in self.__dict__:
            self._ni_keys, self._ni_pre, self._ni_bounds, self._ni_exact, self._ni_n = {}, set(), set(), set(), 0
        new = self._names_order[self._ni_n:]              # only the names learnt since the last call
        for n, tok in new:
            for k, order in enumerate(_name_orders(n)):
                key = "".join(order)
                if len(key) < 4:
                    continue
                if k == 0 and key not in self._ni_exact:
                    self._ni_keys[key] = tok
                    self._ni_exact.add(key)
                else:
                    self._ni_keys.setdefault(key, tok)
                self._ni_pre.update(key[:m] for m in range(1, len(key) + 1))
                acc = ""
                for w in order[:-1]:
                    acc += w
                    self._ni_bounds.add(acc)
        self._ni_n += len(new)
        return self._ni_keys, self._ni_pre, self._ni_bounds

    def _may_hold_value(self, text: str) -> bool:
        if self._value_keys is None:
            self._value_keys = set(self._values)
            self._always_values = any(not _VALUE_WORD_RE.fullmatch(v) for v in self._values)
        if self._always_values or not text.isascii():
            return True
        return any(w in self._value_keys for w in _VALUE_WORD_RE.findall(text.lower()))

    def finds(self, text: str, contacts: bool = False, names: bool = True) -> str | None:
        """What worker data ``text`` holds, without changing it or the vault: "an SSN", "a worker's name" (one
        learnt from a payroll), and with ``contacts`` also "an email address" or "a phone number"; None if none.
        Used to decide whether a file (or its name) may go to AWS at all (see screen.py)."""
        if not text:
            return None
        text = TOKEN_RE.sub(" ", text)          # tokens are what we want to see: "SSN-e412240181" is not an SSN
        for _prefix, kind, pat in self.patterns:
            if kind == "ssn" and pat.search(text):
                return "an SSN"
        if SSN_LABELLED.search(text):
            return "an SSN"
        rx = self._compiled() if names else None
        if rx is not None and self._may_hold_name(text) and rx.search(text):
            return "a worker's name"
        if contacts:
            for _prefix, kind, pat in self.patterns:
                if kind in ("email", "phone") and pat.search(text):
                    return "an email address" if kind == "email" else "a phone number"
        return None

    def name_in_words(self, words: list[str]) -> bool:
        """True if a learnt name of two or more words appears in ``words`` (lower case), its words next to each other
        in either order: "rivera ana" and "ana rivera" both match "Rivera, Ana". For file and folder names, where
        "Ana_Rivera", "Rivera-Ana" or "RiveraAna" are common (see screen.name_left)."""
        return bool(self.name_spans(words))

    def name_spans(self, words: list[str]) -> list[tuple[int, int, str]]:
        """Where learnt names are in ``words`` (lower case): (first word, last word + 1, token), longest first,
        left to right, not overlapping. The token is the one of the name as written there: "ana rivera" gets the
        token of "Ana Rivera", "rivera ana" the one of "Rivera, Ana" (so it comes back in that order)."""
        if self.__dict__.get("_tuples_n") != len(self._names):
            tuples = {}
            for n, tok in self._names.items():
                ws = tuple(w for w in re.split(r"[\s,]+", n) if w)
                if len(ws) >= 2:
                    tuples[ws] = tok
            for ws, tok in list(tuples.items()):
                tuples.setdefault(tuple(reversed(ws)), tok)
            self._tuples, self._tuples_n = tuples, len(self._names)
            self._tuple_lens = sorted({len(t) for t in tuples}, reverse=True)
        out, i = [], 0
        while i < len(words):
            for n in self._tuple_lens:
                tok = self._tuples.get(tuple(words[i:i + n])) if i + n <= len(words) else None
                if tok:
                    out.append((i, i + n, tok))
                    i += n
                    break
            else:
                i += 1
        return out

    @property
    def name_count(self) -> int:
        """How many name spellings are searched for; grows as payrolls teach new names."""
        return len(self._names)

    def scrub(self, text: str) -> str:
        """Same text, same result (until new names or IDs are learnt), so results are remembered: an export
        that keeps every text inline repeats the same trade or contractor in hundreds of thousands of cells."""
        if not text:
            return text
        hit = self._memo.get(text)
        if hit is not None:
            return hit
        out = self._scrub(text)
        if len(self._memo) > 500_000:
            self._memo.clear()
        self._memo[text] = out
        return out

    def _scrub(self, text: str) -> str:
        rx = self._compiled()
        if rx is not None and self._may_hold_name(text):
            text = rx.sub(self._name_token, text)
        if self._values and not _BARE_NUMBER_RE.fullmatch(text) and self._may_hold_value(text):
            if self._value_re is None:
                alts = sorted(self._values, key=len, reverse=True)
                self._value_re = re.compile(r"(?<![\w-])(?:%s)(?![\w-])" % "|".join(map(re.escape, alts)),
                                            re.IGNORECASE)
            text = self._value_re.sub(lambda m: self._values[m.group(0).casefold()], text)
        for prefix, kind, pat in self.patterns:
            def repl(m, prefix=prefix, kind=kind):
                tok = self.k.token(prefix, kind, m.group(0))
                self.vault.add(tok, kind, m.group(0))
                return tok
            text = pat.sub(repl, text)
        return text


# Letters compared as their plain Latin equivalent in names: ones that don't split into a base letter and an accent
# ("Łukasz" = "Lukasz", "Søren" = "Soren", "Yılmaz" = "Yilmaz"), and Cyrillic and Greek letters that look Latin
# ("Аna" with a Cyrillic А). Applied to both the learnt names and the file names.
_FOLD = str.maketrans({
    "ł": "l", "ø": "o", "ı": "i", "đ": "d", "ð": "d", "ħ": "h", "ŧ": "t", "ŀ": "l", "ĸ": "k", "ŋ": "n",
    "æ": "ae", "œ": "oe", "þ": "th", "ĳ": "ij",
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "і": "i", "ј": "j", "ѕ": "s",
    "ԁ": "d", "һ": "h", "к": "k", "м": "m", "н": "h", "т": "t", "в": "b", "ё": "e", "ї": "i", "ԛ": "q", "ԝ": "w",
    "α": "a", "ε": "e", "ο": "o", "ρ": "p", "τ": "t", "υ": "u", "ν": "v", "κ": "k", "ι": "i", "χ": "x", "μ": "m",
    "ϲ": "c", "ϳ": "j", "β": "b", "η": "n",
})


def _is_word_char(ch: str) -> bool:
    c = unicodedata.category(ch)
    return (c[0] == "L" and c != "Lm") or c[0] == "M"           # letters and their accents; not "ʼ"


def name_words(text: str) -> list[tuple[int, int, str]]:
    """The words of a file or folder name: (start, end, lower-case letters without accents). Split at anything that
    isn't a letter and where a capital follows a small letter ("RiveraAna")."""
    out, i, n = [], 0, len(text)
    def prev_letter(j):                                # the letter before j, past any accent written apart
        k = j - 1
        while k > i and unicodedata.combining(text[k]):
            k -= 1
        return text[k]

    def cut(j):                                        # a word ends before text[j]
        a, b = prev_letter(j), text[j]
        if a.islower() and b.isupper():
            return True                                # "anaRivera"
        nxt = text[j + 1] if j + 1 < n else ""
        if a.isupper() and b.isupper() and nxt.islower():
            return True                                # "TSAna": "TS" + "Ana"
        if a.isupper() and b.islower() and j >= 2 and text[j - 2].isupper():
            return True                                # "RIVERAhours": "RIVERA" + "hours"
        return False
    while i < n:
        if not _is_word_char(text[i]):
            i += 1
            continue
        j = i + 1
        while j < n and _is_word_char(text[j]) and not cut(j):
            j += 1
        w = "".join(ch for ch in unicodedata.normalize("NFKD", text[i:j].casefold())       # accents dropped:
                    if _is_word_char(ch) and not unicodedata.combining(ch)).translate(_FOLD)  # "Pena" is "Peña"
        if w:
            out.append((i, j, w))
        i = j
    return out


def _name_orders(name: str) -> list[list[str]]:
    """The word orders a learnt name may be written in, the order as learnt first. "Rivera, Ana": last then first,
    and first then last; "Mary-Jane Watson" (no comma): every rotation."""
    name = unicodedata.normalize("NFKC", name).casefold()
    if "," in name:
        last, _, first = name.partition(",")
        lw, fw = [w for _, _, w in name_words(last)], [w for _, _, w in name_words(first)]
        orders = [lw + fw, fw + lw] if lw and fw else []
    else:
        ws = [w for _, _, w in name_words(name)]
        orders = [ws[k:] + ws[:k] for k in range(len(ws))]
    return [o for o in orders if len(o) >= 2]


# A token next to the space #53's opaque names put there because it was glued to a letter ("ProjInfo" -> "TXT-… Info"):
# its kind ends with "+b", "+a" or "+ba" (pii_tokenizer/names.py _kind), and the way back takes that space out
_GLUE_RE = re.compile(r"( ?)\b([A-Z]{2,5}-[0-9a-f]{%d})\b( ?)" % TOKEN_HEX)


def unglue(text: str, vault: Vault, esc=str) -> str:
    """``text`` with each glued token (see _GLUE_RE) put back with the space next to it taken out (``esc``: how the
    value is written, e.g. XML-escaped)."""
    if not text or "-" not in text:
        return text

    def repl(m):
        e = vault.entries.get(m.group(2))
        if not e or "+" not in str(e[0]):
            return m.group(0)
        glue = str(e[0]).rsplit("+", 1)[1]
        before = "" if "b" in glue else m.group(1)
        after = "" if "a" in glue else m.group(3)
        return before + esc(str(e[2])) + after
    return _GLUE_RE.sub(repl, text)


def detokenize_text(text: str, vault: Vault, unknown: set | None = None) -> str:
    """Replace every token in ``text`` with its original value. Unknown tokens are left as they are
    (and collected in ``unknown`` so the caller can report them)."""
    text = unglue(text, vault)
    def repl(m):
        val = vault.original(m.group(0))
        if val is None:
            if unknown is not None:
                unknown.add(m.group(0))
            return m.group(0)
        return str(val)
    return TOKEN_RE.sub(repl, text)
