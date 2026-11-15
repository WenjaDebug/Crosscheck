"""Zertifikate für das Seitenfenster ohne openssl, nur mit der Standardbibliothek (für Windows).

Am Mac legt openssl die Zertifikate an. Windows hat meist kein openssl, und ein Paket dafür soll Gegenprobe nicht
brauchen. Darum steht hier alles Nötige von Hand: ein RSA-Schlüssel mit 2048 Bit (Primzahlen nach Miller-Rabin), das
Zertifikat in DER/ASN.1 und die Unterschrift mit SHA-256 nach PKCS #1 v1.5. Es entsteht dasselbe wie mit openssl und
_CA_CONFIG/_LEAF_EXTENSIONS in addin.py: eine Zertifizierungsstelle nur für diesen Rechner und damit ein Zertifikat
für localhost und 127.0.0.1.
"""

from __future__ import annotations

import base64
import hashlib
import math
import os
import re
import secrets
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

E = 65537
BITS = 2048

# Kennungen (OID), wie sie in X.509 stehen.
RSA = "1.2.840.113549.1.1.1"
SHA256_RSA = "1.2.840.113549.1.1.11"
SHA256 = "2.16.840.1.101.3.4.2.1"
COMMON_NAME = "2.5.4.3"
BASIC_CONSTRAINTS = "2.5.29.19"
KEY_USAGE = "2.5.29.15"
EXTENDED_KEY_USAGE = "2.5.29.37"
SUBJECT_ALT_NAME = "2.5.29.17"
SUBJECT_KEY_ID = "2.5.29.14"
AUTHORITY_KEY_ID = "2.5.29.35"
SERVER_AUTH = "1.3.6.1.5.5.7.3.1"


# --- DER schreiben ---


def _tlv(tag: int, content: bytes) -> bytes:
    size = len(content)
    if size < 0x80:
        return bytes([tag, size]) + content
    raw = size.to_bytes((size.bit_length() + 7) // 8, "big")
    return bytes([tag, 0x80 | len(raw)]) + raw + content


def _integer(value: int) -> bytes:
    # Eine positive Zahl, deren erstes Bit gesetzt ist, braucht vorne eine 0, sonst wäre sie negativ.
    raw = value.to_bytes(value.bit_length() // 8 + 1, "big") if value else b"\x00"
    return _tlv(0x02, raw)


def _oid(text: str) -> bytes:
    parts = [int(p) for p in text.split(".")]
    out = bytearray([parts[0] * 40 + parts[1]])
    for part in parts[2:]:
        chunk = [part & 0x7F]
        part >>= 7
        while part:
            chunk.append(0x80 | (part & 0x7F))
            part >>= 7
        out += bytes(reversed(chunk))
    return _tlv(0x06, bytes(out))


def _seq(*items: bytes) -> bytes:
    return _tlv(0x30, b"".join(items))


def _set(*items: bytes) -> bytes:
    return _tlv(0x31, b"".join(items))


def _bits(data: bytes, unused: int = 0) -> bytes:
    return _tlv(0x03, bytes([unused]) + data)


def _octets(data: bytes) -> bytes:
    return _tlv(0x04, data)


def _null() -> bytes:
    return b"\x05\x00"


def _true() -> bytes:
    return b"\x01\x01\xff"


def _time(when: datetime) -> bytes:
    # Bis 2049 UTCTime, danach GeneralizedTime (RFC 5280, 4.1.2.5).
    when = when.astimezone(timezone.utc)
    if when.year < 2050:
        return _tlv(0x17, when.strftime("%y%m%d%H%M%SZ").encode("ascii"))
    return _tlv(0x18, when.strftime("%Y%m%d%H%M%SZ").encode("ascii"))


def _name(common_name: str) -> bytes:
    return _seq(_set(_seq(_oid(COMMON_NAME), _tlv(0x0C, common_name.encode("utf-8")))))


def _extension(oid: str, value: bytes, critical: bool = False) -> bytes:
    return _seq(_oid(oid), *((_true(),) if critical else ()), _octets(value))


def _algorithm(oid: str) -> bytes:
    return _seq(_oid(oid), _null())


# --- DER lesen ---


def _read(data: bytes, pos: int = 0) -> tuple[int, int, int]:
    """(Tag, Anfang des Inhalts, Ende) des Elements ab pos."""
    tag, size = data[pos], data[pos + 1]
    pos += 2
    if size & 0x80:
        count = size & 0x7F
        if not 0 < count <= 4:
            raise ValueError("DER: Länge passt nicht")
        size = int.from_bytes(data[pos:pos + count], "big")
        pos += count
    if pos + size > len(data):
        raise ValueError("DER: zu kurz")
    return tag, pos, pos + size


def _children(data: bytes, start: int, end: int) -> list[tuple[int, int, int, int]]:
    """Die Elemente zwischen start und end: (Tag, Beginn des Elements, Anfang des Inhalts, Ende)."""
    out = []
    pos = start
    while pos < end:
        tag, inner, stop = _read(data, pos)
        out.append((tag, pos, inner, stop))
        pos = stop
    return out


def _oid_text(raw: bytes) -> str:
    first = raw[0]
    parts = [min(first // 40, 2), first - 40 * min(first // 40, 2)]
    value = 0
    for byte in raw[1:]:
        value = (value << 7) | (byte & 0x7F)
        if not byte & 0x80:
            parts.append(value)
            value = 0
    return ".".join(str(p) for p in parts)


# --- RSA ---


_SMALL = [p for p in range(3, 2000, 2) if all(p % q for q in range(3, int(p ** 0.5) + 1, 2))]


def _probably_prime(n: int, rounds: int = 40) -> bool:
    """Miller-Rabin mit zufälligen Basen; eine zusammengesetzte Zahl übersteht 40 Runden mit höchstens 4^-40."""
    if n < 2:
        return False
    for p in [2, *_SMALL]:
        if n % p == 0:
            return n == p
    d, s = n - 1, 0
    while d % 2 == 0:
        d //= 2
        s += 1
    for _ in range(rounds):
        a = secrets.randbelow(n - 3) + 2
        x = pow(a, d, n)
        if x in (1, n - 1):
            continue
        for _ in range(s - 1):
            x = pow(x, 2, n)
            if x == n - 1:
                break
        else:
            return False
    return True


def _prime(bits: int) -> int:
    while True:
        # Die obersten zwei Bits gesetzt: Das Produkt zweier solcher Zahlen hat genau 2·bits Bits.
        candidate = secrets.randbits(bits) | (0b11 << (bits - 2)) | 1
        if candidate % E != 1 and _probably_prime(candidate):
            return candidate


class Key:
    """Ein RSA-Schlüssel: n, e, d und die beiden Primzahlen p, q."""

    def __init__(self, n: int, e: int, d: int, p: int, q: int) -> None:
        self.n, self.e, self.d, self.p, self.q = n, e, d, p, q

    @classmethod
    def new(cls, bits: int = BITS) -> "Key":
        while True:
            p, q = _prime(bits // 2), _prime(bits // 2)
            if p == q:
                continue
            n = p * q
            phi = (p - 1) * (q - 1)
            if n.bit_length() != bits or math.gcd(E, phi) != 1:
                continue
            d = pow(E, -1, (p - 1) * (q - 1) // math.gcd(p - 1, q - 1))
            key = cls(n, E, d, max(p, q), min(p, q))
            # Einmal hin und zurück: So fällt ein Rechenfehler hier auf, nicht erst in Excel.
            probe = secrets.randbelow(n - 2) + 2
            if pow(pow(probe, E, n), d, n) == probe:
                return key

    def public(self) -> bytes:
        """RSAPublicKey (PKCS #1)."""
        return _seq(_integer(self.n), _integer(self.e))

    def public_info(self) -> bytes:
        """SubjectPublicKeyInfo, wie es im Zertifikat steht."""
        return _seq(_algorithm(RSA), _bits(self.public()))

    def key_id(self) -> bytes:
        """Subject Key Identifier wie openssl („hash“): SHA-1 über den öffentlichen Schlüssel (RFC 5280, 4.2.1.2)."""
        return hashlib.sha1(self.public()).digest()

    def private(self) -> bytes:
        """PKCS #8, wie openssl ihn heute schreibt („BEGIN PRIVATE KEY“)."""
        rsa = _seq(_integer(0), _integer(self.n), _integer(self.e), _integer(self.d), _integer(self.p),
                   _integer(self.q), _integer(self.d % (self.p - 1)), _integer(self.d % (self.q - 1)),
                   _integer(pow(self.q, -1, self.p)))
        return _seq(_integer(0), _algorithm(RSA), _octets(rsa))

    def sign(self, data: bytes) -> bytes:
        """SHA-256 mit RSA nach PKCS #1 v1.5 (RFC 8017, 8.2)."""
        size = (self.n.bit_length() + 7) // 8
        info = _seq(_algorithm(SHA256), _octets(hashlib.sha256(data).digest()))
        block = b"\x00\x01" + b"\xff" * (size - len(info) - 3) + b"\x00" + info
        m = int.from_bytes(block, "big")
        # Mit dem chinesischen Restsatz, etwa viermal so schnell.
        dp, dq, q_inv = self.d % (self.p - 1), self.d % (self.q - 1), pow(self.q, -1, self.p)
        m1, m2 = pow(m, dp, self.p), pow(m, dq, self.q)
        s = m2 + self.q * ((q_inv * (m1 - m2)) % self.p)
        if pow(s, self.e, self.n) != m:
            raise RuntimeError("Die Unterschrift stimmt nicht")
        return s.to_bytes(size, "big")

    @classmethod
    def load(cls, der: bytes) -> "Key":
        """Aus PKCS #8 oder PKCS #1 („BEGIN RSA PRIVATE KEY“)."""
        tag, inner, end = _read(der)
        parts = _children(der, inner, end)
        if len(parts) >= 3 and parts[1][0] == 0x30 and parts[2][0] == 0x04:
            # PKCS #8: darin steckt der Schlüssel nach PKCS #1.
            _, _, start, stop = parts[2]
            return cls.load(der[start:stop])
        numbers = [int.from_bytes(der[i:j], "big") for t, _, i, j in parts if t == 0x02]
        if len(numbers) < 6:
            raise ValueError("Das ist kein RSA-Schlüssel")
        _, n, e, d, p, q = numbers[:6]
        return cls(n, e, d, p, q)


# --- PEM ---


def pem(kind: str, der: bytes) -> str:
    body = base64.b64encode(der).decode("ascii")
    lines = "\n".join(body[i:i + 64] for i in range(0, len(body), 64))
    return f"-----BEGIN {kind}-----\n{lines}\n-----END {kind}-----\n"


def unpem(text: str) -> bytes:
    """Der erste Block einer PEM-Datei als DER."""
    found = re.search(r"-----BEGIN ([A-Z ]+)-----(.*?)-----END \1-----", text, re.S)
    if not found:
        raise ValueError("Keine PEM-Datei")
    return base64.b64decode("".join(found.group(2).split()))


def _write(path: Path, text: str, private: bool = False) -> None:
    """Erst daneben schreiben, dann ersetzen: Ein halbes Zertifikat bleibt nie liegen."""
    with tempfile.NamedTemporaryFile("w", encoding="ascii", dir=path.parent, delete=False, suffix=".tmp",
                                     newline="\n") as file:
        file.write(text)
    if private:
        os.chmod(file.name, 0o600)
    Path(file.name).replace(path)


# --- Zertifikate ---


def _serial() -> int:
    # Positiv und nie 0; 12 Byte wie mit openssl -set_serial in addin.py.
    return secrets.randbits(95) | 1


def _certificate(subject: bytes, key: Key, issuer: bytes, signer: Key, days: int, extensions: list[bytes],
                 now: datetime | None = None) -> bytes:
    now = (now or datetime.now(timezone.utc)).replace(microsecond=0)
    tbs = _seq(
        _tlv(0xA0, _integer(2)),  # Version 3
        _integer(_serial()),
        _algorithm(SHA256_RSA),
        issuer,
        _seq(_time(now), _time(now + timedelta(days=days))),
        subject,
        key.public_info(),
        _tlv(0xA3, _seq(*extensions)),
    )
    return _seq(tbs, _algorithm(SHA256_RSA), _bits(signer.sign(tbs)))


def authority(cert: Path, key_path: Path, name: str, days: int = 3650, key: Key | None = None) -> None:
    """Die Zertifizierungsstelle: basicConstraints critical CA:TRUE, keyUsage critical keyCertSign und cRLSign."""
    key = key or Key.new()
    subject = _name(name)
    der = _certificate(subject, key, subject, key, days, [
        _extension(BASIC_CONSTRAINTS, _seq(_true()), critical=True),
        # keyCertSign (Bit 5) und cRLSign (Bit 6): 0000 0110, das letzte Bit ungenutzt.
        _extension(KEY_USAGE, _bits(b"\x06", 1), critical=True),
        _extension(SUBJECT_KEY_ID, _octets(key.key_id())),
        _extension(AUTHORITY_KEY_ID, _seq(_tlv(0x80, key.key_id()))),
    ])
    _write(key_path, pem("PRIVATE KEY", key.private()), private=True)
    _write(cert, pem("CERTIFICATE", der))


def server(cert: Path, key_path: Path, ca: Path, ca_key: Path, days: int, key: Key | None = None) -> None:
    """Das Zertifikat für localhost und 127.0.0.1, unterschrieben von der Stelle (wie _LEAF_EXTENSIONS)."""
    key = key or Key.new()
    signer = Key.load(unpem(ca_key.read_text(encoding="ascii")))
    authority_der = unpem(ca.read_text(encoding="ascii"))
    info = _parts(authority_der)
    # Die Ausstellerin genau so, wie sie in ihrem eigenen Zertifikat steht.
    issuer = authority_der[info["subject"][0]:info["subject"][1]]
    issuer_id = info.get("key_id") or signer.key_id()
    names = _seq(_tlv(0x82, b"localhost"), _tlv(0x87, bytes([127, 0, 0, 1])))
    der = _certificate(_name("localhost"), key, issuer, signer, days, [
        _extension(BASIC_CONSTRAINTS, _seq()),
        # digitalSignature (Bit 0) und keyEncipherment (Bit 2): 1010 0000, die letzten fünf Bits ungenutzt.
        _extension(KEY_USAGE, _bits(b"\xa0", 5), critical=True),
        _extension(EXTENDED_KEY_USAGE, _seq(_oid(SERVER_AUTH))),
        _extension(SUBJECT_ALT_NAME, names),
        _extension(SUBJECT_KEY_ID, _octets(key.key_id())),
        _extension(AUTHORITY_KEY_ID, _seq(_tlv(0x80, issuer_id))),
    ])
    _write(key_path, pem("PRIVATE KEY", key.private()), private=True)
    _write(cert, pem("CERTIFICATE", der))


def _parts(der: bytes) -> dict:
    """Aus einem Zertifikat: wo Subject steht, das Ende der Gültigkeit, die Erweiterungen und die eigene Kennung."""
    _, inner, end = _read(der)
    _, _, tbs_inner, tbs_end = _children(der, inner, end)[0]
    fields = _children(der, tbs_inner, tbs_end)
    if fields and fields[0][0] == 0xA0:
        fields = fields[1:]
    # serialNumber, signature, issuer, validity, subject, subjectPublicKeyInfo, dann [1] [2] [3]
    _, _, valid_inner, valid_end = fields[3]
    times = _children(der, valid_inner, valid_end)
    tag, _, t_inner, t_end = times[1]
    text = der[t_inner:t_end].decode("ascii")
    when = datetime.strptime(text, "%y%m%d%H%M%SZ" if tag == 0x17 else "%Y%m%d%H%M%SZ").replace(tzinfo=timezone.utc)
    found: dict = {"subject": (fields[4][1], fields[4][3]), "expires": when, "extensions": set()}
    for tag, _, ext_inner, _ in fields[6:]:
        if tag != 0xA3:
            continue
        _, seq_inner, seq_end = _read(der, ext_inner)
        for _, _, item_inner, item_end in _children(der, seq_inner, seq_end):
            parts = _children(der, item_inner, item_end)
            oid = _oid_text(der[parts[0][2]:parts[0][3]])
            found["extensions"].add(oid)
            if oid == SUBJECT_KEY_ID:
                # Im OCTET STRING der Erweiterung steckt die Kennung selbst als OCTET STRING.
                _, id_inner, id_end = _read(der, parts[-1][2])
                found["key_id"] = der[id_inner:id_end]
    return found


def expires(cert: Path) -> datetime:
    """Bis wann das Zertifikat gilt, ohne openssl."""
    return _parts(unpem(cert.read_text(encoding="ascii")))["expires"]


def linked(cert: Path) -> bool:
    """Nennt das Zertifikat seine Zertifizierungsstelle (Authority Key Identifier)?"""
    return AUTHORITY_KEY_ID in _parts(unpem(cert.read_text(encoding="ascii")))["extensions"]


def der(cert: Path) -> bytes:
    return unpem(cert.read_text(encoding="ascii"))


def thumbprint(cert: Path) -> str:
    """Der Fingerabdruck (SHA-1), unter dem Windows das Zertifikat führt."""
    return hashlib.sha1(der(cert)).hexdigest().upper()
