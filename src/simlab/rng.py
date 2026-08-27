from __future__ import annotations

import hashlib


def derive_seed(base_seed: int, namespace: str) -> int:
    """Derive a stable 64-bit seed for an independent random stream."""

    payload = f"{base_seed}\0{namespace}".encode()
    digest = hashlib.blake2b(
        payload,
        digest_size=8,
        person=b"simlab-rng-v1",
    ).digest()
    return int.from_bytes(digest, byteorder="big", signed=False)
