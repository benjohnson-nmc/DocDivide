"""
embed_key.py  --  Run this ONCE to generate obfuscated key literals for docdivide.py

Usage:
    python embed_key.py

Paste the two output lines (_SALT and _ENCODED_KEY) into docdivide.py,
then rebuild with:  pyinstaller docdivide.spec
"""

import os
import getpass


def encode_key(key: str) -> tuple[bytes, bytes]:
    key_bytes = key.encode()
    salt = os.urandom(32)
    salt_extended = (salt * (len(key_bytes) // len(salt) + 1))[:len(key_bytes)]
    encoded = bytes(a ^ b for a, b in zip(key_bytes, salt_extended))
    return salt, encoded


def main():
    print("DocDivide -- API Key Embedder")
    print("=" * 40)
    key = getpass.getpass("Enter your Anthropic API key (sk-ant-...): ").strip()
    if not key.startswith("sk-ant-"):
        print("Warning: key does not look like an Anthropic key (expected sk-ant-...)")

    salt, encoded = encode_key(key)

    print("\nCopy these two lines into docdivide.py (replace the existing _SALT and _ENCODED_KEY lines):\n")
    print(f"_SALT = {salt!r}")
    print(f"_ENCODED_KEY = {encoded!r}")
    print("\nDone. Rebuild with:  pyinstaller docdivide.spec")


if __name__ == "__main__":
    main()
