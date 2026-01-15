import os
import base64
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.fernet import Fernet

DATA_DIR = os.getenv("DATA_DIR", "./data")
SALT_PATH = os.path.join(DATA_DIR, "enc_salt.bin")

def _get_or_create_salt() -> bytes:
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(SALT_PATH):
        with open(SALT_PATH, "rb") as f:
            return f.read()
    salt = os.urandom(16)
    with open(SALT_PATH, "wb") as f:
        f.write(salt)
    return salt

def derive_fernet_key(passphrase: str) -> bytes:
    if not passphrase or len(passphrase) < 8:
        raise ValueError("Passphrase must be at least 8 characters.")
    salt = _get_or_create_salt()
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=200_000,
    )
    return base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))

def encrypt_text(plaintext: str, passphrase: str) -> bytes:
    key = derive_fernet_key(passphrase)
    return Fernet(key).encrypt(plaintext.encode("utf-8"))

def decrypt_text(ciphertext: bytes, passphrase: str) -> str:
    key = derive_fernet_key(passphrase)
    return Fernet(key).decrypt(ciphertext).decode("utf-8")
