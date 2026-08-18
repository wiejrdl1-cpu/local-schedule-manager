from __future__ import annotations

import base64
import ctypes
import os
from ctypes import wintypes

from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


class InvalidPasswordError(ValueError):
    pass


class DataProtector:
    """Encrypts sensitive values with a key derived from the app password."""

    VERIFICATION_TEXT = "업무기한관리-password-check"

    def __init__(self, password: str, salt_b64: str):
        salt = base64.urlsafe_b64decode(salt_b64.encode("ascii"))
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=600_000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode("utf-8")))
        self._fernet = Fernet(key)

    @staticmethod
    def new_salt() -> str:
        return base64.urlsafe_b64encode(os.urandom(16)).decode("ascii")

    def encrypt(self, value: str | None) -> str:
        if not value:
            return ""
        return self._fernet.encrypt(value.encode("utf-8")).decode("ascii")

    def decrypt(self, value: str | None) -> str:
        if not value:
            return ""
        try:
            return self._fernet.decrypt(value.encode("ascii")).decode("utf-8")
        except (InvalidToken, ValueError) as exc:
            raise InvalidPasswordError("비밀번호가 올바르지 않거나 데이터가 손상되었습니다.") from exc

    def make_verifier(self) -> str:
        return self.encrypt(self.VERIFICATION_TEXT)

    def verify(self, verifier: str) -> None:
        if self.decrypt(verifier) != self.VERIFICATION_TEXT:
            raise InvalidPasswordError("비밀번호가 올바르지 않습니다.")


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob_from_bytes(value: bytes) -> tuple[_DataBlob, ctypes.Array]:
    buffer = ctypes.create_string_buffer(value)
    blob = _DataBlob(len(value), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def protect_for_windows_user(value: str) -> str:
    """Protect a secret so only the current Windows user can recover it."""
    raw = value.encode("utf-8")
    input_blob, input_buffer = _blob_from_bytes(raw)
    output_blob = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(input_blob), None, None, None, None, 0, ctypes.byref(output_blob)
    ):
        raise OSError("Windows 보안 저장소에 자동 잠금 해제 정보를 저장하지 못했습니다.")
    try:
        protected = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        return base64.b64encode(protected).decode("ascii")
    finally:
        ctypes.windll.kernel32.LocalFree(output_blob.pbData)
        del input_buffer


def unprotect_for_windows_user(value: str) -> str:
    protected = base64.b64decode(value.encode("ascii"))
    input_blob, input_buffer = _blob_from_bytes(protected)
    output_blob = _DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(input_blob), None, None, None, None, 0, ctypes.byref(output_blob)
    ):
        raise OSError("이 Windows 사용자 계정에서는 자동 잠금 해제 정보를 읽을 수 없습니다.")
    try:
        raw = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        return raw.decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(output_blob.pbData)
        del input_buffer
