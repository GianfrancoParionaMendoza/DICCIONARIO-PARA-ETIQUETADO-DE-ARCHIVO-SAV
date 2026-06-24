from __future__ import annotations

import hashlib

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from datetime import date

def _h(text: str) -> str:
    return hashlib.sha256(text.encode()).hexdigest()
class UserRecord:
    def __init__(self, password_hash: str, role: str, expires: date):
        self.password_hash = password_hash
        self.role = role
        self.expires = expires

    @property
    def is_expired(self) -> bool:
        return date.today() > self.expires
_CREDENTIALS: dict[str, UserRecord] = {
    _h("@dminGian"): UserRecord(
        password_hash=_h("boby$924"),
        role="administrador",
        expires=date(2026, 12, 31),
    ),
    _h("analista01"): UserRecord(
        password_hash=_h("pass1234"),
        role="analista",
        expires=date(2026, 12, 31),
    ),
}


class LoginScreen(QMainWindow):
    login_successful = Signal(str)   # emite el rol del usuario

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Acceso — Sistema de SAV")
        self.setFixedSize(380, 340)
        self._build_ui()

    def _build_ui(self):
        root = QWidget()
        root.setStyleSheet("background-color: #f4f6fb;")
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(48, 40, 48, 40)
        layout.setSpacing(12)
        layout.setAlignment(Qt.AlignTop)

        lbl_title = QLabel("Sistema de ENCAL2026")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet(
            "font-size: 16px; font-weight: bold; color: #212529; margin-bottom: 8px;"
        )

        lbl_user = QLabel("Usuario")
        lbl_user.setStyleSheet("font-size: 12px; color: #495057;")
        self._txt_user = QLineEdit()
        self._txt_user.setPlaceholderText("Ingresa tu usuario")
        self._txt_user.setStyleSheet(self._input_style())

        lbl_pass = QLabel("Contraseña")
        lbl_pass.setStyleSheet("font-size: 12px; color: #495057;")
        self._txt_pass = QLineEdit()
        self._txt_pass.setPlaceholderText("Ingresa tu contraseña")
        self._txt_pass.setEchoMode(QLineEdit.Password)
        self._txt_pass.setStyleSheet(self._input_style())
        self._txt_pass.returnPressed.connect(self._login)

        self._lbl_error = QLabel("")
        self._lbl_error.setAlignment(Qt.AlignCenter)
        self._lbl_error.setWordWrap(True)
        self._lbl_error.setStyleSheet("color: #dc3545; font-size: 11px;")

        btn = QPushButton("Ingresar")
        btn.setStyleSheet("""
            QPushButton {
                background-color: #4361ee;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #3451d1; }
            QPushButton:pressed { background-color: #2a40b8; }
        """)
        btn.clicked.connect(self._login)

        layout.addWidget(lbl_title)
        layout.addSpacing(8)
        layout.addWidget(lbl_user)
        layout.addWidget(self._txt_user)
        layout.addWidget(lbl_pass)
        layout.addWidget(self._txt_pass)
        layout.addWidget(self._lbl_error)
        layout.addSpacing(4)
        layout.addWidget(btn)

    def _input_style(self) -> str:
        return """
            QLineEdit {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                padding: 8px 10px;
                font-size: 12px;
                background: white;
            }
            QLineEdit:focus { border-color: #4361ee; }
        """

    def _login(self):
        user = self._txt_user.text().strip()
        pwd  = self._txt_pass.text()

        record = _CREDENTIALS.get(_h(user))

        if record is None or record.password_hash != _h(pwd):
            self._show_error("Usuario o contraseña incorrectos.")
            return

        if record.is_expired:
            expires_str = record.expires.strftime("%d/%m/%Y")
            self._show_error(f"Cuenta expirada el {expires_str}.\nContacte al administrador.")
            return

        # Login exitoso
        self._lbl_error.setText("")
        self.login_successful.emit(record.role)

    def _show_error(self, msg: str):
        self._lbl_error.setText(msg)
        self._txt_pass.clear()
        self._txt_pass.setFocus()