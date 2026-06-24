from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QLabel,
    QMainWindow,
    QVBoxLayout,
    QWidget,
)
from core.etiquetador import Etiquetador

class NavCard(QFrame):
    clicked = Signal()

    def __init__(self, icon: str, title: str, description: str, enabled: bool = True, parent=None):
        super().__init__(parent)
        self._enabled_card = enabled
        self._setup_ui(icon, title, description)
        self._apply_style()
        if enabled:
            self.setCursor(Qt.PointingHandCursor)

    def _setup_ui(self, icon: str, title: str, description: str):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(10)
        layout.setContentsMargins(24, 28, 24, 28)

        lbl_icon = QLabel(icon)
        lbl_icon.setAlignment(Qt.AlignCenter)
        lbl_icon.setStyleSheet("font-size: 40px;")

        lbl_title = QLabel(title)
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet(
            "font-size: 15px; font-weight: bold; color: #212529;"
            if self._enabled_card else
            "font-size: 15px; font-weight: bold; color: #adb5bd;"
        )

        lbl_desc = QLabel(description)
        lbl_desc.setAlignment(Qt.AlignCenter)
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet(
            "font-size: 11px; color: #6c757d;"
            if self._enabled_card else
            "font-size: 11px; color: #ced4da;"
        )

        layout.addWidget(lbl_icon)
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_desc)

    def _apply_style(self):
        if self._enabled_card:
            self.setStyleSheet("""
                NavCard {
                    background-color: #ffffff;
                    border: 1px solid #dee2e6;
                    border-radius: 12px;
                    min-width: 200px;
                    min-height: 170px;
                }
                NavCard:hover {
                    background-color: #f0f4ff;
                    border: 2px solid #4361ee;
                }
            """)
        else:
            self.setStyleSheet("""
                NavCard {
                    background-color: #f8f9fa;
                    border: 1px dashed #dee2e6;
                    border-radius: 12px;
                    min-width: 200px;
                    min-height: 170px;
                }
            """)

    def mousePressEvent(self, event):
        if self._enabled_card:
            self.clicked.emit()


class Home(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de SAV")
        self.resize(800, 520)
        self._build_ui()

    def _build_ui(self):
        root = QWidget()
        root.setStyleSheet("background-color: #f4f6fb;")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setAlignment(Qt.AlignTop)
        outer.setContentsMargins(40, 40, 40, 40)
        outer.setSpacing(8)

        lbl_title = QLabel("Automatización de Diccionarios y Etiquetado de Archivos SAV")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet("font-size: 22px; font-weight: bold; color: #212529;")

        lbl_subtitle = QLabel("Presione una opcion para continua \nNota: Si es la primera vez usando el sistema creara de manera automatica \n una carpeta en el particion 'D:' dentro debera introducir lo necesario \npara las operaciones que requiera")
        lbl_subtitle.setAlignment(Qt.AlignCenter)
        lbl_subtitle.setStyleSheet("font-size: 13px; color: #6c757d; margin-bottom: 20px;")

        outer.addWidget(lbl_title)
        outer.addWidget(lbl_subtitle)

        grid = QGridLayout()
        grid.setSpacing(20)
        grid.setAlignment(Qt.AlignCenter)

        modules = [
            ("📊", "Generador de Diccionarios",  "Genera de manera automatica el diccionario de datos",  True),
            ("🔍", "Etiquetador de Diccionario",  "Etiquetar archivos sav",     True),
        ]

        self._cards: list[NavCard] = []
        for i, (icon, title, desc, enabled) in enumerate(modules):
            card = NavCard(icon, title, desc, enabled)
            self._cards.append(card)
            grid.addWidget(card, i // 2, i % 2)

        outer.addLayout(grid)

    @property
    def card_generador_diccionario(self) -> NavCard:
        return self._cards[0]

    @property
    def card_etiquetador(self) -> NavCard:
        return self._cards[1]
    
