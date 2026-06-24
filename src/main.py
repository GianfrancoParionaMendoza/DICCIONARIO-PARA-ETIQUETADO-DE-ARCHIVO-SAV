# main.py
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PySide6.QtWidgets import QApplication, QMessageBox
from src.ui.login import LoginScreen
from src.ui.home import Home
from src.core.generador_diccionario import GeneradorDiccionario
from src.core.etiquetador import Etiquetador
from datetime import datetime
import shutil
# ── Rutas fijas en disco D ────────────────────────────────────────────────
RUTA_DICCIONARIO_INSUMO  = Path("D:/ENCAL2026SYS/Diccionario/Insumo")
RUTA_DICCIONARIO_SALIDA  = Path("D:/ENCAL2026SYS/Diccionario/Resultado")

RUTA_ETIQUETADOR_EXCEL_DIR   = Path("D:/ENCAL2026SYS/Etiquetador/Insumo")
RUTA_ETIQUETADOR_INSUMO  = Path("D:/ENCAL2026SYS/Etiquetador/Insumo")
RUTA_ETIQUETADOR_SALIDA  = Path("D:/ENCAL2026SYS/Etiquetador/Resultado")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    home = Home()
    home.show()
    #login = LoginScreen()
    # def on_login():
    #     login.hide()
    #     home.show()

    # login.login_successful.connect(on_login)
    # login.show()

    generador:   GeneradorDiccionario | None = None
    etiquetador: Etiquetador          | None = None

    def execute_generador():
        nonlocal generador

        # Crear carpetas si no existen
        RUTA_DICCIONARIO_INSUMO.mkdir(parents=True, exist_ok=True)
        RUTA_DICCIONARIO_SALIDA.mkdir(parents=True, exist_ok=True)

        sav_files = list(RUTA_DICCIONARIO_INSUMO.glob("*.sav")) + \
                    list(RUTA_DICCIONARIO_INSUMO.glob("*.SAV"))
        if not sav_files:
            QMessageBox.warning(
                home, "Sin archivos",
                f"La carpeta fue creada pero no tiene archivos .sav:\n{RUTA_DICCIONARIO_INSUMO}"
            )
            return
        
        timestamp = datetime.now().strftime("%y%m%d_%H%M")
        lote_dir = RUTA_DICCIONARIO_SALIDA/ timestamp
        lote_dir.mkdir(parents=True, exist_ok=True)
        primer_sav = sav_files[0]
        partes = primer_sav.stem.split("_")

        codigo = partes[1] if len(partes) > 1 else primer_sav.stem
        salida = lote_dir/ f"Diccionario_{codigo}.xlsx"

        if generador is None:
            generador = GeneradorDiccionario(home)
        log = lote_dir/ f"insumo"
        shutil.copytree(str(RUTA_DICCIONARIO_INSUMO), str(log))

        generador.ejecutar(
            carpeta=str(RUTA_DICCIONARIO_INSUMO),
            salida=str(salida),
        )

    def execute_etiquetador():
        nonlocal etiquetador

        # Crear carpetas si no existen
        RUTA_ETIQUETADOR_EXCEL_DIR.mkdir(parents=True, exist_ok=True)
        RUTA_ETIQUETADOR_INSUMO.mkdir(parents=True, exist_ok=True)
        RUTA_ETIQUETADOR_SALIDA.mkdir(parents=True, exist_ok=True)

        # Buscar el primer .xlsx en la carpeta de etiquetas
        archivos_xlsx = list(RUTA_ETIQUETADOR_EXCEL_DIR.glob("*.xlsx"))
        if not archivos_xlsx:
            QMessageBox.critical(
                home, "Excel no encontrado",
                f"No se encontró ningún archivo .xlsx en:\n{RUTA_ETIQUETADOR_EXCEL_DIR}"
            )
            return
        xlsx_file = archivos_xlsx[0]  # toma el primero que encuentre

        # Buscar SAV
        sav_files = list(RUTA_ETIQUETADOR_INSUMO.glob("*.sav")) + \
                    list(RUTA_ETIQUETADOR_INSUMO.glob("*.SAV"))
        if not sav_files:
            QMessageBox.warning(
                home, "Sin archivos",
                f"La carpeta no tiene archivos .sav:\n{RUTA_ETIQUETADOR_INSUMO}"
            )
            return

        if etiquetador is None:
            etiquetador = Etiquetador(home)

        etiquetador.ejecutar(
            excel_path=str(xlsx_file),
            carpeta_sav=str(RUTA_ETIQUETADOR_INSUMO),
            carpeta_salida=str(RUTA_ETIQUETADOR_SALIDA),
        )

    home.card_generador_diccionario.clicked.connect(execute_generador)
    home.card_etiquetador.clicked.connect(execute_etiquetador)


    sys.exit(app.exec())


if __name__ == "__main__":
    main()