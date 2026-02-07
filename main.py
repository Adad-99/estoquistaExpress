"""
Sistema de Pedido de Almoxarifado
Ponto de entrada da aplicação.
"""

import sys
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from interface.main_window import MainWindow
from utils import get_resource_path


def carregar_estilos(app: QApplication) -> None:
    """Carrega os estilos QSS da aplicação."""
    styles_path = get_resource_path('resources/styles.qss')    
    if styles_path.exists():
        with open(styles_path, 'r', encoding='utf-8') as f:
            stylesheet = f.read()
            app.setStyleSheet(stylesheet)


def main():
    """Função principal da aplicação."""
    # Habilitar High DPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    # Criar aplicação
    app = QApplication(sys.argv)
    app.setApplicationName("Sistema de Pedido")
    app.setOrganizationName("Almoxarifado")
    
    # Carregar estilos
    carregar_estilos(app)
    
    # Criar e exibir janela principal
    window = MainWindow()
    window.show()
    
    # Executar aplicação
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
