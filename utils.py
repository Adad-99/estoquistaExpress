"""
Utilitários para detecção de ambiente e caminhos de recursos.
"""
import sys
import os
from pathlib import Path


def is_frozen() -> bool:
    """
    Verifica se a aplicação está rodando como executável (PyInstaller).
    
    Returns:
        True se for EXE, False se for Python
    """
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def get_base_path() -> Path:
    """
    Obtém o caminho base da aplicação.
    
    Returns:
        Caminho base (diretório do script ou do executável)
    """
    if is_frozen():
        # Executável PyInstaller - retorna pasta onde está o .exe
        return Path(sys.executable).parent
    else:
        # Script Python - retornar o diretório do projeto (onde está main.py)
        return Path(__file__).parent


def get_resource_path(relative_path: str) -> Path:
    """
    Obtém o caminho absoluto de um recurso.
    A pasta resources deve estar sempre ao lado do executável ou do main.py
    
    Args:
        relative_path: Caminho relativo do recurso (ex: 'resources/styles.qss')
        
    Returns:
        Caminho absoluto do recurso
    """
    # Sempre busca na pasta base (ao lado do .exe ou main.py)
    base_path = get_base_path()
    return base_path / relative_path
