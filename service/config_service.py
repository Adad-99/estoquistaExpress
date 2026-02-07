"""
Serviço de gerenciamento de configurações.
Salva e carrega configurações do sistema em formato JSON.
"""

import os
import json
import tempfile
from pathlib import Path
from typing import Any, Optional


class ConfigService:
    """Serviço para gerenciar configurações do aplicativo."""
    
    def __init__(self):
        """Inicializa o serviço de configuração."""
        # Usar pasta temporária do Windows para persistência
        temp_dir = tempfile.gettempdir()
        self.config_dir = os.path.join(temp_dir, "PedidoAlmoxarifado")
        
        # Criar diretório se não existir
        os.makedirs(self.config_dir, exist_ok=True)
        
        self.config_file = os.path.join(self.config_dir, "config.json")
        self.config = self._carregar_config()
    
    def _carregar_config(self) -> dict:
        """Carrega as configurações do arquivo JSON."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                return {}
        return {}
    
    def _salvar_config(self) -> bool:
        """Salva as configurações no arquivo JSON."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            return False
    
    def obter_planilha_padrao(self) -> Optional[str]:
        """
        Obtém o caminho da planilha padrão.
        
        Returns:
            Caminho da planilha padrão ou None se não configurado
        """
        return self.config.get('planilha_padrao')
    
    def definir_planilha_padrao(self, caminho: str) -> bool:
        """
        Define o caminho da planilha padrão.
        
        Args:
            caminho: Caminho completo da planilha padrão
            
        Returns:
            True se salvo com sucesso, False caso contrário
        """
        if not os.path.exists(caminho):
            return False
        
        self.config['planilha_padrao'] = caminho
        return self._salvar_config()
    
    def obter_config(self, chave: str, padrao=None):
        """Obtém uma configuração específica."""
        return self.config.get(chave, padrao)
    
    def definir_config(self, chave: str, valor: Any) -> None:
        """
        Define um valor de configuração genérica.
        
        Args:
            chave: Chave da configuração
            valor: Valor a ser salvo
        """
        self.config[chave] = valor
        self._salvar_config()
    
    def obter_ultimo_setor(self) -> str:
        """
        Obtém o último setor digitado.
        
        Returns:
            Último setor ou string vazia
        """
        return self.config.get('ultimo_setor', '')
    
    def definir_ultimo_setor(self, setor: str) -> None:
        """
        Salva o último setor digitado.
        
        Args:
            setor: Nome do setor
        """
        self.config['ultimo_setor'] = setor
        self._salvar_config()
