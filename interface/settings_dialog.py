"""
Diálogo de configurações para definir a planilha padrão.
"""

import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt
from service.config_service import ConfigService
from service.requisicao_service import PedidoService


class SettingsDialog(QDialog):
    """Diálogo para configurar a planilha padrão."""
    
    def __init__(self, config_service: ConfigService, parent=None):
        super().__init__(parent)
        self.config_service = config_service
        self.init_ui()
        self.carregar_configuracoes()
    
    def init_ui(self):
        """Inicializa a interface do diálogo."""
        self.setWindowTitle("Configurações")
        self.setFixedSize(650, 200)
        self.setModal(True)
        
        # Layout principal
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Título
        titulo = QLabel("Planilha Modelo")
        titulo.setObjectName("dialogTitle")
        layout.addWidget(titulo)
        
        # Descrição
        descricao = QLabel("Selecione o arquivo Excel que será usado como modelo para criar as Pedido:")
        descricao.setWordWrap(True)
        descricao.setStyleSheet("color: #666666; font-size: 12px;")
        layout.addWidget(descricao)
        
        # Planilha padrão
        planilha_layout = QHBoxLayout()
        planilha_layout.setSpacing(10)
        
        self.planilha_input = QLineEdit()
        self.planilha_input.setObjectName("inputField")
        self.planilha_input.setReadOnly(True)
        self.planilha_input.setPlaceholderText("Nenhuma planilha selecionada...")
        self.planilha_input.setMinimumHeight(40)
        
        self.btn_selecionar = QPushButton("Selecionar")
        self.btn_selecionar.setObjectName("selectButton")
        self.btn_selecionar.setFixedWidth(120)
        self.btn_selecionar.setMinimumHeight(40)
        self.btn_selecionar.clicked.connect(self.selecionar_planilha)
        
        planilha_layout.addWidget(self.planilha_input)
        planilha_layout.addWidget(self.btn_selecionar)
        
        layout.addLayout(planilha_layout)
        
        # Botões de ação
        botoes_layout = QHBoxLayout()
        botoes_layout.addStretch()
        
        self.btn_cancelar = QPushButton("Cancelar")
        self.btn_cancelar.setFixedWidth(110)
        self.btn_cancelar.setMinimumHeight(38)
        self.btn_cancelar.clicked.connect(self.reject)
        
        self.btn_salvar = QPushButton("Salvar")
        self.btn_salvar.setObjectName("primaryButton")
        self.btn_salvar.setFixedWidth(110)
        self.btn_salvar.setMinimumHeight(38)
        self.btn_salvar.clicked.connect(self.salvar)
        
        botoes_layout.addWidget(self.btn_cancelar)
        botoes_layout.addWidget(self.btn_salvar)
        
        layout.addLayout(botoes_layout)
        
        self.setLayout(layout)
    
    def carregar_configuracoes(self):
        """Carrega as configurações atuais."""
        planilha = self.config_service.obter_planilha_padrao()
        if planilha:
            self.planilha_input.setText(planilha)
    
    def selecionar_planilha(self):
        """Abre diálogo para selecionar planilha padrão."""
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Planilha Padrão",
            "",
            "Arquivos Excel (*.xlsx *.xls)"
        )
        
        if arquivo:
            # Validar planilha
            valido, mensagem = PedidoService.validar_planilha_padrao(arquivo)
            
            if valido:
                self.planilha_input.setText(arquivo)
            else:
                QMessageBox.warning(
                    self,
                    "Planilha Inválida",
                    mensagem
                )
    
    def salvar(self):
        """Salva as configurações."""
        planilha = self.planilha_input.text()
        
        if not planilha:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, selecione uma planilha padrão."
            )
            return
        
        if not os.path.exists(planilha):
            QMessageBox.critical(
                self,
                "Erro",
                "A planilha selecionada não existe mais."
            )
            return
        
        # Salvar configuração
        if self.config_service.definir_planilha_padrao(planilha):
            QMessageBox.information(
                self,
                "Sucesso",
                "Configurações salvas com sucesso!"
            )
            self.accept()
        else:
            QMessageBox.critical(
                self,
                "Erro",
                "Erro ao salvar configurações."
            )
