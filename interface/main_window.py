"""
Janela principal do sistema de criação de Pedido.
Interface moderna e intuitiva com PySide6.
"""

import os
import sys
import subprocess
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog,
    QMessageBox, QFrame, QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QFont

from service.config_service import ConfigService
from service.requisicao_service import PedidoService
from utils import get_resource_path
from .settings_dialog import SettingsDialog


class MainWindow(QMainWindow):
    """Janela principal da aplicação."""
    
    def __init__(self):
        super().__init__()
        self.config_service = ConfigService()
        self.ultimo_arquivo_criado = None
        self.tema_escuro = self.config_service.obter_config('tema_escuro', False)
        
        # Para arrastar a janela
        self.dragging = False
        self.offset = QPoint()
        
        # Inicializar interface
        self.init_ui()
        self.aplicar_tema()
        
        # Carregar valores salvos DEPOIS de criar os widgets
        self.carregar_valores_salvos()
        
        # Configurar planilha padrão se não existir
        self.configurar_planilha_padrao_inicial()
        self.verificar_configuracao_inicial()
    
    def init_ui(self):
        """Inicializa a interface."""
        # Janela sem frame
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setFixedSize(700, 430)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Barra de título customizada
        title_bar = self.criar_barra_titulo()
        main_layout.addWidget(title_bar)
        
        # Conteúdo
        content = QWidget()
        content.setObjectName("contentArea")
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(15, 15, 15, 8)  # Margem inferior reduzida
        content_layout.setSpacing(10)
        
        # 1. Campo Setor (topo)
        setor_label = QLabel("Nome do Setor")
        setor_label.setObjectName("fieldLabel")
        setor_label.setFixedHeight(18)
        content_layout.addWidget(setor_label)
        
        self.setor_input = QLineEdit()
        self.setor_input.setObjectName("inputField")
        self.setor_input.setPlaceholderText("Digite o nome do setor...")
        self.setor_input.setFixedHeight(32)
        content_layout.addWidget(self.setor_input)
        
        # 2. Campo Pasta Destino + Botão
        pasta_label = QLabel("Pasta de Destino")
        pasta_label.setObjectName("fieldLabel")
        pasta_label.setFixedHeight(18)
        content_layout.addWidget(pasta_label)
        
        pasta_layout = QHBoxLayout()
        pasta_layout.setSpacing(10)
        
        self.pasta_input = QLineEdit()
        self.pasta_input.setObjectName("inputField")
        self.pasta_input.setPlaceholderText("Selecione a pasta de destino...")
        self.pasta_input.setFixedHeight(32)
        self.pasta_input.setReadOnly(True)
        pasta_layout.addWidget(self.pasta_input)
        
        btn_selecionar_pasta = QPushButton("Selecionar Pasta")
        btn_selecionar_pasta.setObjectName("selectButton")
        btn_selecionar_pasta.setFixedHeight(32)
        btn_selecionar_pasta.setFixedWidth(130)
        btn_selecionar_pasta.clicked.connect(self.selecionar_pasta)
        pasta_layout.addWidget(btn_selecionar_pasta)
        
        content_layout.addLayout(pasta_layout)
        
        # 3. Botões processar e Ver Arquivo
        botoes_layout = QHBoxLayout()
        botoes_layout.setSpacing(10)
        
        self.btn_processar = QPushButton("Criar Pedido")
        self.btn_processar.setObjectName("processButton")
        self.btn_processar.setFixedHeight(36)
        self.btn_processar.clicked.connect(self.criar_Pedido)
        botoes_layout.addWidget(self.btn_processar)
        
        botoes_layout.addStretch()
        
        self.btn_ver_arquivo = QPushButton("Ver Arquivo")
        self.btn_ver_arquivo.setObjectName("viewButton")
        self.btn_ver_arquivo.setFixedHeight(36)
        self.btn_ver_arquivo.setEnabled(False)
        self.btn_ver_arquivo.clicked.connect(self.ver_ultimo_arquivo)
        botoes_layout.addWidget(self.btn_ver_arquivo)
        
        content_layout.addLayout(botoes_layout)
        
        # 4. Histórico
        historico_label = QLabel("Histórico")
        historico_label.setObjectName("historicoLabel")
        historico_label.setFixedHeight(20)  # Aumentado para não cortar
        content_layout.addWidget(historico_label)
        
        self.lista_historico = QListWidget()
        self.lista_historico.setObjectName("historicoList")
        self.lista_historico.setFixedHeight(110)  # 3 itens visíveis + scroll
        self.lista_historico.itemDoubleClicked.connect(self.abrir_arquivo_historico)
        content_layout.addWidget(self.lista_historico)
        
        # Status
        self.status_label = QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)
        content_layout.addWidget(self.status_label)
        
        # Assinatura com link
        assinatura = QLabel('<a href="https://github.com/Adad-99" style="color: #888888; text-decoration: none;">Desenvolvido por Guilherme Adad</a>')
        assinatura.setObjectName("assinatura")
        assinatura.setAlignment(Qt.AlignCenter)
        assinatura.setOpenExternalLinks(True)  # Permitir abrir links externos
        assinatura.setCursor(Qt.PointingHandCursor)  # Cursor de mão
        assinatura.setTextFormat(Qt.RichText)  # Aceitar HTML
        assinatura.setContentsMargins(0, 0, 0, 0)  # SEM margem
        content_layout.addWidget(assinatura)
        
        content.setLayout(content_layout)
        main_layout.addWidget(content)
        
        central_widget.setLayout(main_layout)
        
        # Carregar histórico
        self.carregar_historico()
    
    def carregar_valores_salvos(self):
        """Carrega os últimos valores salvos nos campos."""
        # Carregar último setor
        ultimo_setor = self.config_service.obter_ultimo_setor()
        if ultimo_setor:
            self.setor_input.setText(ultimo_setor)
        
        # Carregar última pasta
        ultima_pasta = self.config_service.obter_config('ultima_pasta', '')
        if ultima_pasta:
            self.pasta_input.setText(ultima_pasta)
    
    def configurar_planilha_padrao_inicial(self):
        """Configura a planilha padrão na primeira execução."""
        
        planilha_atual = self.config_service.obter_planilha_padrao()
        
        # Se não houver planilha configurada, usar a padrão
        if not planilha_atual:
            planilha_padrao = str(get_resource_path('resources/padrao.xlsx'))
            
            # Verificar se existe
            if os.path.exists(planilha_padrao):
                self.config_service.definir_planilha_padrao(planilha_padrao)
    
    def criar_barra_titulo(self) -> QWidget:
        """Cria a barra de título customizada."""
        title_bar = QWidget()
        title_bar.setObjectName("titleBar")
        title_bar.setFixedHeight(45)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(15, 0, 10, 0)
        layout.setSpacing(10)
        
        # Título
        titulo = QLabel("Estoquista Express")
        titulo.setObjectName("titleBarLabel")
        layout.addWidget(titulo)
        
        layout.addStretch()
        
        # Botão de tema (Sol = claro, Lua = escuro)
        self.btn_tema = QPushButton("☾" if self.tema_escuro else "☀")
        self.btn_tema.setObjectName("themeButton")
        self.btn_tema.setFixedSize(35, 35)
        self.btn_tema.setToolTip("Alternar tema")
        self.btn_tema.clicked.connect(self.alternar_tema)
        layout.addWidget(self.btn_tema)
        
        # Botão de configurações
        btn_settings = QPushButton("⚙")
        btn_settings.setObjectName("settingsButton")
        btn_settings.setFixedSize(35, 35)
        btn_settings.setToolTip("Configurações")
        btn_settings.clicked.connect(self.abrir_configuracoes)
        layout.addWidget(btn_settings)
        
        # Separador
        separator = QFrame()
        separator.setObjectName("separator")
        separator.setFixedWidth(1)
        separator.setFixedHeight(25)
        layout.addWidget(separator)
        
        # Botão minimizar
        btn_minimize = QPushButton("−")
        btn_minimize.setObjectName("minimizeButton")
        btn_minimize.setFixedSize(35, 35)
        btn_minimize.clicked.connect(self.showMinimized)
        layout.addWidget(btn_minimize)
        
        # Botão fechar
        btn_close = QPushButton("×")
        btn_close.setObjectName("closeButton")
        btn_close.setFixedSize(35, 35)
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)
        
        title_bar.setLayout(layout)
        return title_bar
    
    def carregar_historico(self):
        """Carrega o histórico de Pedido criadas."""
        ultimas = self.config_service.obter_config('ultimas_requisicoes', [])
        self.lista_historico.clear()
        # Adiciona os itens em ordem inversa para que os mais recentes apareçam no topo
        for req_info in reversed(ultimas):
            # req_info agora é uma tupla (arquivo_path, pasta, setor)
            if isinstance(req_info, list) and len(req_info) == 3:
                arquivo_path, pasta, setor = req_info
                nome_arquivo = os.path.basename(arquivo_path)
                item_texto = f"{nome_arquivo} - {pasta}/{setor}"
                item = QListWidgetItem(item_texto)
                item.setData(Qt.UserRole, arquivo_path)
                self.lista_historico.addItem(item)
            else: # Compatibilidade com formato antigo
                self.lista_historico.addItem(req_info)
    
    def adicionar_historico(self, arquivo: str, pasta: str, setor: str):
        """Adiciona item ao histórico (no topo da lista)."""
        ultimas = self.config_service.obter_config('ultimas_requisicoes', [])
        
        # Novo formato: armazenar o caminho completo e informações para exibição
        item_data = [arquivo, pasta, setor]
        
        # Remover se já existir para evitar duplicatas e mover para o topo
        if item_data in ultimas:
            ultimas.remove(item_data)
        
        ultimas.append(item_data)
        
        # Manter apenas 50 últimas
        ultimas = ultimas[-50:]
        
        self.config_service.definir_config('ultimas_requisicoes', ultimas)
        self.carregar_historico() # Recarrega a lista para exibir o novo item no topo
    
    def abrir_arquivo_historico(self, item):
        """Abre arquivo do histórico ao dar duplo clique."""
        caminho = item.data(Qt.UserRole) # Obter o caminho completo do arquivo
        
        if caminho and os.path.exists(caminho):
            self.abrir_arquivo_path(caminho)
        else:
            QMessageBox.warning(self, "Aviso", "Arquivo não encontrado ou caminho inválido!")
    
    def ver_ultimo_arquivo(self):
        """Abre o último arquivo criado."""
        if self.ultimo_arquivo_criado and os.path.exists(self.ultimo_arquivo_criado):
            self.abrir_arquivo_path(self.ultimo_arquivo_criado)
        else:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo foi criado ainda!")
    
    def verificar_configuracao_inicial(self):
        """Verifica se há uma planilha padrão configurada."""
        # A lógica de configuração inicial da planilha foi movida para configurar_planilha_padrao_inicial
        # Este método pode ser removido ou adaptado se houver outras verificações iniciais.
        pass
    
    def alternar_tema(self):
        """Alterna entre tema claro e escuro."""
        self.tema_escuro = not self.tema_escuro
        self.config_service.definir_config('tema_escuro', self.tema_escuro)
        # Sol = claro, Lua = escuro
        self.btn_tema.setText("☾" if self.tema_escuro else "☀")
        self.aplicar_tema()
    
    def aplicar_tema(self):
        """Aplica o tema atual (claro ou escuro)."""
        if self.tema_escuro:
            self.setProperty("theme", "dark")
        else:
            self.setProperty("theme", "light")
        
        # Força o recarregamento do estilo
        self.style().unpolish(self)
        self.style().polish(self)
        
        # Atualiza todos os widgets filhos
        for widget in self.findChildren(QWidget):
            widget.style().unpolish(widget)
            widget.style().polish(widget)
    
    def abrir_configuracoes(self):
        """Abre o diálogo de configurações."""
        dialog = SettingsDialog(self.config_service, self)
        dialog.exec()
    
    def selecionar_pasta(self):
        """Abre diálogo para selecionar pasta de destino."""
        # Abrir na pasta Downloads por padrão
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        
        pasta = QFileDialog.getExistingDirectory(
            self,
            "Selecionar Pasta de Destino",
            downloads_path  # Caminho inicial
        )
        
        if pasta:
            self.pasta_input.setText(pasta)
            # Salvar última pasta usada
            self.config_service.definir_config('ultima_pasta', pasta)
            self.atualizar_status(
                f"Pasta selecionada: {os.path.basename(pasta)}",
                "info"
            )
    
    def criar_Pedido(self):
        """Cria uma nova Pedido."""
        # Validações
        setor = self.setor_input.text().strip()
        pasta = self.pasta_input.text()
        
        if not setor:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, informe o setor!"
            )
            return
        
        if not pasta:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, selecione a pasta de destino!"
            )
            return
        
        # Obter planilha padrão
        planilha_padrao = self.config_service.obter_planilha_padrao()
        
        if not planilha_padrao or not os.path.exists(planilha_padrao):
            QMessageBox.critical(
                self,
                "Erro",
                "Planilha padrão não configurada ou não encontrada.\n"
                "Por favor, configure nas configurações."
            )
            return
        
        # Criar Pedido
        sucesso, mensagem, arquivo = PedidoService.criar_Pedido(
            setor, pasta, planilha_padrao
        )
        
        if sucesso:
            self.ultimo_arquivo_criado = arquivo
            self.btn_ver_arquivo.setEnabled(True)
            
            # Salvar último setor digitado
            self.config_service.definir_ultimo_setor(setor)
            
            # Adicionar ao histórico (no topo)
            self.adicionar_historico(arquivo, pasta, setor)
            
            # Mostrar apenas na barra de status
            self.atualizar_status(f"✓ Pedido criado com sucesso! Arquivo: {os.path.basename(arquivo)}", "success")
        else:
            self.atualizar_status(mensagem, "error")
            QMessageBox.critical(
                self,
                "Erro",
                mensagem
            )
    
    def abrir_arquivo_path(self, caminho):
        """Abre um arquivo Excel pelo caminho."""
        try:
            if sys.platform == 'win32':
                os.startfile(caminho)
            elif sys.platform == 'darwin':  # macOS
                subprocess.call(['open', caminho])
            else:  # linux
                subprocess.call(['xdg-open', caminho])
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erro",
                f"Erro ao abrir arquivo:\n{str(e)}"
            )
    
    def atualizar_status(self, mensagem: str, tipo: str = "info"):
        """
        Atualiza a mensagem de status.
        
        Args:
            mensagem: Texto da mensagem
            tipo: Tipo da mensagem (info, success, error)
        """
        self.status_label.setText(mensagem)
        self.status_label.setProperty("statusType", tipo)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
    
    # Eventos para arrastar a janela
    def mousePressEvent(self, event):
        """Detecta clique para arrastar."""
        if event.button() == Qt.LeftButton:
            # Verifica se clicou na barra de título
            title_bar = self.findChild(QWidget, "titleBar")
            if title_bar and title_bar.geometry().contains(event.pos()):
                self.dragging = True
                self.offset = event.pos()
    
    def mouseMoveEvent(self, event):
        """Move a janela ao arrastar."""
        if self.dragging:
            self.move(self.mapToGlobal(event.pos() - self.offset))
    
    def mouseReleaseEvent(self, event):
        """Para de arrastar."""
        if event.button() == Qt.LeftButton:
            self.dragging = False
