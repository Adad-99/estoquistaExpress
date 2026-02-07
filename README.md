# Sistema de Pedido de Almoxarifado

Sistema desktop para automação de processos de requisição de almoxarifado, desenvolvido em Python com interface gráfica moderna usando PySide6.

## 📋 Descrição

Aplicação desktop que facilita o gerenciamento e processamento de requisições de almoxarifado, com interface gráfica intuitiva e funcionalidades de processamento de planilhas Excel.

## 🚀 Funcionalidades

- Interface gráfica moderna e responsiva
- Processamento de planilhas Excel
- Sistema de requisições automatizado
- Suporte a High DPI
- Estilos personalizados com QSS

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **PySide6** - Framework para interface gráfica (Qt for Python)
- **openpyxl** - Manipulação de arquivos Excel

## 📦 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Passos para instalação

1. Clone o repositório:
```bash
git clone <url-do-repositorio>
cd "Automação aumoxarifado"
```

2. Crie um ambiente virtual (recomendado):
```bash
python -m venv venv
```

3. Ative o ambiente virtual:
   - **Windows:**
     ```bash
     venv\Scripts\activate
     ```
   - **Linux/Mac:**
     ```bash
     source venv/bin/activate
     ```

4. Instale as dependências:
```bash
pip install -r requirements.txt
```

## ▶️ Como Usar

Execute o arquivo principal:
```bash
python main.py
```

## 📁 Estrutura do Projeto

```
Automação aumoxarifado/
├── interface/          # Módulos da interface gráfica
├── service/           # Lógica de negócio e serviços
├── resources/         # Recursos (estilos, ícones, etc.)
├── main.py           # Ponto de entrada da aplicação
├── utils.py          # Funções utilitárias
├── requirements.txt  # Dependências do projeto
├── setup.iss         # Script de instalação (Inno Setup)
└── padrao.xlsx       # Arquivo de template Excel
```

## 🔧 Desenvolvimento

### Dependências

- `PySide6==6.6.1` - Framework Qt para Python
- `openpyxl==3.1.2` - Biblioteca para manipulação de arquivos Excel

### Gerando Executável

O projeto inclui configuração para geração de instalador Windows usando Inno Setup (`setup.iss`).

## 📝 Licença

[Especifique a licença do projeto]

## 👤 Autor

[Seu nome/organização]

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.

## 📧 Contato

[Suas informações de contato]
# estoquistaExpress
