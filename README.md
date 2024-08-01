# QtExpoTool 🚀

QtExpoTool é uma aplicação em Python desenvolvida com PyQt5, pandas e docx para importar, processar e exportar planilhas do Excel para documentos Word e PDF. Ela oferece uma interface amigável para gerenciar o processo de conversão com atualizações de progresso em tempo real.

## Propósito 🎯

O objetivo principal do QtExpoTool é lidar com grandes arquivos de planilha e, usando um modelo predefinido, exportar automaticamente esses dados para um documento Word. Isso é extremamente útil para a criação de relatórios de alta qualidade de forma rápida e eficiente. Além disso, a ferramenta possui funções de reformatação de tabelas que garantem a correta formatação de valores numéricos. A combinação das bibliotecas numpy e pandas garante rapidez e eficiência no processamento dos dados. A interface moderna e atraente é fornecida pelo framework Qt, tornando a experiência do usuário mais agradável.

## Funcionalidades ✨

- **Importar arquivos Excel (.xlsx, .xls)**
- **Exportar dados para:**
  - Arquivos Excel (.xlsx) 
  - Documentos Word (.docx) 
  - Documentos PDF (.pdf) 
- **Exibir dados em uma visualização de tabela com opções de formatação**
- **Atualizar o progresso com uma barra de progresso**

## Requisitos ⚠️

- Python 3.7+
- Pacotes Python necessários:
  - PyQt5
  - pandas
  - python-docx
  - comtypes

## Instalação 🖥️

### Clonar o Repositório

```bash
git clone https://github.com/DesignerDjalma/QtExpoTool.git
cd QtExpoTool
```

### Criar e Ativar o Ambiente Virtual

É recomendado usar um ambiente virtual para gerenciar as dependências:

```bash
python -m venv venv
source venv/bin/activate # No Windows use venv\Scripts\activate
```

### Instalar Dependências

```bash
pip install -r requirements.txt
```

Se você não possui um arquivo `requirements.txt`, pode instalar os pacotes necessários manualmente:

```bash
pip install pyqt5 pandas python-docx comtypes
```

### Executar a Aplicação

```bash
python xlsxparapdf.py
```

## Uso 📚

1. **Importar Arquivo Excel**: Clique no botão "Importar Planilha" para selecionar e carregar um arquivo Excel.
2. **Definir Local de Exportação**: Clique no botão "..." ao lado de "Local de Exportação" para selecionar a pasta onde os arquivos exportados serão salvos.
3. **Exportar Dados**:
   - Clique no botão "Exportar .xlsx" para exportar os dados como um arquivo Excel.
   - Clique no botão "Exportar .docx" para exportar os dados como um documento Word.
   - Clique no botão "Exportar .pdf" para exportar os dados como um documento PDF.
4. **Sobre**: Clique no botão "Sobre" para obter informações sobre a aplicação.

## Nota Importante 📝

O modelo de documento .docx para ser usado estará na pasta `docs`.

## Licença 📄

Este projeto está licenciado sob a Licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## Agradecimentos 🙌

- A biblioteca PyQt5 pela interface gráfica
- A biblioteca pandas pela manipulação de dados
- A biblioteca python-docx pela criação de documentos Word

