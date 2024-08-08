# BOT XML OBOTICARIO AUTOMATICO

Este projeto automatiza a extração e o processamento de documentos eletrônicos (XML) do site VarejoFacil, utilizando Selenium WebDriver para interagir com a interface web e scripts Python para mover e descompactar os arquivos baixados.

## Requisitos

Certifique-se de ter os seguintes softwares instalados em sua máquina:

- Python 3.7+
- Google Chrome (última versão)
- ChromeDriver (o WebDriver Manager cuidará da instalação)

## Instalação

1. Clone o repositório:

   ```bash
   git clone https://github.com/EnzzoHosaki/bot-xml-oboticario
   cd repositorio

2. Crie um ambiente virtual e ative-o:
    ```bash
    python -m venv venv
    source venv/bin/activate  # No Windows, use `venv\Scripts\activate`

3. Instale as dependências:
    ```bash
    Instale as dependências:

## Instalação

1. Para executar o script principal e iniciar a automação, execute:
    ```bash
    python bot_xml_oboticario_manual.py

2. O script executará as seguintes etapas:

    2.1. Acessará o site VarejoFacil e fará login.
    2.2. Navegará até a página de exportação de documentos eletrônicos.
    2.3. Preencherá o formulário com os parâmetros apropriados e iniciará a exportação.
    2.4. Fará o download dos arquivos exportados e os moverá para os diretórios apropriados, dentro da Rede da RPS, descompactando-os conforme necessário.

## Detalhes do Script

`get_xml()`
Esta função usa Selenium para:

    - Fazer login no site.
    - Navegar até a página de exportação de documentos.
    - Preencher o formulário de exportação com os parâmetros especificados.
    - Iniciar o processo de exportação e download dos arquivos XML.

`move_xml_directory()`
Esta função:

    - Move os arquivos baixados para os diretórios apropriados.
    - Descompacta os arquivos e organiza a estrutura de diretórios.

## Dependências
As principais dependências usadas neste projeto são:

- `selenium`: Para automação do navegador.
- `webdriver-manager`: Para gerenciar a instalação do ChromeDriver.

## Arquivo `requirements.txt`
selenium==4.11.2
webdriver-manager==4.0.0

## Criação do Executável
Para criar um executável deste script, diga os passos abaixo:

1. Instale o PyInstaller:
    ```bash
    pip install pyinstaller
2. Gere o executável:
    ```bash
    pyinstaller --onefile --windowed bot_xml_oboticario_manual.py

O executável será criado na pasta `dist` dentro do diretório do projeto.

## Licença
Este projeto está licenciado sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.
