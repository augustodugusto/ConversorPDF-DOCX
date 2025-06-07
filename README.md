Conversor PDF para DOCX
Uma aplicação de desktop simples e eficiente para converter arquivos PDF em documentos .docx totalmente editáveis, focada em preservar o layout original e a estrutura do texto.

Visão Geral
Este projeto foi desenvolvido como um estudo de caso prático em engenharia de software e manipulação de dados. O objetivo era criar uma solução funcional e distribuível para um problema comum, aplicando conceitos desde a arquitetura do software até o empacotamento final para o usuário.

Principais Funcionalidades
Interface Gráfica Intuitiva: Uma janela simples e limpa que qualquer usuário pode operar sem dificuldades.

Conversão de Alta Fidelidade: Foco em manter o texto fluido e os parágrafos corretamente estruturados, ao invés de criar caixas de texto flutuantes.

Processamento em Segundo Plano: A interface não congela durante a conversão, graças ao uso de threading para tarefas demoradas.

Autocontido: O instalador inclui todas as dependências necessárias (incluindo o Pandoc). O usuário final não precisa instalar nada a mais.

Código Aberto: O código está disponível para estudo, modificação e contribuição da comunidade.

Como Funciona: A Arquitetura
O processo de conversão é dividido em duas etapas principais para garantir o melhor resultado possível:

graph LR
    A["<div style='font-weight:bold; font-size:16px;'>Arquivo PDF</div><div style='font-size:12px;'>Entrada do Usuário</div>"] -->|"Etapa 1: Extração com PyMuPDF"| B["<div style='font-weight:bold; font-size:16px;'>HTML Semântico</div><div style='font-size:12px;'>Estrutura Intermediária</div>"]
    B -->|"Etapa 2: Conversão com Pandoc"| C["<div style='font-weight:bold; font-size:16px;'>Arquivo DOCX</div><div style='font-size:12px;'>Saída Editável</div>"]

    style A fill:#D9534F,stroke:#333,stroke-width:2px,color:#fff
    style B fill:#f0f0f0,stroke:#333,stroke-width:2px,color:#333
    style C fill:#428BCA,stroke:#333,stroke-width:2px,color:#fff

Extração (PDF → HTML): A biblioteca PyMuPDF analisa o PDF e extrai seu conteúdo de forma estruturada, gerando um arquivo HTML que representa o fluxo de texto e as imagens.

Conversão (HTML → DOCX): A ferramenta Pandoc é utilizada para converter este HTML intermediário em um arquivo .docx compatível com o Microsoft Word e outros editores.

Instalação e Uso (Para Usuários)
Para usar a aplicação, basta baixar e executar o instalador.

Baixe o Instalador:

Acesse a seção de Releases deste repositório.

Baixe o arquivo setup_conversor_pdf_v1.0.exe.

Execute o Instalador:

Dê um duplo clique no arquivo baixado.

Siga as instruções na tela. O programa será instalado e um atalho será criado na sua Área de Trabalho e no Menu Iniciar.

Use o Programa:

Abra o "Conversor PDF para DOCX" pelo atalho.

Clique no botão "Selecionar PDF e Converter".

Escolha o arquivo PDF que deseja converter e aguarde. O arquivo .docx final será salvo na mesma pasta do PDF original e aberto automaticamente.

Para Desenvolvedores (Rodando do Código-Fonte)
Se você deseja executar o projeto a partir do código-fonte para estudar ou contribuir:

Pré-requisitos:

Python 3.9 ou superior.

Pandoc: É necessário ter o Pandoc instalado no seu sistema. Você pode baixá-lo em pandoc.org/installing.html.

Passos:

Clone o repositório:

git clone https://github.com/augustodugusto/ConversorPDF-DOCX.git
cd ConversorPDF-DOCX

Crie um ambiente virtual (recomendado):

python -m venv venv
# No Windows
.\venv\Scripts\activate

Instale as dependências:

pip install -r requirements.txt

Execute o script principal:

python conversor.py

Licença
Este projeto está licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.

Agradecimentos
Este projeto não seria possível sem o trabalho incrível das comunidades por trás das seguintes ferramentas:

PyMuPDF

Pandoc

PyInstaller

Inno Setup