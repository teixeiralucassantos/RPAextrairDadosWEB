# Web Data Exporter

## Descrição do Projeto

O **Web Data Exporter** é uma aplicação desenvolvida em Python que utiliza uma interface gráfica para extrair dados de uma tabela disponível em um site e exportá-los para um arquivo Excel. Este projeto foi construído com a intenção de demonstrar como integrar diferentes tecnologias para criar uma solução automatizada de extração de dados da web, um conceito central em RPA (Robotic Process Automation).

## Lógica de Funcionamento

1. **Interface Gráfica**: O programa utiliza a biblioteca Tkinter para criar uma interface amigável, onde os usuários podem visualizar os dados extraídos em uma tabela organizada.
2. **Extração de Dados**: A aplicação utiliza Selenium para automação de navegadores, permitindo a interação com o site alvo. O Selenium navega até a página da tabela, aguarda o carregamento e coleta os dados de cada linha da tabela.
3. **Processamento de Dados**: Os dados coletados são inseridos em uma `Treeview`, que permite a visualização dos dados em tempo real na interface gráfica.
4. **Exportação para Excel**: Ao clicar no botão de exportação, os dados são salvos em um arquivo Excel usando a biblioteca OpenPyXL. A estrutura da planilha é definida, e os dados são inseridos linha por linha.

## Tecnologias Envolvidas

- **Python**: Linguagem de programação utilizada para o desenvolvimento do projeto.
- **Tkinter**: Biblioteca padrão do Python para criação de interfaces gráficas.
- **Selenium**: Biblioteca para automação de navegadores, utilizada para interagir com o site e coletar dados.
- **OpenPyXL**: Biblioteca para manipulação de arquivos Excel, utilizada para salvar os dados extraídos em um formato legível.

## RPA (Automação de Processos Robóticos)

RPA refere-se à tecnologia que permite a automação de tarefas repetitivas e baseadas em regras normalmente realizadas por humanos. Com RPA, software "robôs" podem ser programados para executar uma variedade de tarefas, como:

- Navegação em sites
- Extração de dados
- Preenchimento de formulários
- E muito mais

Este projeto é um exemplo prático de como a automação pode ser utilizada para melhorar a eficiência e a precisão na coleta de dados, reduzindo a necessidade de intervenção manual.

## Conceitos Técnicos

- **Web Scraping**: O processo de extração de dados de websites. É uma técnica comum utilizada para coletar informações de forma automatizada.
- **API**: (Interface de Programação de Aplicações) - Em alguns casos, a coleta de dados pode ser feita através de APIs que fornecem acesso direto aos dados em formatos estruturados, como JSON ou XML. Neste projeto, a extração foi feita diretamente da interface web.
- **Excel File Manipulation**: O uso de bibliotecas como OpenPyXL permite a criação e edição de arquivos Excel programaticamente, possibilitando a organização de dados de maneira eficiente.

## Conclusão

O **Web Data Exporter** exemplifica uma aplicação simples, mas poderosa, de automação que pode ser expandida e adaptada para atender a várias necessidades de coleta de dados. Este projeto é uma demonstração prática do potencial da RPA e da automação de tarefas no dia a dia de profissionais de diversas áreas.
