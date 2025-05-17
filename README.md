# Gerador de Arquivos por Técnico

Este projeto é uma aplicação desktop simples desenvolvida em Python com interface gráfica usando `customtkinter`, que permite importar uma planilha Excel, selecionar uma aba e gerar arquivos `.txt` individuais para cada técnico listado.

## 📋 Funcionalidades

- Interface gráfica amigável.
- Permite selecionar uma planilha Excel (`.xlsx` ou `.xls`).
- Permite escolher a aba desejada da planilha.
- Filtra e organiza os dados com base em colunas específicas.
- Gera arquivos `.txt` separados por técnico com seus respectivos dados.
- Salva os arquivos em uma pasta chamada `Arquivos_por_Tecnico` no mesmo diretório da planilha.

## 🛠️ Tecnologias e Bibliotecas

- `pandas` – para leitura e manipulação de dados.
- `customtkinter` – para interface gráfica moderna.
- `tkinter` – para diálogos de seleção de arquivos e mensagens.
- `os` – para manipulação de arquivos e diretórios.

## 🧾 Como usar

1. Execute o script.
2. Clique em **Selecionar Planilha** e escolha um arquivo `.xlsx` ou `.xls`.
3. Selecione a aba da planilha desejada no menu suspenso.
4. Clique em **Processar Aba Selecionada**.
5. Os arquivos `.txt` serão gerados automaticamente na pasta `Arquivos_por_Tecnico`.

## 🧠 Estrutura esperada da planilha

A planilha deve conter pelo menos as seguintes colunas (com esses nomes ou variações com espaços/maiúsculas/minúsculas):

- `Novo TEC`
- `Endereço`
- `Novo End`
- `Intervalo de Tempo`
- `Contrato`

As colunas podem estar em qualquer ordem, desde que existam.

## 📁 Saída

Os arquivos `.txt` serão gerados com o nome do técnico (retirado da coluna **"Novo TEC"**), separados por tabulação (`.tsv`), e salvos na pasta `Arquivos_por_Tecnico`.

## ⚠️ Erros comuns

- A planilha não possui alguma das colunas obrigatórias.
- A aba selecionada não contém dados válidos.
- O arquivo não é um Excel válido.

