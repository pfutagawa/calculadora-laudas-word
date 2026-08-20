# Calculadora de Laudas e Orçamentos

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Plataforma](https://img.shields.io/badge/plataforma-Windows-0078D4?logo=windows&logoColor=white)
![Interface](https://img.shields.io/badge/interface-Tkinter-2C3E50)
![Status](https://img.shields.io/badge/status-funcional-2E8B57)

Aplicativo desktop desenvolvido em Python para automatizar a contagem de caracteres em documentos do Microsoft Word, calcular a quantidade de laudas e gerar uma planilha de orçamento.

O programa processa vários arquivos de uma pasta, inclusive em subpastas, e apresenta os resultados em uma interface gráfica responsiva. A contagem é feita pelo próprio Microsoft Word, preservando o mesmo critério usado pelo editor para caracteres com espaços.

## Demonstração

<p align="center">
  <img src="docs/interface-placeholder.svg" width="900" alt="Espaço reservado para captura da interface do programa">
</p>

> Para adicionar a captura real, salve-a como `docs/interface.png` e substitua, acima, `docs/interface-placeholder.svg` por `docs/interface.png`.

## Problema abordado

Em trabalhos de tradução, revisão e preparação editorial, o orçamento frequentemente depende do volume de texto. Verificar manualmente vários documentos, copiar a contagem de caracteres e calcular o valor de cada arquivo é um processo repetitivo e sujeito a erros.

Este projeto transforma esse fluxo em uma operação única:

1. o usuário seleciona a pasta dos documentos;
2. o programa localiza os arquivos `.doc` e `.docx`;
3. o Microsoft Word calcula os caracteres com espaços;
4. a aplicação converte o total em laudas e valores;
5. os resultados são exportados para uma planilha Excel.

## Funcionalidades

- Seleção gráfica da pasta de trabalho.
- Busca opcional em todas as subpastas.
- Processamento em lote de arquivos `.doc` e `.docx`.
- Exclusão automática de arquivos temporários do Word (`~$`).
- Valor por lauda configurável.
- Quantidade de caracteres por lauda configurável.
- Indicadores de documentos, caracteres, laudas e valor total.
- Tabela detalhada com resultado individual por arquivo.
- Barra de progresso e cancelamento do processamento.
- Tratamento individual de erros, sem interromper todo o lote.
- Exportação de planilha Excel com cabeçalhos, filtros e resumo financeiro.

## Tecnologias utilizadas

| Tecnologia | Aplicação no projeto |
|---|---|
| Python | Lógica, automação e organização da aplicação |
| Tkinter/ttk | Interface gráfica desktop |
| pywin32 | Comunicação COM com o Microsoft Word |
| openpyxl | Criação e formatação da planilha Excel |
| threading + queue | Processamento sem bloquear a interface |

## Decisões de projeto

### Interface desktop

Tkinter foi escolhido porque o fluxo depende de recursos locais do Windows: seleção de pastas e automação do Microsoft Word. Nesse contexto, uma aplicação desktop elimina a necessidade de manter um servidor web local e oferece uma experiência mais direta ao usuário.

### Processamento em segundo plano

A automação do Word pode levar alguns segundos por arquivo. Para impedir que a janela pareça travada, o trabalho é executado em uma thread separada. Uma fila envia os resultados para a interface principal de forma segura.

### Tolerância a falhas

Cada documento é processado de maneira independente. Se um arquivo estiver corrompido, protegido ou inacessível, o erro é registrado na tabela e os documentos seguintes continuam sendo analisados. O Word e os documentos também são fechados em blocos de limpeza para reduzir processos residuais.

## Como executar

### Pré-requisitos

- Windows 10 ou 11;
- Python 3.10 ou superior;
- Microsoft Word instalado.

### Instalação

Clone o repositório e entre na pasta:

```bash
git clone https://github.com/pfutagawa/calculadora-laudas-word.git
cd calculadora-laudas-word
```

Crie e ative um ambiente virtual:

```powershell
py -m venv .venv
.venv\Scripts\activate
```

Instale as dependências:

```powershell
py -m pip install -r requirements.txt
```

Execute o aplicativo:

```powershell
py app.py
```

Depois da instalação, também é possível abrir o programa pelo arquivo `executar.bat`.

## Como usar

1. Clique em **Selecionar pasta…**.
2. Informe o valor cobrado por lauda.
3. Confirme a quantidade de caracteres que representa uma lauda.
4. Escolha se deseja incluir subpastas.
5. Clique em **Processar documentos**.
6. Confira os resultados e selecione **Salvar planilha Excel…**.

## Estrutura do projeto

```text
calculadora-laudas-word/
├── app.py
├── executar.bat
├── requirements.txt
├── requirements-dev.txt
├── README.md
├── .gitignore
├── .github/workflows/tests.yml
├── docs/
│   └── interface-placeholder.svg
└── tests/
    └── test_formatacao.py
```

## Testes

Os testes automatizados verificam a interpretação de valores monetários e a formatação brasileira usada na interface.

```powershell
py -m pip install -r requirements-dev.txt
py -m unittest discover -s tests -v
```

O workflow do GitHub Actions executa os testes automaticamente no Windows a cada `push` ou `pull request`.

## Competências demonstradas

- Automação de processos operacionais repetitivos.
- Integração entre Python, Word e Excel.
- Desenvolvimento de interface orientada ao usuário.
- Processamento concorrente e comunicação segura entre threads.
- Validação de entradas e tratamento de exceções.
- Organização, documentação e testes de um projeto versionado.

## Limitações atuais

- A contagem exata depende do Microsoft Word e, portanto, o processamento funciona somente no Windows.
- Arquivos protegidos por senha não podem ser processados automaticamente.
- O cancelamento ocorre depois que o documento atual termina de ser analisado.

## Possíveis evoluções

- Geração de executável com PyInstaller.
- Histórico local de orçamentos.
- Perfis com diferentes clientes e valores por lauda.
- Exportação adicional para PDF.
- Testes de integração com uma instância controlada do Word.
