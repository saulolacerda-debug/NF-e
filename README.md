# NF-e Reader

Leitor de **NF-e (Nota Fiscal Eletrônica)** com exportação para **XLSX** e **CSV**, escrito em Python.

## Funcionalidades

- Leitura de arquivos XML de NF-e (versão 4.00)
- Extração de dados: identificação, emitente, destinatário, itens, totais e protocolo
- Exportação para **XLSX** com abas separadas por seção (uma única NF-e)
- Exportação em **lote** para XLSX (uma linha por NF-e, arquivo resumo)
- Exportação para **CSV**
- Processamento de diretórios inteiros de XMLs
- Interface de linha de comando (CLI)

## Instalação

```bash
pip install -r requirements.txt
pip install -e .
```

## Uso via CLI

### Exportar uma única NF-e para XLSX (com abas)

```bash
nfe-reader nota.xml -o nota.xlsx
```

### Exportar múltiplas NF-e para um XLSX resumo

```bash
nfe-reader nota1.xml nota2.xml nota3.xml --batch -o resumo.xlsx
```

### Exportar um diretório de XMLs para XLSX

```bash
nfe-reader --dir /caminho/para/xmls -o resumo.xlsx
```

### Exportar para CSV

```bash
nfe-reader nota1.xml nota2.xml --format csv -o resumo.csv
```

## Uso como biblioteca Python

```python
from nfe_reader import NFEParser, NFEExporter

# Ler uma NF-e
parser = NFEParser()
data = parser.parse_file("nota.xml")

print(data.emitente.nome)          # EMPRESA EMISSORA LTDA
print(data.identificacao.numero)   # 12345
print(data.totais.valor_nota)      # 1100.00

# Ver todos os itens
for item in data.itens:
    print(f"{item.numero}. {item.descricao} - R$ {item.valor_total}")

# Exportar para XLSX
exporter = NFEExporter()
exporter.export_xlsx(data, "saida.xlsx")

# Exportar lote para CSV
records = parser.parse_directory("/caminho/para/xmls")
exporter.export_csv(records, "resumo.csv")
```

## Estrutura do Projeto

```
nfe_reader/
├── __init__.py       # Pacote principal
├── parser.py         # Parser de XML NF-e
├── exporter.py       # Exportador XLSX/CSV
└── cli.py            # Interface de linha de comando
tests/
├── fixtures/
│   └── sample_nfe.xml    # NF-e de exemplo para testes
├── test_parser.py         # Testes do parser
└── test_exporter.py       # Testes do exportador
```

## Dados Extraídos

| Seção | Campos |
|-------|--------|
| Identificação | Chave, número, série, modelo, data de emissão, natureza da operação, UF, finalidade |
| Emitente | CNPJ/CPF, nome, fantasia, IE, CRT, endereço completo |
| Destinatário | CNPJ/CPF, nome, IE, e-mail, endereço completo |
| Itens | Código, descrição, NCM, CFOP, unidade, quantidade, valor unitário, valor total, ICMS, IPI, PIS, COFINS |
| Totais | Base ICMS, valor ICMS, valor ST, valor produtos, frete, IPI, PIS, COFINS, valor da nota |
| Protocolo | Número, data de autorização, código de status, motivo |

## Executar os Testes

```bash
pip install pytest
pytest
```

## Bibliotecas Utilizadas

- [`openpyxl`](https://openpyxl.readthedocs.io/) – geração de arquivos XLSX
- `xml.etree.ElementTree` (stdlib) – leitura de XML
- `csv` (stdlib) – exportação CSV
- `argparse` (stdlib) – interface de linha de comando
