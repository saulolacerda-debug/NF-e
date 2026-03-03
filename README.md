# NF-e

Leitor de NF-e (Nota Fiscal Eletrônica) com exportação para xlsx.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```python
from nfe_reader import NFeReader

reader = NFeReader("nota_fiscal.xml")

# Dados da NF-e
print(reader.get_identificacao())
print(reader.get_emitente())
print(reader.get_destinatario())
print(reader.get_itens())
print(reader.get_totais())

# Exportar para xlsx
reader.exportar_xlsx("nota_fiscal.xlsx")
```

## Dados extraídos

| Método | Descrição |
|---|---|
| `get_identificacao()` | Número, série, data de emissão, natureza da operação, chave de acesso |
| `get_emitente()` | CNPJ/CPF, nome, endereço do emitente |
| `get_destinatario()` | CNPJ/CPF, nome, endereço do destinatário |
| `get_itens()` | Código, descrição, NCM, CFOP, quantidade, valores, impostos (ICMS, IPI, PIS, COFINS) |
| `get_totais()` | Totais de produtos, frete, desconto, impostos e valor da NF-e |
| `exportar_xlsx(path)` | Exporta todos os dados para um arquivo xlsx com abas separadas |

## Testes

```bash
python -m pytest tests/
```
