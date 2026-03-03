"""Command-line interface for nfe_reader."""

from __future__ import annotations

import argparse
import sys
from typing import List

from .exporter import NFEExporter
from .parser import NFEData, NFEParseError, NFEParser


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="nfe-reader",
        description="Leitor de NF-e (Nota Fiscal Eletrônica) com exportação para XLSX/CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  # Exportar uma única NF-e para XLSX
  nfe-reader nota.xml -o nota.xlsx

  # Exportar múltiplas NF-e para um único XLSX (resumo)
  nfe-reader *.xml --batch -o resumo.xlsx

  # Exportar uma pasta de XMLs para CSV
  nfe-reader --dir /caminho/xmls --format csv -o resumo.csv
""",
    )

    # Input
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        "arquivos",
        nargs="*",
        metavar="ARQUIVO",
        help="Um ou mais arquivos XML de NF-e.",
        default=[],
    )
    input_group.add_argument(
        "--dir",
        metavar="DIRETÓRIO",
        help="Diretório contendo arquivos XML de NF-e.",
    )

    # Output
    parser.add_argument(
        "-o",
        "--output",
        metavar="SAÍDA",
        required=True,
        help="Caminho do arquivo de saída (ex: saida.xlsx ou saida.csv).",
    )
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Formato de exportação: xlsx (padrão) ou csv.",
    )

    # Mode
    parser.add_argument(
        "--batch",
        action="store_true",
        default=False,
        help=(
            "Para XLSX: gera um único arquivo com uma linha por NF-e "
            "(modo resumo). Sem este flag, cada NF-e gera seu próprio XLSX "
            "com abas separadas (somente para um único arquivo)."
        ),
    )

    return parser


def main(argv: List[str] | None = None) -> int:
    """Entry point for the CLI. Returns the exit code."""
    parser = _build_parser()

    # Handle the edge case where arquivos is specified as a positional list but
    # argparse doesn't mark it required because it's nargs="*"
    args = parser.parse_args(argv)
    if not args.dir and not args.arquivos:
        parser.error("Informe um ou mais arquivos XML ou use --dir para especificar um diretório.")

    nfe_parser = NFEParser()
    records: List[NFEData] = []

    if args.dir:
        try:
            records = nfe_parser.parse_directory(args.dir)
        except NFEParseError as exc:
            print(f"Erro: {exc}", file=sys.stderr)
            return 1
        if not records:
            print(f"Nenhum XML de NF-e encontrado em: {args.dir}", file=sys.stderr)
            return 1
    else:
        for filepath in args.arquivos:
            try:
                records.append(nfe_parser.parse_file(filepath))
            except NFEParseError as exc:
                print(f"[AVISO] {exc}", file=sys.stderr)

    if not records:
        print("Nenhuma NF-e pôde ser processada.", file=sys.stderr)
        return 1

    exporter = NFEExporter()

    try:
        if args.format == "csv":
            out = exporter.export_csv(records, args.output)
            print(f"CSV exportado: {out}")
        elif len(records) == 1 and not args.batch:
            out = exporter.export_xlsx(records[0], args.output)
            print(f"XLSX exportado: {out}")
        else:
            out = exporter.export_batch_xlsx(records, args.output)
            print(f"XLSX (batch) exportado: {out}")
    except Exception as exc:  # noqa: BLE001
        print(f"Erro ao exportar: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
