# -*- coding: utf-8 -*-
"""Iniciador unificado para geração de documentações do Tableau e Power BI."""

import sys

def print_usage():
    print("Uso: python main.py [T|P] [caminho_do_arquivo] [--format all|markdown|json|excel|rtf|docx]")
    print("")
    print("Onde [T|P] indica o sistema de origem:")
    print("  T - Tableau")
    print("  P - Power BI")

def main():
    if len(sys.argv) < 2:
        print_usage()
        sys.exit(1)

    system_arg = sys.argv[1].upper()
    
    if system_arg not in ('T', 'P'):
        print_usage()
        print(f"\nErro: O primeiro parâmetro deve ser 'T' ou 'P'. Recebido: {sys.argv[1]}")
        sys.exit(1)

    # Remove o parâmetro T/P de sys.argv para que o argparse do parser interno funcione normalmente
    sys.argv.pop(1)

    if system_arg == 'T':
        from tableaudoc.tableau_doc import main as tableau_main
        tableau_main()
    elif system_arg == 'P':
        from tableaudoc.powerbi_doc import main as powerbi_main
        powerbi_main()

if __name__ == "__main__":
    main()
