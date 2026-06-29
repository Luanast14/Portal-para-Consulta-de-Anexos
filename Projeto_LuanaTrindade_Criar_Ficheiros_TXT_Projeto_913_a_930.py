from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================


#Faltava criar os ficheiros [nº de projeto]_[empresa/autor]_falta.txt do projeto 913 ao 930



PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

PASTA_REL_AP_PDF = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\04_PROJETOS"
    )

NUMERO_INICIAL = 913
NUMERO_FINAL = 930

# True = só simula, não cria ficheiros.
# False = cria mesmo os ficheiros .txt vazios.
MODO_SIMULACAO = False


# ============================================================
# PROCESSO
# ============================================================

def main():
    if not PASTA_REL_AP_PDF.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {PASTA_REL_AP_PDF}")

    criados = 0
    ja_existiam = 0

    for numero in range(NUMERO_INICIAL, NUMERO_FINAL + 1):
        nome_ficheiro = f"{numero:03d}_AP_falta.txt"
        caminho_ficheiro = PASTA_REL_AP_PDF / nome_ficheiro

        if caminho_ficheiro.exists():
            print(f"já existia: {nome_ficheiro}")
            ja_existiam += 1
            continue

        if not MODO_SIMULACAO:
            caminho_ficheiro.write_text("", encoding="utf-8")

        print(f"{'simulação' if MODO_SIMULACAO else 'criado'}: {nome_ficheiro}")
        criados += 1

    print("\nProcesso concluído.")

    if MODO_SIMULACAO:
        print("Simulação concluída: nenhum ficheiro foi criado.")
    else:
        print(f"Ficheiros criados: {criados}")
        print(f"Ficheiros que já existiam: {ja_existiam}")


if __name__ == "__main__":
    main()