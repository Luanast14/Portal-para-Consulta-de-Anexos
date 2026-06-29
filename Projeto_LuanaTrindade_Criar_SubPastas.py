from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\04_PROJETOS"
)

# True = só simula, não cria subpastas.
# False = cria mesmo as subpastas.
MODO_SIMULACAO = False

NUMERO_FINAL = 930

SUBPASTAS_CRIAR = [
    "REL_docx",
    "REL_pdf",
    "ESPOLIO",
    "RESTAURO",
    "DESENHO",
    "DIGITALIZACOES",
    "SIG",
    "FOTO_ORG",
    "FOTO_COPIA",
    "FTGRM",
    "ADMN_PROJ",
    "BIBLIO_PROJ",
]


# ============================================================
# FUNÇÕES
# ============================================================

def gerar_pastas_projeto() -> list[Path]:
    pastas = []

    pastas.append(PASTA_BASE / "Projeto_0")
    pastas.append(PASTA_BASE / "Projeto_00")
    pastas.append(PASTA_BASE / "Projeto_000")

    for numero in range(1, NUMERO_FINAL + 1):
        pastas.append(PASTA_BASE / f"Projeto_{numero:03d}")

    return pastas


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def main():
    if not PASTA_BASE.exists():
        raise FileNotFoundError(f"Pasta base não encontrada: {PASTA_BASE}")

    pastas_projeto = gerar_pastas_projeto()

    total_subpastas_previstas = len(pastas_projeto) * len(SUBPASTAS_CRIAR)

    criadas = 0
    ja_existiam = 0
    projetos_inexistentes = 0

    print(f"Pasta base: {PASTA_BASE}")
    print(f"Pastas de projeto previstas: {len(pastas_projeto)}")
    print(f"Subpastas por projeto: {len(SUBPASTAS_CRIAR)}")
    print(f"Total de subpastas previstas: {total_subpastas_previstas}")
    print(f"Modo simulação: {MODO_SIMULACAO}")
    print()

    for pasta_projeto in pastas_projeto:
        if not pasta_projeto.exists():
            print(f"Pasta de projeto não existe: {pasta_projeto.name}")
            projetos_inexistentes += 1
            continue

        for nome_subpasta in SUBPASTAS_CRIAR:
            caminho_subpasta = pasta_projeto / nome_subpasta

            if caminho_subpasta.exists():
                print(f"já existia: {pasta_projeto.name}\\{nome_subpasta}")
                ja_existiam += 1
                continue

            if not MODO_SIMULACAO:
                caminho_subpasta.mkdir(parents=True, exist_ok=True)

            print(
                f"{'simulação' if MODO_SIMULACAO else 'criada'}: "
                f"{pasta_projeto.name}\\{nome_subpasta}"
            )
            criadas += 1

    print("\nProcesso concluído.")

    if MODO_SIMULACAO:
        print("Simulação concluída: nenhuma subpasta foi criada.")
        print("Para criar mesmo as subpastas, muda MODO_SIMULACAO para False.")
    else:
        print(f"Subpastas criadas: {criadas}")
        print(f"Subpastas que já existiam: {ja_existiam}")
        print(f"Pastas de projeto inexistentes: {projetos_inexistentes}")


if __name__ == "__main__":
    main()