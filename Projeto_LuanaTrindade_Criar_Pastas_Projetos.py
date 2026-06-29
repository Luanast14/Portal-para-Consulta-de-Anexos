from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

# True = só simula, não cria pastas.
# False = cria mesmo as pastas.
MODO_SIMULACAO = False

NUMERO_FINAL = 930


# ============================================================
# PROCESSO
# ============================================================

def gerar_nomes_pastas() -> list[str]:
    nomes = []

    # Casos especiais pedidos
    nomes.append("Projeto_0")
    nomes.append("Projeto_00")
    nomes.append("Projeto_000")

    # De 001 até 930
    for numero in range(1, NUMERO_FINAL + 1):
        nomes.append(f"Projeto_{numero:03d}")

    return nomes


def main():
    if not PASTA_BASE.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {PASTA_BASE}")

    nomes_pastas = gerar_nomes_pastas()

    criadas = 0
    ja_existiam = 0

    print(f"Total de pastas previstas: {len(nomes_pastas)}")
    print(f"Pasta base: {PASTA_BASE}")
    print(f"Modo simulação: {MODO_SIMULACAO}")
    print()

    for nome_pasta in nomes_pastas:
        caminho_pasta = PASTA_BASE / nome_pasta

        if caminho_pasta.exists():
            print(f"já existia: {nome_pasta}")
            ja_existiam += 1
            continue

        if not MODO_SIMULACAO:
            caminho_pasta.mkdir(parents=True, exist_ok=True)

        print(f"{'simulação' if MODO_SIMULACAO else 'criada'}: {nome_pasta}")
        criadas += 1

    print("\nProcesso concluído.")

    if MODO_SIMULACAO:
        print("Simulação concluída: nenhuma pasta foi criada.")
        print("Para criar mesmo as pastas, muda MODO_SIMULACAO para False.")
    else:
        print(f"Pastas criadas: {criadas}")
        print(f"Pastas que já existiam: {ja_existiam}")


if __name__ == "__main__":
    main()