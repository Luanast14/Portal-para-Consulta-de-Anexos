from pathlib import Path
import re


# ============================================================
# 1. CONFIGURAÇÕES — ALTERAR AQUI
# ============================================================


#Ficheiro para que as pastas criadas ignorem as letras minúsculas à frente dos nº de projetos
#(e.g. existem dois ficheiros, 178a_... e 178b_...; o código só deve criar uma única pasta com o
#nome 178_...)




PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

PASTA_REL_AP_PDF = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\04_PROJETOS"
)
# False = cria mesmo os ficheiros .txt vazios.
# True = só mostra o que faria, sem criar ficheiros.
MODO_SIMULACAO = False

# Ignora estes ficheiros ao analisar a pasta.
EXTENSOES_IGNORAR = {".txt"}

# O número AP pode ir até 930.
LIMITE_MAXIMO_AP = 930

# Modo recomendado:
# "MIN_MAX_EXISTENTE" = cria faltas entre o menor e o maior número AP encontrado.
# "000_930" = cria faltas de 000 até 930.
MODO_INTERVALO_AP = "MIN_MAX_EXISTENTE"


# ============================================================
# 2. FUNÇÕES AUXILIARES
# ============================================================

def extrair_numero_ap_do_nome(nome_ficheiro: str) -> int | None:
    """
    Extrai apenas o número inteiro de ficheiros AP.

    Exemplos:
    134_AP_REL_IA_BJA.pdf    -> 134
    140a_AP_REL_IA_BJA.pdf   -> 140
    140b_AP_REL_IA_BJA.pdf   -> 140
    034c_AP_REL_CR_CHV.pdf   -> 34

    Ignora:
    010_HT_REL_IA_BJA.pdf
    135_AP_falta.txt
    """

    caminho = Path(nome_ficheiro)
    stem = caminho.stem

    # Ignorar ficheiros de falta
    if re.search(r"_falta$", stem, flags=re.IGNORECASE):
        return None

    # Separar por underscore, hífen ou espaços
    partes = re.split(r"[_\-\s]+", stem)

    if len(partes) < 2:
        return None

    primeiro_bloco = partes[0]
    segundo_bloco = partes[1].upper()

    # Só queremos ficheiros AP
    if segundo_bloco != "AP":
        return None

    # Aceita 134, 140a, 140b, 034c, etc.
    match = re.fullmatch(r"0*(\d{1,3})([a-z])?", primeiro_bloco, flags=re.IGNORECASE)

    if not match:
        return None

    numero = int(match.group(1))

    if 0 <= numero <= LIMITE_MAXIMO_AP:
        return numero

    return None


def carregar_numeros_ap_existentes(pasta: Path) -> set[int]:
    """
    Lê os ficheiros da pasta e devolve os números AP existentes.
    As letras são ignoradas.

    Exemplo:
    140a_AP_...
    140b_AP_...

    devolve apenas:
    {140}
    """

    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    numeros = set()

    for ficheiro in pasta.iterdir():
        if not ficheiro.is_file():
            continue

        if ficheiro.suffix.lower() in EXTENSOES_IGNORAR:
            continue

        numero = extrair_numero_ap_do_nome(ficheiro.name)

        if numero is not None:
            numeros.add(numero)

    return numeros


def definir_intervalo_a_verificar(numeros_existentes: set[int]) -> list[int]:
    if not numeros_existentes:
        return []

    if MODO_INTERVALO_AP == "000_930":
        return list(range(0, LIMITE_MAXIMO_AP + 1))

    menor = min(numeros_existentes)
    maior = max(numeros_existentes)

    return list(range(menor, maior + 1))


def criar_txt_falta(numero: int) -> tuple[str, str]:
    numero_formatado = f"{numero:03d}"
    nome_txt = f"{numero_formatado}_AP_falta.txt"
    caminho_txt = PASTA_REL_AP_PDF / nome_txt

    if caminho_txt.exists():
        return nome_txt, "já existia"

    if not MODO_SIMULACAO:
        caminho_txt.write_text("", encoding="utf-8")

    return nome_txt, "criado" if not MODO_SIMULACAO else "simulação"


# ============================================================
# 3. PROCESSO PRINCIPAL
# ============================================================

def main():
    numeros_existentes = carregar_numeros_ap_existentes(PASTA_REL_AP_PDF)
    intervalo = definir_intervalo_a_verificar(numeros_existentes)

    numeros_em_falta = [
        numero for numero in intervalo
        if numero not in numeros_existentes
    ]

    print(f"Números AP existentes encontrados: {len(numeros_existentes)}")

    if numeros_existentes:
        print(f"Menor número AP encontrado: {min(numeros_existentes):03d}")
        print(f"Maior número AP encontrado: {max(numeros_existentes):03d}")

    print(f"Modo de intervalo: {MODO_INTERVALO_AP}")
    print(f"Números AP em falta: {len(numeros_em_falta)}")
    print(f"Modo simulação: {MODO_SIMULACAO}")

    criados = 0
    ja_existiam = 0

    for numero in numeros_em_falta:
        nome_txt, estado = criar_txt_falta(numero)

        if estado == "criado":
            criados += 1
        elif estado == "já existia":
            ja_existiam += 1

        print(f"{estado}: {nome_txt}")

    print("\nProcesso concluído.")

    if MODO_SIMULACAO:
        print("Simulação concluída: nenhum ficheiro foi criado.")
    else:
        print(f"Ficheiros .txt criados: {criados}")
        print(f"Ficheiros .txt que já existiam: {ja_existiam}")


if __name__ == "__main__":
    main()