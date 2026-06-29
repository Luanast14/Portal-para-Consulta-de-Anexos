from pathlib import Path
import re
import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES — ALTERAR AQUI
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

PASTA_REL_AP_PDF = Path(
    r"C:\Users\pedromaia\Downloads\ESTRAÇÃO_REAL_FINAL\14_REL_AP_pdf"
)

EXCEL_LISTA_NUMEROS = PASTA_BASE / "Lista_numeros_projeto.xlsx"

# False = cria mesmo os ficheiros .txt vazios.
# True = apenas mostra o que faria, sem criar ficheiros.
MODO_SIMULACAO = False

# O código ignora estes ficheiros ao verificar se já existe projeto.
EXTENSOES_IGNORAR_AO_VERIFICAR = {".txt"}


# ============================================================
# 2. FUNÇÕES AUXILIARES
# ============================================================

def normalizar_codigo_projeto(valor) -> str:
    """
    Normaliza números de projeto vindos do Excel.

    Aceita:
    - 2 -> 002
    - 002 -> 002
    - 019a -> 019a
    - HT3 -> HT003
    - HT003 -> HT003
    - HT003a -> HT003a
    """

    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.lower() in {"", "nan", "none"}:
        return ""

    texto = texto.replace("[", "").replace("]", "").strip()

    # Corrige casos lidos como 2.0, 34.0, etc.
    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]

    texto_sem_espacos = re.sub(r"\s+", "", texto)

    # Projetos HT: HT000 a HT015, com eventual letra final
    match_ht = re.fullmatch(
        r"(?i)HT0*(\d{1,3})([a-z]+)?",
        texto_sem_espacos
    )

    if match_ht:
        numero = int(match_ht.group(1))
        letras = match_ht.group(2) or ""

        return f"HT{numero:03d}{letras.lower()}"

    # Projetos normais: 000 a 912, com eventual letra final
    match_normal = re.fullmatch(
        r"0*(\d{1,3})([a-z]+)?",
        texto_sem_espacos,
        flags=re.IGNORECASE
    )

    if match_normal:
        numero = int(match_normal.group(1))
        letras = match_normal.group(2) or ""

        return f"{numero:03d}{letras.lower()}"

    return texto_sem_espacos


def extrair_codigo_projeto_do_nome(nome_ficheiro: str) -> str:
    """
    Extrai o número de projeto a partir do nome de um ficheiro.

    Ignora ficheiros do tipo:
    021_falta.txt
    """

    nome = Path(nome_ficheiro).name
    stem = Path(nome).stem

    if re.fullmatch(r".+_falta", stem, flags=re.IGNORECASE):
        return ""

    # Primeiro tenta projetos HT: HT000, HT001a, etc.
    match_ht = re.search(
        r"(?i)(?<![A-Z0-9])HT0*(\d{1,3})([a-z]+)?(?![A-Z0-9])",
        stem
    )

    if match_ht:
        numero = int(match_ht.group(1))
        letras = match_ht.group(2) or ""

        return f"HT{numero:03d}{letras.lower()}"

    # Depois tenta projetos normais: 000, 019a, 034b, etc.
    match_normal = re.search(
        r"(?<!\d)0*(\d{1,3})([a-z]+)?(?![A-Z0-9])",
        stem,
        flags=re.IGNORECASE
    )

    if match_normal:
        numero = int(match_normal.group(1))
        letras = match_normal.group(2) or ""

        return f"{numero:03d}{letras.lower()}"

    return ""


def carregar_lista_numeros(excel: Path) -> list[str]:
    """
    Lê todos os valores do Excel, sem assumir cabeçalho.
    Assim não perde o primeiro número da lista.
    """

    if not excel.exists():
        raise FileNotFoundError(f"Excel não encontrado: {excel}")

    todas_folhas = pd.read_excel(
        excel,
        sheet_name=None,
        header=None,
        dtype=str
    )

    numeros = []

    for nome_folha, df in todas_folhas.items():
        for valor in df.to_numpy().flatten():
            codigo = normalizar_codigo_projeto(valor)

            if codigo:
                numeros.append(codigo)

    # Remove duplicados mantendo a ordem
    vistos = set()
    unicos = []

    for numero in numeros:
        if numero not in vistos:
            vistos.add(numero)
            unicos.append(numero)

    return unicos


def carregar_numeros_existentes_na_pasta(pasta: Path) -> set[str]:
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    existentes = set()

    for ficheiro in pasta.iterdir():
        if not ficheiro.is_file():
            continue

        if ficheiro.suffix.lower() in EXTENSOES_IGNORAR_AO_VERIFICAR:
            continue

        codigo = extrair_codigo_projeto_do_nome(ficheiro.name)

        if codigo:
            existentes.add(codigo)

    return existentes


def criar_txt_falta(pasta: Path, numero: str) -> tuple[str, str]:
    nome_txt = f"{numero}_falta.txt"
    caminho_txt = pasta / nome_txt

    if caminho_txt.exists():
        return nome_txt, "já existia"

    if not MODO_SIMULACAO:
        caminho_txt.write_text("", encoding="utf-8")

    return nome_txt, "criado" if not MODO_SIMULACAO else "simulação"


# ============================================================
# 3. PROCESSO PRINCIPAL
# ============================================================

def main():
    numeros_lista = carregar_lista_numeros(EXCEL_LISTA_NUMEROS)
    numeros_existentes = carregar_numeros_existentes_na_pasta(PASTA_REL_AP_PDF)

    numeros_em_falta = [
        numero for numero in numeros_lista
        if numero not in numeros_existentes
    ]

    print(f"Números na lista: {len(numeros_lista)}")
    print(f"Números encontrados em ficheiros existentes: {len(numeros_existentes)}")
    print(f"Números em falta: {len(numeros_em_falta)}")
    print(f"Modo simulação: {MODO_SIMULACAO}")

    criados = 0
    ja_existiam = 0

    for numero in numeros_em_falta:
        nome_txt, estado = criar_txt_falta(PASTA_REL_AP_PDF, numero)

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