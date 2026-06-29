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
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\04_PROJETOS"
)

EXCEL_LISTA_NUMEROS = PASTA_BASE / "Lista_numeros_projeto.xlsx"

# False = cria mesmo os ficheiros .txt vazios.
# True = só simula, não cria ficheiros.
MODO_SIMULACAO = False

# O código ignora ficheiros .txt ao verificar se o projeto já existe.
EXTENSOES_IGNORAR_AO_VERIFICAR = {".txt"}

# Projetos HT que devem existir como falta, se não houver ficheiro correspondente.
HT_NUMEROS_ESPERADOS = {"010", "016", "017", "018"}

# Projetos AP podem ir até 930.
LIMITE_MAXIMO_AP = 930


# ============================================================
# 2. FUNÇÕES AUXILIARES
# ============================================================

def normalizar_numero(valor) -> str:
    """
    Normaliza números vindos do Excel.

    Exemplos:
    1 -> 001
    1.0 -> 001
    001 -> 001
    34a -> 034a
    """

    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.lower() in {"", "nan", "none"}:
        return ""

    texto = texto.replace("[", "").replace("]", "").strip()
    texto = re.sub(r"\s+", "", texto)

    # Corrige valores lidos como 1.0, 34.0, etc.
    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]

    # Remove eventual AP/HT vindo do Excel, mas guarda só o número.
    texto = re.sub(r"(?i)^AP", "", texto)
    texto = re.sub(r"(?i)^HT", "", texto)

    match = re.fullmatch(r"0*(\d{1,3})([a-z]+)?", texto, flags=re.IGNORECASE)

    if not match:
        return ""

    numero = int(match.group(1))
    letras = match.group(2) or ""

    return f"{numero:03d}{letras.lower()}"


def carregar_lista_numeros_ap(excel: Path) -> list[str]:
    """
    Lê todos os números do Excel.
    Considera estes números como projetos AP.
    Só aceita números até 930.
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

    for _nome_folha, df in todas_folhas.items():
        for valor in df.to_numpy().flatten():
            numero = normalizar_numero(valor)

            if not numero:
                continue

            numero_base = int(numero[:3])

            if 0 <= numero_base <= LIMITE_MAXIMO_AP:
                numeros.append(numero)

    # Remover duplicados mantendo ordem
    vistos = set()
    unicos = []

    for numero in numeros:
        if numero not in vistos:
            vistos.add(numero)
            unicos.append(numero)

    return unicos


def extrair_projeto_empresa_do_nome(nome_ficheiro: str) -> tuple[str, str]:
    """
    Extrai número de projeto e empresa do nome de um ficheiro existente.

    Exemplos:
    010_AP_REL_IA_BJA.pdf -> 010, AP
    010_HT_REL_IA_BJA.pdf -> 010, HT
    034a_AP_REL_CR_CHV.pdf -> 034a, AP
    """

    nome = Path(nome_ficheiro).name
    stem = Path(nome).stem

    # Ignorar ficheiros do tipo *_falta
    if re.search(r"_falta$", stem, flags=re.IGNORECASE):
        return "", ""

    partes = re.split(r"[_\-\s]+", stem)

    numero = ""
    empresa = ""

    # Caso normalizado esperado: 010_AP_...
    if len(partes) >= 2:
        possivel_numero = normalizar_numero(partes[0])
        possivel_empresa = partes[1].upper()

        if possivel_numero and possivel_empresa in {"AP", "HT"}:
            return possivel_numero, possivel_empresa

    # Fallback: procurar número e empresa no nome
    match_numero = re.search(r"(?<!\d)(\d{3}[a-z]?)(?!\d)", stem, flags=re.IGNORECASE)

    if match_numero:
        numero = normalizar_numero(match_numero.group(1))

    if re.search(r"(?<![A-Z0-9])AP(?![A-Z0-9])", stem, flags=re.IGNORECASE):
        empresa = "AP"
    elif re.search(r"(?<![A-Z0-9])HT(?![A-Z0-9])", stem, flags=re.IGNORECASE):
        empresa = "HT"

    return numero, empresa


def carregar_projetos_existentes(pasta: Path) -> set[tuple[str, str]]:
    """
    Devolve pares existentes:
    {("010", "AP"), ("010", "HT"), ...}
    """

    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    existentes = set()

    for ficheiro in pasta.iterdir():
        if not ficheiro.is_file():
            continue

        if ficheiro.suffix.lower() in EXTENSOES_IGNORAR_AO_VERIFICAR:
            continue

        numero, empresa = extrair_projeto_empresa_do_nome(ficheiro.name)

        if numero and empresa:
            existentes.add((numero, empresa))

    return existentes


def criar_txt_falta(pasta: Path, numero: str, empresa: str) -> tuple[str, str]:
    """
    Cria ficheiro vazio:
    xxx_AP_falta.txt
    xxx_HT_falta.txt
    """

    nome_txt = f"{numero}_{empresa}_falta.txt"
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
    numeros_ap = carregar_lista_numeros_ap(EXCEL_LISTA_NUMEROS)
    projetos_existentes = carregar_projetos_existentes(PASTA_REL_AP_PDF)

    print(f"Números AP na lista: {len(numeros_ap)}")
    print(f"Projetos existentes encontrados na pasta: {len(projetos_existentes)}")
    print(f"Projetos HT esperados: {', '.join(sorted(HT_NUMEROS_ESPERADOS))}")
    print(f"Modo simulação: {MODO_SIMULACAO}")

    criados = 0
    ja_existiam = 0

    # --------------------------------------------------------
    # A) Criar faltas AP
    # --------------------------------------------------------

    for numero in numeros_ap:
        if (numero, "AP") in projetos_existentes:
            continue

        nome_txt, estado = criar_txt_falta(
            pasta=PASTA_REL_AP_PDF,
            numero=numero,
            empresa="AP"
        )

        if estado == "criado":
            criados += 1
        elif estado == "já existia":
            ja_existiam += 1

        print(f"{estado}: {nome_txt}")

    # --------------------------------------------------------
    # B) Criar faltas HT apenas para 010, 016, 017, 018
    # --------------------------------------------------------

    for numero in sorted(HT_NUMEROS_ESPERADOS):
        if (numero, "HT") in projetos_existentes:
            continue

        nome_txt, estado = criar_txt_falta(
            pasta=PASTA_REL_AP_PDF,
            numero=numero,
            empresa="HT"
        )

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