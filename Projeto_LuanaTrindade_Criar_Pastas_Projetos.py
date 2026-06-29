from pathlib import Path
from datetime import datetime
import re
import unicodedata

import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

EXCEL_NOMENCLATURA = PASTA_BASE / "Implementar_Critério_Nomenclatura_REL_AP_pdf.xlsx"

# Se não souberes a folha, deixa None para usar a primeira.
FOLHA_EXCEL = None

# Se não souberes o nome exato da coluna, deixa None.
# O código vai tentar encontrar "Nome novo" automaticamente.
COLUNA_NOME_NOVO = None

# True = não cria nem renomeia nada; só cria relatório.
# False = cria/renomeia mesmo as pastas.
MODO_SIMULACAO = True

# True = se já existir Projeto_034, tenta renomear para Projeto_034_CHV.
# False = não renomeia pastas antigas; apenas cria as novas.
RENOMEAR_PASTAS_EXISTENTES = True

# True = cria também as subpastas dentro de cada pasta de projeto.
CRIAR_SUBPASTAS = True

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
]

RELATORIO_SAIDA = PASTA_BASE / f"relatorio_pastas_projeto_municipio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


# ============================================================
# 2. FUNÇÕES AUXILIARES
# ============================================================

def remover_acentos(texto: str) -> str:
    texto = unicodedata.normalize("NFD", str(texto))
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def normalizar_texto(texto: str) -> str:
    texto = remover_acentos(str(texto))
    texto = texto.lower().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def encontrar_coluna(df: pd.DataFrame, nomes_possiveis: list[str]) -> str:
    mapa = {
        normalizar_texto(col): col
        for col in df.columns
    }

    for nome in nomes_possiveis:
        chave = normalizar_texto(nome)

        if chave in mapa:
            return mapa[chave]

    raise ValueError(
        "Não encontrei a coluna necessária.\n"
        f"Colunas existentes no Excel: {list(df.columns)}\n"
        f"Colunas procuradas: {nomes_possiveis}"
    )


def limpar_nome_ficheiro(valor) -> str:
    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.lower() in {"", "nan", "none"}:
        return ""

    texto = texto.replace("\\", "/")
    return Path(texto).name.strip()


def extrair_info_nome_novo(nome_ficheiro: str) -> dict:
    """
    Extrai número de projeto e código de município a partir de nomes do tipo:

    034c_AP_REL_CR_CHV.pdf
    140a_AP_REL_IA_BJA.pdf
    010_HT_REL_IA_CHV.pdf

    A letra do projeto é ignorada para criar a pasta:
    034c -> Projeto_034_CHV
    140a -> Projeto_140_BJA
    """

    nome = limpar_nome_ficheiro(nome_ficheiro)

    if not nome:
        return {
            "valido": False,
            "nome_ficheiro": nome_ficheiro,
            "numero_projeto": "",
            "empresa": "",
            "codigo_municipio": "",
            "nome_pasta": "",
            "motivo": "Nome vazio"
        }

    stem = Path(nome).stem

    # Ignorar ficheiros do tipo 913_AP_falta.txt
    if re.search(r"_falta$", stem, flags=re.IGNORECASE):
        return {
            "valido": False,
            "nome_ficheiro": nome,
            "numero_projeto": "",
            "empresa": "",
            "codigo_municipio": "",
            "nome_pasta": "",
            "motivo": "Ficheiro de falta; não tem código de município"
        }

    partes = stem.split("_")

    if len(partes) < 5:
        return {
            "valido": False,
            "nome_ficheiro": nome,
            "numero_projeto": "",
            "empresa": "",
            "codigo_municipio": "",
            "nome_pasta": "",
            "motivo": "Nome não tem estrutura suficiente"
        }

    bloco_projeto = partes[0].strip()
    empresa = partes[1].strip().upper()
    codigo_municipio = partes[-1].strip().upper()

    # Aceita 034, 034a, 034aa, 140b, etc.
    match_projeto = re.fullmatch(
        r"0*(\d{1,3})([a-z]+)?",
        bloco_projeto,
        flags=re.IGNORECASE
    )

    if not match_projeto:
        return {
            "valido": False,
            "nome_ficheiro": nome,
            "numero_projeto": "",
            "empresa": empresa,
            "codigo_municipio": codigo_municipio,
            "nome_pasta": "",
            "motivo": "Não foi possível extrair número de projeto"
        }

    numero = int(match_projeto.group(1))
    numero_formatado = f"{numero:03d}"

    if empresa not in {"AP", "HT"}:
        return {
            "valido": False,
            "nome_ficheiro": nome,
            "numero_projeto": numero_formatado,
            "empresa": empresa,
            "codigo_municipio": codigo_municipio,
            "nome_pasta": "",
            "motivo": "Empresa não reconhecida"
        }

    if not re.fullmatch(r"[A-Z]{2,6}", codigo_municipio):
        return {
            "valido": False,
            "nome_ficheiro": nome,
            "numero_projeto": numero_formatado,
            "empresa": empresa,
            "codigo_municipio": codigo_municipio,
            "nome_pasta": "",
            "motivo": "Código de município inválido ou ausente"
        }

    nome_pasta = f"Projeto_{numero_formatado}_{codigo_municipio}"

    return {
        "valido": True,
        "nome_ficheiro": nome,
        "numero_projeto": numero_formatado,
        "empresa": empresa,
        "codigo_municipio": codigo_municipio,
        "nome_pasta": nome_pasta,
        "motivo": "OK"
    }


def ler_excel_nomenclatura() -> pd.DataFrame:
    if not EXCEL_NOMENCLATURA.exists():
        raise FileNotFoundError(f"Excel não encontrado: {EXCEL_NOMENCLATURA}")

    if FOLHA_EXCEL is None:
        xls = pd.ExcelFile(EXCEL_NOMENCLATURA)
        folha = xls.sheet_names[0]
    else:
        folha = FOLHA_EXCEL

    df = pd.read_excel(EXCEL_NOMENCLATURA, sheet_name=folha)

    if COLUNA_NOME_NOVO is None:
        coluna_nome_novo = encontrar_coluna(
            df,
            [
                "Nome novo",
                "Nome Novo",
                "nome_novo",
                "Novo nome",
                "Novo Nome",
                "Nome final",
                "Nome Final",
            ]
        )
    else:
        coluna_nome_novo = COLUNA_NOME_NOVO

    df["_folha_usada"] = folha
    df["_coluna_nome_novo"] = coluna_nome_novo

    return df


def criar_subpastas(pasta_projeto: Path) -> tuple[int, int]:
    criadas = 0
    ja_existiam = 0

    if not CRIAR_SUBPASTAS:
        return criadas, ja_existiam

    for subpasta in SUBPASTAS_CRIAR:
        caminho_subpasta = pasta_projeto / subpasta

        if caminho_subpasta.exists():
            ja_existiam += 1
            continue

        if not MODO_SIMULACAO:
            caminho_subpasta.mkdir(parents=True, exist_ok=True)

        criadas += 1

    return criadas, ja_existiam


def criar_ou_renomear_pasta(nome_pasta_novo: str, numero_projeto: str) -> dict:
    pasta_nova = PASTA_BASE / nome_pasta_novo
    pasta_antiga = PASTA_BASE / f"Projeto_{numero_projeto}"

    resultado = {
        "pasta_antiga": str(pasta_antiga),
        "pasta_nova": str(pasta_nova),
        "acao": "",
        "estado": "",
        "subpastas_criadas": 0,
        "subpastas_ja_existiam": 0,
    }

    # Caso 1: a pasta nova já existe
    if pasta_nova.exists():
        sub_criadas, sub_existiam = criar_subpastas(pasta_nova)

        resultado["acao"] = "usar pasta existente"
        resultado["estado"] = "pasta nova já existia"
        resultado["subpastas_criadas"] = sub_criadas
        resultado["subpastas_ja_existiam"] = sub_existiam
        return resultado

    # Caso 2: existe Projeto_xxx e pode ser renomeada
    if RENOMEAR_PASTAS_EXISTENTES and pasta_antiga.exists():
        if not MODO_SIMULACAO:
            pasta_antiga.rename(pasta_nova)

        sub_criadas, sub_existiam = criar_subpastas(pasta_nova)

        resultado["acao"] = "renomear pasta antiga"
        resultado["estado"] = "renomeada" if not MODO_SIMULACAO else "simulação: seria renomeada"
        resultado["subpastas_criadas"] = sub_criadas
        resultado["subpastas_ja_existiam"] = sub_existiam
        return resultado

    # Caso 3: criar pasta nova
    if not MODO_SIMULACAO:
        pasta_nova.mkdir(parents=True, exist_ok=True)

    sub_criadas, sub_existiam = criar_subpastas(pasta_nova)

    resultado["acao"] = "criar pasta nova"
    resultado["estado"] = "criada" if not MODO_SIMULACAO else "simulação: seria criada"
    resultado["subpastas_criadas"] = sub_criadas
    resultado["subpastas_ja_existiam"] = sub_existiam
    return resultado


# ============================================================
# 3. PROCESSO PRINCIPAL
# ============================================================

def main():
    df = ler_excel_nomenclatura()

    folha_usada = df["_folha_usada"].iloc[0]
    coluna_nome_novo = df["_coluna_nome_novo"].iloc[0]

    print(f"Excel analisado: {EXCEL_NOMENCLATURA}")
    print(f"Folha usada: {folha_usada}")
    print(f"Coluna usada: {coluna_nome_novo}")
    print(f"Pasta base: {PASTA_BASE}")
    print(f"Modo simulação: {MODO_SIMULACAO}")
    print()

    registos_extraidos = []

    for idx, row in df.iterrows():
        nome_novo = row.get(coluna_nome_novo, "")
        info = extrair_info_nome_novo(nome_novo)

        registos_extraidos.append({
            "linha_excel": idx + 2,
            "nome_novo": nome_novo,
            **info
        })

    df_extraidos = pd.DataFrame(registos_extraidos)

    df_validos = df_extraidos[df_extraidos["valido"] == True].copy()

    # Uma pasta por combinação Projeto_xxx + código município
    df_pastas = (
        df_validos[["numero_projeto", "codigo_municipio", "nome_pasta"]]
        .drop_duplicates()
        .sort_values(["numero_projeto", "codigo_municipio"])
        .reset_index(drop=True)
    )

    # Identificar projetos com mais do que um código de município
    conflitos = (
        df_pastas
        .groupby("numero_projeto")["codigo_municipio"]
        .nunique()
        .reset_index(name="total_codigos_municipio")
    )

    projetos_varios_codigos = set(
        conflitos.loc[
            conflitos["total_codigos_municipio"] > 1,
            "numero_projeto"
        ]
    )

    relatorio_pastas = []

    for _, row in df_pastas.iterrows():
        numero = row["numero_projeto"]
        codigo = row["codigo_municipio"]
        nome_pasta = row["nome_pasta"]

        acao = criar_ou_renomear_pasta(
            nome_pasta_novo=nome_pasta,
            numero_projeto=numero
        )

        observacao = "OK"

        if numero in projetos_varios_codigos:
            observacao = (
                "Atenção: este número de projeto aparece associado "
                "a mais do que um código de município no Excel"
            )

        relatorio_pastas.append({
            "numero_projeto": numero,
            "codigo_municipio": codigo,
            "nome_pasta": nome_pasta,
            "observacao": observacao,
            **acao
        })

        print(f"{acao['estado']}: {nome_pasta}")

    df_relatorio_pastas = pd.DataFrame(relatorio_pastas)

    resumo = pd.DataFrame([
        {"Indicador": "Pasta base", "Valor": str(PASTA_BASE)},
        {"Indicador": "Excel analisado", "Valor": str(EXCEL_NOMENCLATURA)},
        {"Indicador": "Folha analisada", "Valor": folha_usada},
        {"Indicador": "Coluna analisada", "Valor": coluna_nome_novo},
        {"Indicador": "Modo simulação", "Valor": MODO_SIMULACAO},
        {"Indicador": "Renomear pastas Projeto_xxx existentes", "Valor": RENOMEAR_PASTAS_EXISTENTES},
        {"Indicador": "Criar subpastas internas", "Valor": CRIAR_SUBPASTAS},
        {"Indicador": "Linhas no Excel", "Valor": len(df)},
        {"Indicador": "Nomes válidos extraídos", "Valor": len(df_validos)},
        {"Indicador": "Pastas Projeto_xxx_codigo únicas", "Valor": len(df_pastas)},
        {"Indicador": "Projetos com vários códigos de município", "Valor": len(projetos_varios_codigos)},
    ])

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        df_relatorio_pastas.to_excel(writer, index=False, sheet_name="Pastas_criadas")
        df_extraidos.to_excel(writer, index=False, sheet_name="Extracao_nomes")
        conflitos.to_excel(writer, index=False, sheet_name="Conflitos_codigos")
        resumo.to_excel(writer, index=False, sheet_name="Resumo")

    print("\nProcesso concluído.")
    print(f"Relatório criado em: {RELATORIO_SAIDA}")

    if MODO_SIMULACAO:
        print("Simulação ativa: nenhuma pasta foi criada ou renomeada.")
        print("Se estiver correto, muda MODO_SIMULACAO para False.")
    else:
        print("Pastas criadas/renomeadas com sucesso, quando aplicável.")


if __name__ == "__main__":
    main()