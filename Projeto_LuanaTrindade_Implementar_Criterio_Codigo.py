from pathlib import Path
from datetime import datetime
import re
import unicodedata

import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES — ALTERAR AQUI
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

PASTA_REL_AP_PDF = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\14_REL_AP_pdf"
    )

EXCEL_RENOMEACAO = PASTA_BASE / "Projeto_LuanaTrindade_ImplementarCritério.xlsx"

# Se não souberes o nome da folha, deixa None para usar a primeira folha.
FOLHA_EXCEL = None

# True = só cria relatório, NÃO renomeia ficheiros.
# False = renomeia mesmo os ficheiros.
MODO_SIMULACAO = False

# False = só analisa ficheiros diretamente dentro da pasta 14_REL_AP_pdf.
# True = também procura em subpastas.
PROCURAR_RECURSIVAMENTE = False

# Se o nome novo no Excel vier sem extensão, usa a extensão original do ficheiro.
MANTER_EXTENSAO_ORIGINAL_SE_NOME_NOVO_NAO_TIVER = True

RELATORIO_SAIDA = PASTA_BASE / f"relatorio_renomeacao_14_REL_AP_pdf_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


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


def obter_coluna(df: pd.DataFrame, nomes_possiveis: list[str]) -> str:
    mapa = {
        normalizar_texto(coluna): coluna
        for coluna in df.columns
    }

    for nome in nomes_possiveis:
        chave = normalizar_texto(nome)

        if chave in mapa:
            return mapa[chave]

    raise ValueError(
        f"Não encontrei nenhuma destas colunas: {nomes_possiveis}\n"
        f"Colunas existentes no Excel: {list(df.columns)}"
    )


def extrair_nome_simples(valor: str) -> str:
    """
    Recebe nome ou caminho e devolve só o nome do ficheiro.
    Funciona com caminhos Windows.
    """
    texto = str(valor).strip()

    if texto.lower() in {"", "nan", "none"}:
        return ""

    texto = texto.replace("\\", "/")
    return Path(texto).name.strip()


def chave_nome(nome: str, usar_stem: bool = False) -> str:
    nome = extrair_nome_simples(nome)

    if not nome:
        return ""

    p = Path(nome)

    if usar_stem:
        nome = p.stem
    else:
        nome = p.name

    nome = remover_acentos(nome).lower()
    nome = re.sub(r"\s+", " ", nome)
    return nome.strip()


def validar_nome_windows(nome: str) -> tuple[bool, str]:
    """
    Valida se o nome pode ser usado como nome de ficheiro no Windows.
    """
    if not nome or nome.lower() in {"nan", "none"}:
        return False, "nome vazio"

    if re.search(r'[<>:"/\\|?*]', nome):
        return False, "nome contém caracteres inválidos para Windows"

    if nome.endswith(".") or nome.endswith(" "):
        return False, "nome termina com ponto ou espaço"

    reservados = {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }

    stem = Path(nome).stem.upper()

    if stem in reservados:
        return False, "nome reservado pelo Windows"

    return True, "nome válido"


def preparar_nome_novo(nome_novo_excel: str, ficheiro_origem: Path) -> tuple[str, str]:
    """
    Garante que o nome novo é apenas nome de ficheiro e que tem extensão.
    """
    nome_novo = extrair_nome_simples(nome_novo_excel)

    if not nome_novo:
        return "", "nome novo vazio"

    p_novo = Path(nome_novo)

    if not p_novo.suffix and MANTER_EXTENSAO_ORIGINAL_SE_NOME_NOVO_NAO_TIVER:
        nome_novo = nome_novo + ficheiro_origem.suffix

    valido, estado = validar_nome_windows(nome_novo)

    if not valido:
        return "", estado

    return nome_novo, "OK"


def listar_ficheiros_pasta(pasta: Path) -> list[Path]:
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    if PROCURAR_RECURSIVAMENTE:
        return sorted([p for p in pasta.rglob("*") if p.is_file()])

    return sorted([p for p in pasta.glob("*") if p.is_file()])


def criar_indice_ficheiros(ficheiros: list[Path]) -> dict:
    """
    Cria índice para encontrar ficheiros pelo nome atual.
    Indexa tanto pelo nome completo como pelo nome sem extensão.
    """
    indice = {}

    for ficheiro in ficheiros:
        chave_com_ext = chave_nome(ficheiro.name, usar_stem=False)
        chave_sem_ext = chave_nome(ficheiro.name, usar_stem=True)

        if chave_com_ext:
            indice.setdefault(chave_com_ext, []).append(ficheiro)

        if chave_sem_ext:
            indice.setdefault(chave_sem_ext, []).append(ficheiro)

    return indice


def encontrar_ficheiro(nome_atual_excel: str, indice: dict) -> tuple[Path | None, str]:
    """
    Encontra o ficheiro da pasta que corresponde ao Nome atual do Excel.
    """
    nome = extrair_nome_simples(nome_atual_excel)

    if not nome:
        return None, "Nome atual vazio"

    p = Path(nome)

    if p.suffix:
        chave = chave_nome(nome, usar_stem=False)
    else:
        chave = chave_nome(nome, usar_stem=True)

    encontrados = indice.get(chave, [])

    if len(encontrados) == 1:
        return encontrados[0], "ficheiro encontrado"

    if len(encontrados) > 1:
        return None, f"mais do que um ficheiro corresponde ao Nome atual: {len(encontrados)} correspondências"

    return None, "ficheiro não encontrado na pasta"


# ============================================================
# 3. PROCESSO PRINCIPAL
# ============================================================

def main():
    if not EXCEL_RENOMEACAO.exists():
        raise FileNotFoundError(f"Excel não encontrado: {EXCEL_RENOMEACAO}")

    ficheiros_pasta = listar_ficheiros_pasta(PASTA_REL_AP_PDF)
    indice_ficheiros = criar_indice_ficheiros(ficheiros_pasta)

    print(f"Ficheiros encontrados na pasta: {len(ficheiros_pasta)}")
    print(f"Modo simulação: {MODO_SIMULACAO}")

    if FOLHA_EXCEL is None:
        ficheiro_excel = pd.ExcelFile(EXCEL_RENOMEACAO)
        folha_usada = ficheiro_excel.sheet_names[0]
        df = pd.read_excel(EXCEL_RENOMEACAO, sheet_name=folha_usada)
    else:
        folha_usada = FOLHA_EXCEL
        df = pd.read_excel(EXCEL_RENOMEACAO, sheet_name=folha_usada)

    col_nome_atual = obter_coluna(
        df,
        [
            "Nome atual",
            "Nome Atual",
            "nome_atual",
            "Nome do ficheiro atual",
            "Nome do Ficheiro Atual",
            "Ficheiro atual",
            "Ficheiro Atual",
        ]
    )

    col_nome_novo = obter_coluna(
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

    print(f"Folha usada: {folha_usada}")
    print(f"Coluna Nome atual: {col_nome_atual}")
    print(f"Coluna Nome novo: {col_nome_novo}")

    plano = []

    # --------------------------------------------------------
    # A) Criar plano de renomeação
    # --------------------------------------------------------

    for idx, row in df.iterrows():
        nome_atual_excel = row.get(col_nome_atual, "")
        nome_novo_excel = row.get(col_nome_novo, "")

        ficheiro_origem, estado_match = encontrar_ficheiro(
            nome_atual_excel=nome_atual_excel,
            indice=indice_ficheiros,
        )

        if ficheiro_origem is None:
            plano.append({
                "linha_excel": idx + 2,
                "nome_atual_excel": nome_atual_excel,
                "nome_novo_excel": nome_novo_excel,
                "ficheiro_encontrado": False,
                "caminho_origem": "",
                "nome_novo_validado": "",
                "caminho_destino": "",
                "estado_plano": "BLOQUEADO",
                "acao_executada": "não executada",
                "observacoes": estado_match,
            })
            continue

        nome_novo_validado, estado_nome_novo = preparar_nome_novo(
            nome_novo_excel=nome_novo_excel,
            ficheiro_origem=ficheiro_origem,
        )

        if estado_nome_novo != "OK":
            plano.append({
                "linha_excel": idx + 2,
                "nome_atual_excel": nome_atual_excel,
                "nome_novo_excel": nome_novo_excel,
                "ficheiro_encontrado": True,
                "caminho_origem": str(ficheiro_origem),
                "nome_novo_validado": "",
                "caminho_destino": "",
                "estado_plano": "BLOQUEADO",
                "acao_executada": "não executada",
                "observacoes": estado_nome_novo,
            })
            continue

        caminho_destino = ficheiro_origem.with_name(nome_novo_validado)

        if ficheiro_origem.resolve() == caminho_destino.resolve():
            estado_plano = "SEM_ALTERACAO"
            observacoes = "Nome atual já é igual ao Nome novo"
        elif caminho_destino.exists():
            estado_plano = "BLOQUEADO"
            observacoes = "Já existe um ficheiro com o Nome novo no destino"
        else:
            estado_plano = "PRONTO"
            observacoes = "Pronto para renomear"

        plano.append({
            "linha_excel": idx + 2,
            "nome_atual_excel": nome_atual_excel,
            "nome_novo_excel": nome_novo_excel,
            "ficheiro_encontrado": True,
            "caminho_origem": str(ficheiro_origem),
            "nome_novo_validado": nome_novo_validado,
            "caminho_destino": str(caminho_destino),
            "estado_plano": estado_plano,
            "acao_executada": "não executada",
            "observacoes": observacoes,
        })

    df_plano = pd.DataFrame(plano)

    # --------------------------------------------------------
    # B) Bloquear nomes repetidos no Excel
    # --------------------------------------------------------

    if not df_plano.empty:
        prontos = df_plano["estado_plano"].isin(["PRONTO", "SEM_ALTERACAO"])

        duplicados_origem = (
            df_plano.loc[prontos, "caminho_origem"]
            .duplicated(keep=False)
        )

        duplicados_destino = (
            df_plano.loc[prontos, "caminho_destino"]
            .duplicated(keep=False)
        )

        indices_prontos = df_plano.loc[prontos].index

        for idx_local, duplicado in zip(indices_prontos, duplicados_origem):
            if duplicado:
                df_plano.loc[idx_local, "estado_plano"] = "BLOQUEADO"
                df_plano.loc[idx_local, "observacoes"] = "O mesmo ficheiro de origem aparece mais do que uma vez no Excel"

        for idx_local, duplicado in zip(indices_prontos, duplicados_destino):
            if duplicado:
                df_plano.loc[idx_local, "estado_plano"] = "BLOQUEADO"
                df_plano.loc[idx_local, "observacoes"] = "O mesmo Nome novo aparece mais do que uma vez no Excel"

    # --------------------------------------------------------
    # C) Executar renomeação, se MODO_SIMULACAO = False
    # --------------------------------------------------------

    total_renomeados = 0
    total_erros = 0

    if MODO_SIMULACAO:
        df_plano.loc[df_plano["estado_plano"] == "PRONTO", "acao_executada"] = "simulação: não renomeado"
    else:
        for idx, row in df_plano.iterrows():
            if row["estado_plano"] != "PRONTO":
                continue

            origem = Path(row["caminho_origem"])
            destino = Path(row["caminho_destino"])

            try:
                origem.rename(destino)
                df_plano.loc[idx, "acao_executada"] = "renomeado"
                df_plano.loc[idx, "observacoes"] = "Ficheiro renomeado com sucesso"
                total_renomeados += 1

            except Exception as erro:
                df_plano.loc[idx, "estado_plano"] = "ERRO"
                df_plano.loc[idx, "acao_executada"] = "erro"
                df_plano.loc[idx, "observacoes"] = f"Erro ao renomear: {erro}"
                total_erros += 1

    # --------------------------------------------------------
    # D) Identificar ficheiros da pasta sem correspondência no Excel
    # --------------------------------------------------------

    caminhos_usados = set(
        str(Path(c).resolve())
        for c in df_plano["caminho_origem"].dropna().astype(str)
        if c
    )

    ficheiros_sem_linha_excel = []

    for ficheiro in ficheiros_pasta:
        if str(ficheiro.resolve()) not in caminhos_usados:
            ficheiros_sem_linha_excel.append({
                "nome_ficheiro": ficheiro.name,
                "caminho_ficheiro": str(ficheiro),
                "observacao": "Ficheiro existe na pasta, mas não foi encontrado na coluna Nome atual do Excel",
            })

    df_sem_excel = pd.DataFrame(ficheiros_sem_linha_excel)

    # --------------------------------------------------------
    # E) Resumo
    # --------------------------------------------------------

    resumo = pd.DataFrame([
        {"Indicador": "Pasta analisada", "Valor": str(PASTA_REL_AP_PDF)},
        {"Indicador": "Excel analisado", "Valor": str(EXCEL_RENOMEACAO)},
        {"Indicador": "Folha analisada", "Valor": folha_usada},
        {"Indicador": "Coluna Nome atual", "Valor": col_nome_atual},
        {"Indicador": "Coluna Nome novo", "Valor": col_nome_novo},
        {"Indicador": "Modo simulação", "Valor": MODO_SIMULACAO},
        {"Indicador": "Linhas no Excel", "Valor": len(df)},
        {"Indicador": "Ficheiros encontrados na pasta", "Valor": len(ficheiros_pasta)},
        {"Indicador": "Linhas prontas para renomear", "Valor": int((df_plano["estado_plano"] == "PRONTO").sum()) if not df_plano.empty else 0},
        {"Indicador": "Linhas bloqueadas", "Valor": int((df_plano["estado_plano"] == "BLOQUEADO").sum()) if not df_plano.empty else 0},
        {"Indicador": "Linhas sem alteração necessária", "Valor": int((df_plano["estado_plano"] == "SEM_ALTERACAO").sum()) if not df_plano.empty else 0},
        {"Indicador": "Ficheiros renomeados", "Valor": total_renomeados},
        {"Indicador": "Erros ao renomear", "Valor": total_erros},
        {"Indicador": "Ficheiros da pasta sem linha no Excel", "Valor": len(df_sem_excel)},
    ])

    # --------------------------------------------------------
    # F) Exportar relatório
    # --------------------------------------------------------

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        df_plano.to_excel(writer, index=False, sheet_name="Plano_renomeacao")
        df_sem_excel.to_excel(writer, index=False, sheet_name="Ficheiros_sem_Excel")
        resumo.to_excel(writer, index=False, sheet_name="Resumo")

    print("\nProcesso concluído.")
    print(f"Relatório criado em: {RELATORIO_SAIDA}")

    if MODO_SIMULACAO:
        print("Modo simulação ativo: nenhum ficheiro foi renomeado.")
        print("Se o relatório estiver correto, muda MODO_SIMULACAO para False.")
    else:
        print(f"Ficheiros renomeados: {total_renomeados}")
        print(f"Erros ao renomear: {total_erros}")


if __name__ == "__main__":
    main()