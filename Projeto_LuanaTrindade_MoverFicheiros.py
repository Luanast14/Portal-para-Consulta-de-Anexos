from pathlib import Path
from datetime import datetime
import re
import shutil

import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES
# ============================================================

PASTA_BASE = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL"
)

PASTA_ORIGEM = Path(
    r"C:\Users\pedromaia\Downloads\EXTRAÇÃO_REAL_FINAL\04_PROJETOS"
)

# True = só simula, não move ficheiros.
# False = move mesmo os ficheiros.
MODO_SIMULACAO = False

# True = cria a subpasta de destino se ela ainda não existir.
# False = bloqueia se a subpasta não existir.
CRIAR_SUBPASTA_DESTINO_SE_NAO_EXISTIR = True

RELATORIO_SAIDA = PASTA_BASE / f"relatorio_mover_ficheiros_para_projetos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


# ============================================================
# 2. MAPA: TIPO DE DOCUMENTO -> SUBPASTA
# ============================================================

MAPA_TIPO_DOCUMENTO = {
    "REL": "REL_pdf",
    "BD": "ESPOLIO",
    "CCP": "ADMN_PROJ",
    "CE": "ADMN_PROJ",
    "DADM": "ADMN_PROJ",
    "EIA": "REL_pdf",
    "EPP": "REL_pdf",
    "FOT": "FOTO_COPIA",
    "IE": "ESPOLIO",
    "MDR": "ADMN_PROJ",
    "BIB": "BIBLIO_PROJ",
    "PCR": "REL_pdf",
    "PM": "REL_pdf",
    "RT": "SIG",
    "TAB": "ESPOLIO",
}


# ============================================================
# 3. FUNÇÕES AUXILIARES
# ============================================================

def normalizar_tipo_documento(valor: str) -> str:
    """
    Normaliza códigos como DAdm, dadm, DADM para DADM.
    """
    return str(valor).strip().upper()


def extrair_numero_base(bloco_numero: str) -> str:
    """
    Extrai apenas o número inteiro do projeto, ignorando letras.

    Exemplos:
    047a -> 047
    047b -> 047
    25   -> 025
    025  -> 025
    """

    texto = str(bloco_numero).strip()

    match = re.fullmatch(r"0*(\d{1,3})([a-z]+)?", texto, flags=re.IGNORECASE)

    if not match:
        return ""

    numero = int(match.group(1))

    return f"{numero:03d}"


def encontrar_pasta_projeto(numero_projeto: str, codigo_municipio: str = "") -> tuple[Path | None, str]:
    """
    Procura a pasta de projeto correspondente.

    Para ficheiros normais:
    Projeto_025_PRT

    Para ficheiros *_falta.txt:
    Projeto_042

    Se não encontrar exatamente, tenta uma alternativa segura.
    """

    if codigo_municipio:
        pasta_com_codigo = PASTA_BASE / f"Projeto_{numero_projeto}_{codigo_municipio}"

        if pasta_com_codigo.exists():
            return pasta_com_codigo, "pasta com código de município encontrada"

        return None, f"não existe a pasta Projeto_{numero_projeto}_{codigo_municipio}"

    pasta_sem_codigo = PASTA_BASE / f"Projeto_{numero_projeto}"

    if pasta_sem_codigo.exists():
        return pasta_sem_codigo, "pasta sem código de município encontrada"

    # Fallback: se existir exatamente uma pasta do género Projeto_042_XXX,
    # usa essa pasta. Isto ajuda caso a pasta já tenha sido renomeada com município.
    candidatas = [
        p for p in PASTA_BASE.iterdir()
        if p.is_dir() and re.fullmatch(
            rf"Projeto_{numero_projeto}_[A-Z]{{2,6}}",
            p.name,
            flags=re.IGNORECASE
        )
    ]

    if len(candidatas) == 1:
        return candidatas[0], "pasta com código encontrada por fallback"

    if len(candidatas) > 1:
        return None, f"existem várias pastas possíveis para Projeto_{numero_projeto}"

    return None, f"não existe a pasta Projeto_{numero_projeto}"


def analisar_ficheiro(nome_ficheiro: str) -> dict:
    """
    Analisa o nome do ficheiro e decide:
    - número de projeto;
    - tipo de documento;
    - código de município;
    - subpasta de destino.

    Aceita:
    025_AP_REL_IA_PRT.pdf
    047a_AP_REL_IA_PRT.pdf
    047b_AP_REL_IA_PRT.pdf
    042_AP_falta.txt
    """

    caminho = Path(nome_ficheiro)
    stem = caminho.stem
    extensao = caminho.suffix.lower()

    partes = stem.split("_")

    # --------------------------------------------------------
    # A) Ficheiros de falta: 042_AP_falta.txt / 010_HT_falta.txt
    # --------------------------------------------------------

    if extensao == ".txt" and len(partes) >= 3 and partes[-1].lower() == "falta":
        numero_projeto = extrair_numero_base(partes[0])
        empresa = partes[1].upper() if len(partes) >= 2 else ""

        if not numero_projeto:
            return {
                "valido": False,
                "numero_projeto": "",
                "empresa": empresa,
                "tipo_documento": "",
                "codigo_municipio": "",
                "subpasta_destino": "",
                "motivo": "não foi possível extrair número de projeto do ficheiro falta"
            }

        return {
            "valido": True,
            "numero_projeto": numero_projeto,
            "empresa": empresa,
            "tipo_documento": "FALTA",
            "codigo_municipio": "",
            "subpasta_destino": "REL_pdf",
            "motivo": "ficheiro falta"
        }

    # --------------------------------------------------------
    # B) Ficheiros normalizados:
    # [nº projeto]_[empresa]_[tipo documento]_[tipo intervenção]_[código município]
    # --------------------------------------------------------

    if len(partes) < 5:
        return {
            "valido": False,
            "numero_projeto": "",
            "empresa": "",
            "tipo_documento": "",
            "codigo_municipio": "",
            "subpasta_destino": "",
            "motivo": "nome não tem estrutura suficiente"
        }

    numero_projeto = extrair_numero_base(partes[0])
    empresa = partes[1].upper()
    tipo_documento = normalizar_tipo_documento(partes[2])
    codigo_municipio = partes[-1].upper()

    if not numero_projeto:
        return {
            "valido": False,
            "numero_projeto": "",
            "empresa": empresa,
            "tipo_documento": tipo_documento,
            "codigo_municipio": codigo_municipio,
            "subpasta_destino": "",
            "motivo": "não foi possível extrair número de projeto"
        }

    if tipo_documento not in MAPA_TIPO_DOCUMENTO:
        return {
            "valido": False,
            "numero_projeto": numero_projeto,
            "empresa": empresa,
            "tipo_documento": tipo_documento,
            "codigo_municipio": codigo_municipio,
            "subpasta_destino": "",
            "motivo": f"tipo de documento não reconhecido: {tipo_documento}"
        }

    subpasta_destino = MAPA_TIPO_DOCUMENTO[tipo_documento]

    return {
        "valido": True,
        "numero_projeto": numero_projeto,
        "empresa": empresa,
        "tipo_documento": tipo_documento,
        "codigo_municipio": codigo_municipio,
        "subpasta_destino": subpasta_destino,
        "motivo": "OK"
    }


def listar_ficheiros_origem() -> list[Path]:
    """
    Lista apenas ficheiros diretamente dentro de 14_REL_AP_pdf.
    Não entra em subpastas.
    """

    if not PASTA_ORIGEM.exists():
        raise FileNotFoundError(f"Pasta de origem não encontrada: {PASTA_ORIGEM}")

    return sorted([
        p for p in PASTA_ORIGEM.iterdir()
        if p.is_file()
    ])


# ============================================================
# 4. PROCESSO PRINCIPAL
# ============================================================

def main():
    if not PASTA_BASE.exists():
        raise FileNotFoundError(f"Pasta base não encontrada: {PASTA_BASE}")

    ficheiros = listar_ficheiros_origem()

    print(f"Pasta origem: {PASTA_ORIGEM}")
    print(f"Ficheiros encontrados diretamente na origem: {len(ficheiros)}")
    print(f"Modo simulação: {MODO_SIMULACAO}")
    print()

    resultados = []

    movidos = 0
    bloqueados = 0
    simulados = 0

    for ficheiro in ficheiros:
        info = analisar_ficheiro(ficheiro.name)

        if not info["valido"]:
            resultados.append({
                "nome_ficheiro": ficheiro.name,
                "caminho_origem": str(ficheiro),
                **info,
                "pasta_projeto": "",
                "caminho_destino": "",
                "estado": "BLOQUEADO",
                "acao": "não movido",
                "observacoes": info["motivo"],
            })
            bloqueados += 1
            print(f"bloqueado: {ficheiro.name} | {info['motivo']}")
            continue

        pasta_projeto, estado_pasta = encontrar_pasta_projeto(
            numero_projeto=info["numero_projeto"],
            codigo_municipio=info["codigo_municipio"]
        )

        if pasta_projeto is None:
            resultados.append({
                "nome_ficheiro": ficheiro.name,
                "caminho_origem": str(ficheiro),
                **info,
                "pasta_projeto": "",
                "caminho_destino": "",
                "estado": "BLOQUEADO",
                "acao": "não movido",
                "observacoes": estado_pasta,
            })
            bloqueados += 1
            print(f"bloqueado: {ficheiro.name} | {estado_pasta}")
            continue

        pasta_destino = pasta_projeto / info["subpasta_destino"]

        if not pasta_destino.exists():
            if CRIAR_SUBPASTA_DESTINO_SE_NAO_EXISTIR:
                if not MODO_SIMULACAO:
                    pasta_destino.mkdir(parents=True, exist_ok=True)
            else:
                resultados.append({
                    "nome_ficheiro": ficheiro.name,
                    "caminho_origem": str(ficheiro),
                    **info,
                    "pasta_projeto": str(pasta_projeto),
                    "caminho_destino": "",
                    "estado": "BLOQUEADO",
                    "acao": "não movido",
                    "observacoes": f"subpasta de destino não existe: {info['subpasta_destino']}",
                })
                bloqueados += 1
                print(f"bloqueado: {ficheiro.name} | subpasta destino não existe")
                continue

        destino = pasta_destino / ficheiro.name

        if destino.exists():
            resultados.append({
                "nome_ficheiro": ficheiro.name,
                "caminho_origem": str(ficheiro),
                **info,
                "pasta_projeto": str(pasta_projeto),
                "caminho_destino": str(destino),
                "estado": "BLOQUEADO",
                "acao": "não movido",
                "observacoes": "já existe ficheiro com o mesmo nome no destino",
            })
            bloqueados += 1
            print(f"bloqueado: {ficheiro.name} | já existe no destino")
            continue

        if MODO_SIMULACAO:
            resultados.append({
                "nome_ficheiro": ficheiro.name,
                "caminho_origem": str(ficheiro),
                **info,
                "pasta_projeto": str(pasta_projeto),
                "caminho_destino": str(destino),
                "estado": "PRONTO",
                "acao": "simulação: seria movido",
                "observacoes": estado_pasta,
            })
            simulados += 1
            print(f"simulação: {ficheiro.name} -> {pasta_projeto.name}\\{info['subpasta_destino']}")
        else:
            try:
                shutil.move(str(ficheiro), str(destino))

                resultados.append({
                    "nome_ficheiro": ficheiro.name,
                    "caminho_origem": str(ficheiro),
                    **info,
                    "pasta_projeto": str(pasta_projeto),
                    "caminho_destino": str(destino),
                    "estado": "MOVIDO",
                    "acao": "movido",
                    "observacoes": estado_pasta,
                })

                movidos += 1
                print(f"movido: {ficheiro.name} -> {pasta_projeto.name}\\{info['subpasta_destino']}")

            except Exception as erro:
                resultados.append({
                    "nome_ficheiro": ficheiro.name,
                    "caminho_origem": str(ficheiro),
                    **info,
                    "pasta_projeto": str(pasta_projeto),
                    "caminho_destino": str(destino),
                    "estado": "ERRO",
                    "acao": "erro",
                    "observacoes": str(erro),
                })

                bloqueados += 1
                print(f"erro: {ficheiro.name} | {erro}")

    df_resultados = pd.DataFrame(resultados)

    resumo = pd.DataFrame([
        {"Indicador": "Pasta base", "Valor": str(PASTA_BASE)},
        {"Indicador": "Pasta origem", "Valor": str(PASTA_ORIGEM)},
        {"Indicador": "Modo simulação", "Valor": MODO_SIMULACAO},
        {"Indicador": "Ficheiros encontrados na origem", "Valor": len(ficheiros)},
        {"Indicador": "Ficheiros simulados como prontos", "Valor": simulados},
        {"Indicador": "Ficheiros movidos", "Valor": movidos},
        {"Indicador": "Ficheiros bloqueados/erro", "Valor": bloqueados},
        {"Indicador": "Criar subpasta destino se não existir", "Valor": CRIAR_SUBPASTA_DESTINO_SE_NAO_EXISTIR},
    ])

    mapa = pd.DataFrame([
        {"Tipo_documento": tipo, "Subpasta_destino": subpasta}
        for tipo, subpasta in MAPA_TIPO_DOCUMENTO.items()
    ])

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        df_resultados.to_excel(writer, index=False, sheet_name="Movimentacao")
        mapa.to_excel(writer, index=False, sheet_name="Mapa_tipos")
        resumo.to_excel(writer, index=False, sheet_name="Resumo")

    print("\nProcesso concluído.")
    print(f"Relatório criado em: {RELATORIO_SAIDA}")

    if MODO_SIMULACAO:
        print("Simulação ativa: nenhum ficheiro foi movido.")
        print("Se estiver tudo correto, muda MODO_SIMULACAO para False.")
    else:
        print(f"Ficheiros movidos: {movidos}")
        print(f"Ficheiros bloqueados/erro: {bloqueados}")


if __name__ == "__main__":
    main()