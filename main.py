"""
Classificação de Manutenção x Não Manutenção - Razão Contábil de Frota
========================================================================

Este script lê o razão contábil (contas de Serviços/Peças para Logística e
Frota) e classifica cada transação em "Manutenção" ou "Não Manutenção",
além de identificar a coluna ATIVIDADE, combinando duas estratégias:

1. PROCV (merge) contra uma base de referência de transações já classificadas
   — preenche tanto MANUTENCAO quanto ATIVIDADE.
2. Regras de palavra-chave (dicionário), aplicadas como camada complementar
   e prioritária sobre o resultado do PROCV para a coluna MANUTENCAO.

Para a coluna TIPO_ATIVIDADE (atividade), a prioridade é diferente:
  a) usa o valor vindo do PROCV, quando encontrado;
  b) se o PROCV não encontrou, mas a transação bateu com alguma
     palavra-chave, marca como "Classificação inferida | palavra-chave"
     (a palavra-chave só indica Manutenção Sim/Não, não a atividade exata);
  c) se nem o PROCV nem a palavra-chave encontraram nada, marca como
     "Classificação inferida | não encontrado".

Ao final, gera um resumo mensal (valores e percentuais) e exporta o
resultado para Excel.
"""

import re
import unicodedata

import pandas as pd

from config import REGRAS_PALAVRA_CHAVE

# Configuração de caminhos

CAMINHO_RAZAO = "data/razao.xlsx"
ABA_RAZAO = "2026"
CAMINHO_BASE = "data/Transacoes_base.xlsx"
CAMINHO_SAIDA = "output/espurgos_2026.xlsx"

CONTAS_FILTRO = (
    "Servicos para Logistica e Frota"
    "|Pecas e Acessorios"
    "|Pecas e Acessorios para Logistica e Frota"
)

ATIVIDADE_INFERIDA_PALAVRA_CHAVE = "Classificação inferida | palavra-chave"
ATIVIDADE_INFERIDA_NAO_ENCONTRADO = "Classificação inferida | não encontrado"


# Limpeza de texto

def limpa_texto(text):
    "Remove acentos, números e caracteres especiais; retorna texto em maiúsculas."
    if pd.isna(text):
        return text
    text = str(text)
    text = unicodedata.normalize("NFKD", text).encode("ASCII", errors="ignore").decode("ASCII")
    text = re.sub(r"[0-9\-]", " ", text)          # retirar números e hífens
    text = re.sub(r"[^A-Za-z\s]", " ", text)       # retirar outros caracteres especiais
    text = re.sub(r"\s+", " ", text).strip()       # compactar espaços
    text = re.sub(r"\bS\s+NF\b", "", text)         # retirar o termo "S NF"
    return text.upper()



# PROCV (merge)

def _normaliza_nome_coluna(nome):
    """Remove espaços nas pontas para comparar nomes de coluna com mais tolerância."""
    return str(nome).strip()


def _encontra_coluna(df, nome_esperado):
    """
    Localiza a coluna `nome_esperado` em `df` tolerando espaços extras nas
    pontas e diferenças de maiúsculas/minúsculas. Retorna o nome EXATO da
    coluna como está em `df`, ou None se não encontrar.
    """
    alvo = _normaliza_nome_coluna(nome_esperado).lower()
    for coluna in df.columns:
        if _normaliza_nome_coluna(coluna).lower() == alvo:
            return coluna
    return None


def preenche_manutencao_procv(
    razao,
    base,
    razao_texto_col="TEXTO_MANUTENCAO",
    base_trans_col="TRANSACAO",
    manut_col="MANUTENCAO",
    atividade_col="ATIVIDADE",
):
    """
    Preenche as colunas MANUTENCAO e TIPO_ATIVIDADE em `razao`, cruzando o
    texto limpo da transação com a base de referência (`base`).

    A base precisa ter as colunas `base_trans_col`, `manut_col` e
    `atividade_col`. O resultado do cruzamento de atividade é gravado na
    coluna `TIPO_ATIVIDADE` de `razao`.
    """
    base_copy = base.copy()

    # localiza as colunas tolerando espaço extra / maiúscula-minúscula, e já
    # troca para o nome exato encontrado em base_copy
    colunas_encontradas = {}
    for nome_esperado in (base_trans_col, manut_col, atividade_col):
        encontrada = _encontra_coluna(base_copy, nome_esperado)
        if encontrada is None:
            raise KeyError(
                f"Coluna '{nome_esperado}' não encontrada em base. "
                f"Colunas disponíveis: {list(base_copy.columns)}"
            )
        colunas_encontradas[nome_esperado] = encontrada

    base_trans_col = colunas_encontradas[base_trans_col]
    manut_col = colunas_encontradas[manut_col]
    atividade_col = colunas_encontradas[atividade_col]

    base_copy[base_trans_col] = base_copy[base_trans_col].astype(str).apply(limpa_texto)

    resultado = razao.copy()
    resultado[razao_texto_col] = resultado[razao_texto_col].astype(str).apply(limpa_texto)

    base_unique = base_copy[
        [base_trans_col, manut_col, atividade_col]
    ].drop_duplicates(subset=[base_trans_col])

    resultado = resultado.merge(
        base_unique,
        left_on=razao_texto_col,
        right_on=base_trans_col,
        how="left",
    )

    resultado["MANUTENCAO"] = resultado[manut_col]
    resultado["TIPO_ATIVIDADE"] = resultado[atividade_col]

    # cuidado: manut_col já é o nome final "MANUTENCAO" por padrão, então só
    # descartamos as colunas auxiliares que não devem sobrar no resultado.
    colunas_para_descartar = [base_trans_col]
    if atividade_col != "TIPO_ATIVIDADE":
        colunas_para_descartar.append(atividade_col)
    if manut_col != "MANUTENCAO":
        colunas_para_descartar.append(manut_col)

    resultado = resultado.drop(columns=colunas_para_descartar)
    return resultado


# Classificação por palavra-chave

def classifica_por_palavra_chave(
    razao,
    regras=REGRAS_PALAVRA_CHAVE,
    texto_col="TEXTO_MANUTENCAO",
    manut_col="MANUTENCAO",
    prioridade=True,
):
    """
    Aplica regras de classificação por palavra-chave sobre a coluna de texto limpo.

    - regras: dict {palavra_chave: classificacao}, busca case-insensitive por
      substring dentro do texto limpo (TEXTO_MANUTENCAO).
    - prioridade=True: a regra de palavra-chave sobrescreve o resultado do PROCV
      sempre que a palavra-chave for encontrada (útil para exceções conhecidas).
    - prioridade=False: a regra só é aplicada nos registros ainda sem
      classificação (MANUTENCAO nulo/NaN) após o PROCV.
    """
    resultado = razao.copy()
    texto_lower = resultado[texto_col].astype(str).str.lower()

    for palavra, classificacao in regras.items():
        contem_palavra = texto_lower.str.contains(palavra.lower(), na=False, regex=False)
        alvo = contem_palavra if prioridade else (contem_palavra & resultado[manut_col].isna())
        resultado.loc[alvo, manut_col] = classificacao

    return resultado


def mask_bate_palavra_chave(razao, regras=REGRAS_PALAVRA_CHAVE, texto_col="TEXTO_MANUTENCAO"):
    """Retorna uma Series booleana: True onde o texto bate com alguma palavra-chave."""
    texto_lower = razao[texto_col].astype(str).str.lower()
    bate = pd.Series(False, index=razao.index)
    for palavra in regras:
        bate = bate | texto_lower.str.contains(palavra.lower(), na=False, regex=False)
    return bate


def preenche_atividade_fallback(
    razao,
    regras=REGRAS_PALAVRA_CHAVE,
    texto_col="TEXTO_MANUTENCAO",
    atividade_col="TIPO_ATIVIDADE",
):
    """
    Preenche TIPO_ATIVIDADE para as linhas que o PROCV não conseguiu classificar.

    - Se a transação bate com alguma palavra-chave (ou seja, MANUTENCAO já foi
      inferida por palavra-chave), marca a atividade como
      "Classificação inferida | palavra-chave" — a palavra-chave indica se é
      manutenção ou não, mas não sabe qual é a atividade exata.
    - Se nem o PROCV nem a palavra-chave encontraram nada, marca como
      "Classificação inferida | não encontrado".
    """
    resultado = razao.copy()
    sem_atividade = resultado[atividade_col].isna()
    bate_palavra_chave = mask_bate_palavra_chave(resultado, regras, texto_col)

    resultado.loc[sem_atividade & bate_palavra_chave, atividade_col] = (
        ATIVIDADE_INFERIDA_PALAVRA_CHAVE
    )
    resultado.loc[sem_atividade & ~bate_palavra_chave, atividade_col] = (
        ATIVIDADE_INFERIDA_NAO_ENCONTRADO
    )

    return resultado

# Resumo mensal

def gera_resumo_mensal(razao):
    resumo = (
        razao.groupby("MES")
        .agg(
            Nao_Manutencao=("VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"] == "Não"].sum()),
            Manutencao=("VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"] == "Sim"].sum()),
            Nan=("VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"].isna()].sum()),
            Valor_Total=("VALOR_MOEDA", "sum"),
        )
        .reset_index()
    )

    percentual = (
        razao.groupby("MES")
        .agg(
            Nao_manutencao_percentual=(
                "VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"] == "Não"].sum() / x.sum() * 100
            ),
            Manutencao_percentual=(
                "VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"] == "Sim"].sum() / x.sum() * 100
            ),
            Nan_percentual=(
                "VALOR_MOEDA", lambda x: x[razao["MANUTENCAO"].isna()].sum() / x.sum() * 100
            ),
            total_percentual=("VALOR_MOEDA", lambda x: x.sum() / x.sum() * 100),
        )
        .reset_index()
    )

    return resumo, percentual

# Fluxo principal

def main():
    razao = pd.read_excel(CAMINHO_RAZAO, sheet_name=ABA_RAZAO)
    base = pd.read_excel(CAMINHO_BASE)

    total_geral = razao["VALOR_MOEDA"].sum()

    razao = razao[
        ["COD_ESTABELECIMENTO", "DATA_TRANSACAO", "NOME_CONTA", "TEXTO_TRANSACAO", "VALOR_MOEDA"]
    ]
    razao["TEXTO_MANUTENCAO"] = razao["TEXTO_TRANSACAO"]
    razao["TIPO_ATIVIDADE"] = None
    razao["VALOR_MOEDA"] = razao["VALOR_MOEDA"].astype("Float64")
    razao["DATA_TRANSACAO"] = pd.to_datetime(razao["DATA_TRANSACAO"], errors="coerce")
    razao["MES"] = razao["DATA_TRANSACAO"].dt.strftime("%Y-%m")

    razao = razao[razao["NOME_CONTA"].str.contains(CONTAS_FILTRO, case=False, na=False)]
    total_filtrado = razao["VALOR_MOEDA"].sum()

    razao["TEXTO_MANUTENCAO"] = razao["TEXTO_MANUTENCAO"].apply(limpa_texto)
    base["TRANSACAO"] = base["TRANSACAO"].apply(limpa_texto)

    razao = preenche_manutencao_procv(razao, base)
    razao = preenche_atividade_fallback(razao)  # completa TIPO_ATIVIDADE onde o PROCV não achou nada
    razao = classifica_por_palavra_chave(razao, prioridade=True)  # ajusta MANUTENCAO (Sim/Não)

    resumo, percentual = gera_resumo_mensal(razao)

    print(f"Valor total do razão (todas as contas): {total_geral:,.2f}")
    print(f"Valor total após filtro de contas de manutenção: {total_filtrado:,.2f}\n")
    print(resumo)
    print(percentual)

    razao.to_excel(CAMINHO_SAIDA, index=False)
    print(f"\nArquivo exportado para: {CAMINHO_SAIDA}")


if __name__ == "__main__":
    main()