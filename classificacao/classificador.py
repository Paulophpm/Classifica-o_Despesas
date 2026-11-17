"""
Preparação do razão e classificação em camadas.

Cada lançamento passa por camadas ordenadas por PRIORIDADE (menor número
primeiro). A primeira camada que reconhecer o lançamento define a CATEGORIA;
as demais não sobrescrevem. As camadas são:

  - Exceções manuais (aba EXCECOES)            -> PRIORIDADE_EXCECAO (0)
  - Regras de palavra-chave (aba REGRAS)       -> prioridade de cada regra
  - Base PROCV com ATIVIDADE mapeada            -> PRIORIDADE_BASE_ATIVIDADE (50)
  - Base PROCV só com Sim/Não                   -> PRIORIDADE_BASE_SIM_NAO (500)
  - Pessoa física detectada                    -> PRIORIDADE_PF (990)
  - Categoria padrão da conta (aba CONTAS)     -> PRIORIDADE_PADRAO_CONTA (995)

Depois, a categoria é convertida no GRUPO_ESPERADO e comparada ao
GRUPO_ATUAL (grupo da conta onde o lançamento caiu) para gerar o STATUS.
"""

import re

import numpy as np
import pandas as pd

from .regras import (
    CATEGORIA_BASE_NAO,
    CATEGORIA_BASE_SIM,
    CATEGORIA_PF,
    GRUPO_AJUSTE,
    GRUPO_NAO_MANUT,
    GRUPO_VERIFICAR,
    Regras,
)
from .texto import limpa_texto_procv, limpa_texto_regra, normaliza_chave

COLUNAS_OBRIGATORIAS = ["DATA_TRANSACAO", "NOME_CONTA", "TEXTO_TRANSACAO", "VALOR_MOEDA"]


# ---------------------------------------------------------------------------
# Preparação
# ---------------------------------------------------------------------------

def _para_numero(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(serie):
        return serie.astype(float)
    texto = serie.astype(str).str.strip()
    # formato brasileiro "1.234,56" -> "1234.56"
    br = texto.str.contains(",", regex=False)
    texto = texto.where(~br, texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(texto, errors="coerce")


def _para_data(serie: pd.Series) -> pd.Series:
    """Aceita datas do Excel, texto ISO (2026-10-05) e texto brasileiro (05/10/2026)."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    texto = serie.astype(str).str.strip()
    iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}")
    datas = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    datas[iso] = pd.to_datetime(texto[iso].str[:10], errors="coerce", format="%Y-%m-%d")
    datas[~iso] = pd.to_datetime(texto[~iso], errors="coerce", dayfirst=True)
    return datas


def prepara_razao(razao: pd.DataFrame, regras: Regras):
    """
    Retorna (analise, fora_do_escopo, excluidos_premissa):
      - analise: lançamentos das contas cadastradas, prontos para classificar
      - fora_do_escopo: lançamentos de contas que não estão na aba CONTAS
      - excluidos_premissa: lançamentos removidos por premissa (tipo de lançamento etc.)
    """
    razao = razao.copy()
    razao.columns = [str(c).strip() for c in razao.columns]
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in razao.columns]
    if faltando:
        raise KeyError(f"O razão não tem as colunas obrigatórias {faltando}. Colunas lidas: {list(razao.columns)}")

    campos_texto = regras.premissa_lista("CAMPOS_TEXTO") or ["TEXTO_TRANSACAO"]
    faltando = [c for c in campos_texto if c not in razao.columns]
    if faltando:
        raise KeyError(f"Premissa CAMPOS_TEXTO cita colunas que não existem no razão: {faltando}")

    razao["VALOR_MOEDA"] = _para_numero(razao["VALOR_MOEDA"])
    razao["DATA_TRANSACAO"] = _para_data(razao["DATA_TRANSACAO"])
    razao["MES"] = razao["DATA_TRANSACAO"].dt.strftime("%Y-%m")

    razao["CHAVE_CONTA"] = razao["NOME_CONTA"].map(normaliza_chave)
    contas = regras.contas.rename(columns={"NOME_CONTA": "CONTA_CADASTRO"})
    razao = razao.merge(contas, on="CHAVE_CONTA", how="left")

    dentro = razao["GRUPO_ATUAL"].notna()
    fora_do_escopo = razao[~dentro].copy()
    analise = razao[dentro].copy()

    # exclusões por premissa
    excluir = pd.Series(False, index=analise.index)
    for premissa, coluna in (("EXCLUIR_TIPO_LANCAMENTO", "TIPO_LANCAMENTO"), ("EXCLUIR_TIPO_TRANSACAO", "TIPO_TRANSACAO")):
        valores = {normaliza_chave(v) for v in regras.premissa_lista(premissa)}
        if valores and coluna in analise.columns:
            excluir |= analise[coluna].map(normaliza_chave).isin(valores)
    excluidos = analise[excluir].copy()
    analise = analise[~excluir].copy()

    analise["TEXTO_ANALISE"] = (
        analise[campos_texto].fillna("").astype(str).agg(" ".join, axis=1).str.strip()
    )
    analise["TEXTO_REGRA"] = analise["TEXTO_ANALISE"].map(limpa_texto_regra)
    analise["TEXTO_PROCV"] = analise["TEXTO_TRANSACAO"].map(limpa_texto_procv)
    return analise.reset_index(drop=True), fora_do_escopo, excluidos


# ---------------------------------------------------------------------------
# Base de referência (PROCV)
# ---------------------------------------------------------------------------

def _acha_coluna(df, nome):
    alvo = nome.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == alvo:
            return c
    return None


def _normaliza_sim_nao(v):
    chave = normaliza_chave(v)
    if chave in {"sim", "s"}:
        return "Sim"
    if chave in {"nao", "n"}:
        return "Não"
    return None


def prepara_base(base: pd.DataFrame | None):
    """
    Retorna (lookup, conflitos):
      - lookup: DataFrame indexado pelo texto limpo, com BASE_MANUTENCAO e BASE_ATIVIDADE
      - conflitos: textos que aparecem na base com classificações diferentes
    """
    vazio = pd.DataFrame(columns=["BASE_MANUTENCAO", "BASE_ATIVIDADE"])
    if base is None or base.empty:
        return vazio, pd.DataFrame()

    col_trans = _acha_coluna(base, "TRANSACAO")
    col_manut = _acha_coluna(base, "MANUTENCAO")
    col_atv = _acha_coluna(base, "ATIVIDADE")
    if col_trans is None or col_manut is None:
        raise KeyError(f"A base precisa das colunas TRANSACAO e MANUTENCAO. Colunas lidas: {list(base.columns)}")

    b = pd.DataFrame({
        "TEXTO_PROCV": base[col_trans].map(limpa_texto_procv),
        "BASE_MANUTENCAO": base[col_manut].map(_normaliza_sim_nao),
        "BASE_ATIVIDADE": base[col_atv] if col_atv is not None else None,
    })
    b["BASE_ATIVIDADE"] = b["BASE_ATIVIDADE"].map(lambda v: None if pd.isna(v) or str(v).strip() == "" else str(v).strip())
    b = b[b["TEXTO_PROCV"].notna() & (b["TEXTO_PROCV"] != "")]

    resumo = b.groupby("TEXTO_PROCV").agg(
        QTD_NA_BASE=("BASE_MANUTENCAO", "size"),
        VALORES_MANUTENCAO=("BASE_MANUTENCAO", lambda s: " | ".join(sorted({str(x) for x in s.dropna()}))),
        VALORES_ATIVIDADE=("BASE_ATIVIDADE", lambda s: " | ".join(sorted({str(x) for x in s.dropna()}))),
        N_MANUT=("BASE_MANUTENCAO", "nunique"),
        N_ATV=("BASE_ATIVIDADE", "nunique"),
    )
    conflitos = resumo[(resumo["N_MANUT"] > 1) | (resumo["N_ATV"] > 1)].drop(columns=["N_MANUT", "N_ATV"])
    conflitos = conflitos.reset_index().rename(columns={"TEXTO_PROCV": "TEXTO_LIMPO_BASE"})

    # mesmo critério do código original: fica a primeira ocorrência
    lookup = b.drop_duplicates(subset="TEXTO_PROCV", keep="first").set_index("TEXTO_PROCV")
    return lookup[["BASE_MANUTENCAO", "BASE_ATIVIDADE"]], conflitos


# ---------------------------------------------------------------------------
# Pessoa física
# ---------------------------------------------------------------------------

def detecta_pf(textos: pd.Series, regras: Regras) -> pd.Series:
    """
    Heurística: o texto tem um NOME e um SOBRENOME conhecidos (ou dois
    sobrenomes) e nenhum termo de empresa (LTDA, ME, COMERCIO...). As listas
    ficam na aba PESSOA_FISICA.
    """
    if not regras.pf_nomes and not regras.pf_sobrenomes:
        return pd.Series(False, index=textos.index)

    padrao_empresa = None
    if regras.pf_empresa:
        alternativas = "|".join(re.escape(t) for t in sorted(regras.pf_empresa, key=len, reverse=True))
        padrao_empresa = re.compile(rf"(?<![A-Z0-9])(?:{alternativas})(?![A-Z0-9])")

    nomes, sobrenomes = regras.pf_nomes, regras.pf_sobrenomes

    def eh_pf(texto: str) -> bool:
        if not texto:
            return False
        if padrao_empresa is not None and padrao_empresa.search(texto):
            return False
        tokens = texto.split()
        tem_nome = any(t in nomes for t in tokens)
        qtd_sobrenome = sum(t in sobrenomes for t in tokens)
        return (tem_nome and qtd_sobrenome >= 1) or qtd_sobrenome >= 2

    return textos.map(eh_pf)


# ---------------------------------------------------------------------------
# Classificação
# ---------------------------------------------------------------------------

def _padrao_regex(termo: str, tipo: str) -> str:
    if tipo == "regex":
        return termo
    escapado = re.escape(termo)
    if tipo == "palavra":
        return rf"(?<![A-Z0-9]){escapado}(?![A-Z0-9])"
    if tipo == "prefixo":
        return rf"(?<![A-Z0-9]){escapado}"
    return escapado  # contem


class _Atribuidor:
    """Guarda a primeira categoria atribuída a cada linha."""

    def __init__(self, n):
        self.categoria = np.full(n, None, dtype=object)
        self.origem = np.full(n, None, dtype=object)
        self.detalhe = np.full(n, None, dtype=object)
        self.prioridade = np.full(n, np.nan)

    def pendentes(self) -> np.ndarray:
        return pd.isna(self.categoria)

    def atribui(self, mask, categoria, origem, detalhe, prioridade):
        mask = np.asarray(mask, dtype=bool) & self.pendentes()
        if not mask.any():
            return
        self.categoria[mask] = categoria[mask] if isinstance(categoria, np.ndarray) else categoria
        self.origem[mask] = origem
        self.detalhe[mask] = detalhe[mask] if isinstance(detalhe, np.ndarray) else detalhe
        self.prioridade[mask] = prioridade


def classifica(analise: pd.DataFrame, regras: Regras, base_lookup: pd.DataFrame) -> pd.DataFrame:
    df = analise.copy().reset_index(drop=True)
    n = len(df)
    atr = _Atribuidor(n)
    texto = df["TEXTO_REGRA"].fillna("")
    grupo_atual_norm = df["GRUPO_ATUAL"].map(normaliza_chave)

    # dados da base (para a camada PROCV e para auditoria)
    if len(base_lookup):
        juntou = df[["TEXTO_PROCV"]].join(base_lookup, on="TEXTO_PROCV")
        df["BASE_MANUTENCAO"] = juntou["BASE_MANUTENCAO"].to_numpy()
        df["BASE_ATIVIDADE"] = juntou["BASE_ATIVIDADE"].to_numpy()
    else:
        df["BASE_MANUTENCAO"] = None
        df["BASE_ATIVIDADE"] = None

    df["POSSIVEL_PF"] = detecta_pf(texto, regras).to_numpy() if regras.premissa_sim("DETECTAR_PF") else False

    camadas = []  # (prioridade, ordem, função)

    # Exceções manuais: texto + conta específica tem preferência sobre texto sem conta
    if len(regras.excecoes):
        def camada_excecoes():
            exc = regras.excecoes
            especificas = {(t, c): cat for t, c, cat in exc[exc["CHAVE_CONTA"] != ""][["TEXTO_REGRA", "CHAVE_CONTA", "CATEGORIA"]].itertuples(index=False)}
            gerais = {t: cat for t, cat in exc[exc["CHAVE_CONTA"] == ""][["TEXTO_REGRA", "CATEGORIA"]].itertuples(index=False)}
            cats = np.array([
                especificas.get((t, c), gerais.get(t)) for t, c in zip(texto, df["CHAVE_CONTA"])
            ], dtype=object)
            atr.atribui(pd.notna(cats), cats, "Exceção manual", "EXCECOES", regras.premissa_int("PRIORIDADE_EXCECAO"))
        camadas.append((regras.premissa_int("PRIORIDADE_EXCECAO"), 0, camada_excecoes))

    # Regras de palavra-chave
    for ordem, regra in enumerate(regras.regras.itertuples(index=False), start=1):
        def camada_regra(regra=regra):
            pend = atr.pendentes()
            if not pend.any():
                return
            alvo = pend.copy()
            if regra.SOMENTE_GRUPO_ATUAL:
                permitidos = {normaliza_chave(g) for g in regra.SOMENTE_GRUPO_ATUAL}
                alvo &= grupo_atual_norm.isin(permitidos).to_numpy()
                if not alvo.any():
                    return
            flags = re.IGNORECASE if regra.TIPO_MATCH == "regex" else 0
            bate = np.zeros(n, dtype=bool)
            bate[alvo] = texto[alvo].str.contains(_padrao_regex(regra.TERMO, regra.TIPO_MATCH), flags=flags, regex=True, na=False).to_numpy()
            detalhe = f"{regra.TIPO_MATCH}: {regra.TERMO_ORIGINAL}"
            atr.atribui(bate, regra.CATEGORIA, "Regra palavra-chave", detalhe, regra.PRIORIDADE)
        camadas.append((float(regra.PRIORIDADE), ordem, camada_regra))

    # Base PROCV
    base_manut = df["BASE_MANUTENCAO"].to_numpy(dtype=object)
    base_atv = df["BASE_ATIVIDADE"].to_numpy(dtype=object)

    def camada_base_atividade():
        cats = np.array([regras.mapa_atividade.get(normaliza_chave(a)) if pd.notna(a) else None for a in base_atv], dtype=object)
        detalhe = np.array([f"ATIVIDADE: {a}" for a in base_atv], dtype=object)
        atr.atribui(pd.notna(cats), cats, "Base PROCV (atividade)", detalhe, regras.premissa_int("PRIORIDADE_BASE_ATIVIDADE"))

    def camada_base_sim_nao():
        cats = np.array([
            CATEGORIA_BASE_SIM if m == "Sim" else CATEGORIA_BASE_NAO if m == "Não" else None for m in base_manut
        ], dtype=object)
        detalhe = np.array([f"MANUTENCAO: {m}" for m in base_manut], dtype=object)
        atr.atribui(pd.notna(cats), cats, "Base PROCV (Sim/Não)", detalhe, regras.premissa_int("PRIORIDADE_BASE_SIM_NAO"))

    camadas.append((regras.premissa_int("PRIORIDADE_BASE_ATIVIDADE"), 10**6, camada_base_atividade))
    camadas.append((regras.premissa_int("PRIORIDADE_BASE_SIM_NAO"), 10**6, camada_base_sim_nao))

    # Pessoa física
    if regras.premissa_sim("DETECTAR_PF"):
        def camada_pf():
            atr.atribui(df["POSSIVEL_PF"].to_numpy(dtype=bool), CATEGORIA_PF, "Pessoa física (heurística)", "PESSOA_FISICA", regras.premissa_int("PRIORIDADE_PF"))
        camadas.append((regras.premissa_int("PRIORIDADE_PF"), 10**6, camada_pf))

    # Categoria padrão da conta
    if regras.premissa_sim("APLICAR_PADRAO_CONTA"):
        def camada_padrao():
            cats = df["CATEGORIA_PADRAO"].map(lambda v: v if isinstance(v, str) and v else None).to_numpy(dtype=object)
            atr.atribui(pd.notna(cats), cats, "Padrão da conta", "CONTAS", regras.premissa_int("PRIORIDADE_PADRAO_CONTA"))
        camadas.append((regras.premissa_int("PRIORIDADE_PADRAO_CONTA"), 10**6, camada_padrao))

    for _, _, func in sorted(camadas, key=lambda c: (c[0], c[1])):
        func()

    df["CATEGORIA"] = atr.categoria
    df["ORIGEM_CLASSIFICACAO"] = atr.origem
    df["REGRA_APLICADA"] = atr.detalhe
    df["PRIORIDADE_APLICADA"] = atr.prioridade
    return aplica_status(df, regras)


def aplica_status(df: pd.DataFrame, regras: Regras) -> pd.DataFrame:
    manut = normaliza_chave(regras.grupo_manutencao)
    comb = normaliza_chave(regras.grupo_combustivel)
    p_alta = regras.premissa_int("PRIORIDADE_BASE_ATIVIDADE")
    p_media = regras.premissa_int("PRIORIDADE_BASE_SIM_NAO")

    df["GRUPO_ESPERADO"] = df["CATEGORIA"].map(lambda c: regras.categorias.get(c) if isinstance(c, str) else None)
    atual = df["GRUPO_ATUAL"].map(normaliza_chave)
    esperado = df["GRUPO_ESPERADO"].map(normaliza_chave)

    status = np.select(
        [
            df["CATEGORIA"].isna(),
            esperado == normaliza_chave(GRUPO_AJUSTE),
            esperado == normaliza_chave(GRUPO_VERIFICAR),
            (esperado == normaliza_chave(GRUPO_NAO_MANUT)) & (atual != manut),
            esperado == normaliza_chave(GRUPO_NAO_MANUT),
            esperado == atual,
        ],
        ["Não classificado", "Ajuste contábil", "Revisar", "Revisar", "Divergente", "OK"],
        default="Divergente",
    )
    df["STATUS"] = status

    def texto_divergencia(linha):
        if linha["STATUS"] != "Divergente":
            return None
        esp = linha["GRUPO_ESPERADO"]
        if esp == GRUPO_NAO_MANUT:
            esp = "Não manutenção"
        return f"{esp} lançado em {linha['GRUPO_ATUAL']}"

    df["DIVERGENCIA"] = df.apply(texto_divergencia, axis=1)
    divergente = df["STATUS"] == "Divergente"
    df["COMBUSTIVEL_EM_MANUTENCAO"] = divergente & (atual == manut) & (esperado == comb)
    df["MANUTENCAO_FORA_DA_CONTA"] = divergente & (atual != manut) & (esperado == manut)
    df["NAO_MANUTENCAO_EM_MANUTENCAO"] = divergente & (atual == manut) & (esperado != manut)

    prio = df["PRIORIDADE_APLICADA"]
    df["CONFIANCA"] = np.select(
        [prio.isna(), prio <= p_alta, prio < p_media],
        ["", "Alta", "Média"],
        default="Baixa",
    )

    # Base PROCV discorda do resultado final? (só quando a base não foi quem decidiu)
    decidido_fora_da_base = ~df["ORIGEM_CLASSIFICACAO"].fillna("").str.startswith("Base PROCV")
    df["CONFLITO_BASE_x_REGRA"] = decidido_fora_da_base & df["CATEGORIA"].notna() & (
        ((df["BASE_MANUTENCAO"] == "Sim") & (esperado != manut) & ~esperado.isin([normaliza_chave(GRUPO_AJUSTE), normaliza_chave(GRUPO_VERIFICAR)]))
        | ((df["BASE_MANUTENCAO"] == "Não") & (esperado == manut))
    )
    return df
