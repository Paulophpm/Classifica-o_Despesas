"""Montagem das abas de saída e exportação para Excel."""

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .regras import Regras
from .texto import normaliza_chave

SEM_CLASSIFICACAO = "(não classificado)"
STATUS_ORDEM = ["OK", "Divergente", "Revisar", "Não classificado", "Ajuste contábil"]


def _soma(df, mask=None):
    sub = df if mask is None else df[mask]
    return float(sub["VALOR_MOEDA"].sum()), int(len(sub))


def _pct(valor, total):
    return valor / total if total else 0.0


# ---------------------------------------------------------------------------
# Abas
# ---------------------------------------------------------------------------

def aba_resumo(mes: pd.DataFrame, competencia: str) -> pd.DataFrame:
    total, qtd_total = _soma(mes)
    linhas = [("Total analisado (contas cadastradas)", total, qtd_total)]
    for st in STATUS_ORDEM:
        linhas.append((f"Status: {st}", *_soma(mes, mes["STATUS"] == st)))
    linhas += [
        ("Combustível lançado em Manutenção", *_soma(mes, mes["COMBUSTIVEL_EM_MANUTENCAO"])),
        ("Manutenção lançada fora da conta de Manutenção", *_soma(mes, mes["MANUTENCAO_FORA_DA_CONTA"])),
        ("Não manutenção lançada em Manutenção", *_soma(mes, mes["NAO_MANUTENCAO_EM_MANUTENCAO"])),
        ("Possível pessoa física (todos os status)", *_soma(mes, mes["POSSIVEL_PF"])),
        ("Base PROCV discorda da regra aplicada", *_soma(mes, mes["CONFLITO_BASE_x_REGRA"])),
    ]
    for conf in ["Alta", "Média", "Baixa"]:
        linhas.append((f"Confiança da classificação: {conf}", *_soma(mes, mes["CONFIANCA"] == conf)))

    df = pd.DataFrame(linhas, columns=["INDICADOR", "VALOR", "QTD_LANCAMENTOS"])
    df["%_DO_TOTAL"] = df["VALOR"].map(lambda v: _pct(v, total))
    df.insert(0, "COMPETENCIA", competencia)
    return df


def aba_resumo_por_grupo(mes: pd.DataFrame) -> pd.DataFrame:
    tab = mes.pivot_table(index="GRUPO_ATUAL", columns="STATUS", values="VALOR_MOEDA", aggfunc="sum", fill_value=0)
    tab = tab.reindex(columns=[s for s in STATUS_ORDEM if s in tab.columns])
    tab["TOTAL"] = tab.sum(axis=1)
    if "OK" in tab.columns:
        tab["%_OK"] = tab["OK"] / tab["TOTAL"].where(tab["TOTAL"] != 0)
    return tab.sort_values("TOTAL", ascending=False).reset_index()


def aba_matriz_grupos(mes: pd.DataFrame) -> pd.DataFrame:
    tmp = mes.assign(GRUPO_ESPERADO=mes["GRUPO_ESPERADO"].fillna(SEM_CLASSIFICACAO))
    tab = tmp.pivot_table(index="GRUPO_ATUAL", columns="GRUPO_ESPERADO", values="VALOR_MOEDA",
                          aggfunc="sum", fill_value=0, margins=True, margins_name="TOTAL")
    return tab.reset_index().rename(columns={"GRUPO_ATUAL": "GRUPO_ATUAL \\ GRUPO_ESPERADO"})


def aba_matriz_categorias(mes: pd.DataFrame) -> pd.DataFrame:
    tmp = mes.assign(
        GRUPO_ESPERADO=mes["GRUPO_ESPERADO"].fillna(SEM_CLASSIFICACAO),
        CATEGORIA=mes["CATEGORIA"].fillna(SEM_CLASSIFICACAO),
    )
    tab = tmp.pivot_table(index=["GRUPO_ESPERADO", "CATEGORIA"], columns="GRUPO_ATUAL", values="VALOR_MOEDA",
                          aggfunc="sum", fill_value=0, margins=True, margins_name="TOTAL")
    return tab.reset_index()


def aba_categoria_por_conta(mes: pd.DataFrame) -> pd.DataFrame:
    tmp = mes.assign(CATEGORIA=mes["CATEGORIA"].fillna(SEM_CLASSIFICACAO),
                     GRUPO_ESPERADO=mes["GRUPO_ESPERADO"].fillna(SEM_CLASSIFICACAO))
    g = tmp.groupby(["GRUPO_ATUAL", "CONTA_CADASTRO", "GRUPO_ESPERADO", "CATEGORIA", "STATUS"], dropna=False).agg(
        VALOR=("VALOR_MOEDA", "sum"), QTD=("VALOR_MOEDA", "size")
    ).reset_index()
    total_conta = g.groupby("CONTA_CADASTRO")["VALOR"].transform("sum")
    g["%_DA_CONTA"] = g["VALOR"] / total_conta.where(total_conta != 0)
    return g.sort_values(["GRUPO_ATUAL", "CONTA_CADASTRO", "VALOR"], ascending=[True, True, False])


COLUNAS_DETALHE = [
    "MES", "DATA_TRANSACAO", "COD_ESTABELECIMENTO", "COMPROVANTE", "DOCUMENTO", "CONTA_CONTABIL", "CENTRO_CUSTO",
    "NOME_CONTA", "GRUPO_ATUAL", "TEXTO_TRANSACAO", "VALOR_MOEDA", "CATEGORIA", "GRUPO_ESPERADO", "STATUS",
    "DIVERGENCIA", "CONFIANCA", "ORIGEM_CLASSIFICACAO", "REGRA_APLICADA", "POSSIVEL_PF", "BASE_MANUTENCAO",
    "BASE_ATIVIDADE", "CONFLITO_BASE_x_REGRA", "COMBUSTIVEL_EM_MANUTENCAO", "MANUTENCAO_FORA_DA_CONTA",
    "NAO_MANUTENCAO_EM_MANUTENCAO", "TEXTO_REGRA",
]


def colunas_detalhe(df: pd.DataFrame) -> pd.DataFrame:
    principais = [c for c in COLUNAS_DETALHE if c in df.columns]
    internas = {"CHAVE_CONTA", "CONTA_CADASTRO", "CATEGORIA_PADRAO", "TEXTO_ANALISE", "TEXTO_PROCV", "PRIORIDADE_APLICADA"}
    resto = [c for c in df.columns if c not in principais and c not in internas]
    return df[principais + resto]


def aba_divergencias(mes: pd.DataFrame) -> pd.DataFrame:
    div = mes[mes["STATUS"] == "Divergente"].copy()
    div["_ABS"] = div["VALOR_MOEDA"].abs()
    return colunas_detalhe(div.sort_values("_ABS", ascending=False).drop(columns="_ABS"))


def aba_pendentes(mes: pd.DataFrame, valor_minimo: float) -> pd.DataFrame:
    """
    Agrupa por texto os lançamentos sem classificação ou a revisar. As três
    primeiras colunas (TEXTO, NOME_CONTA, CATEGORIA) têm o mesmo formato da aba
    EXCECOES: basta preencher CATEGORIA e colar lá.
    """
    pend = mes[mes["STATUS"].isin(["Não classificado", "Revisar"])]
    if pend.empty:
        return pd.DataFrame(columns=["TEXTO", "NOME_CONTA", "CATEGORIA", "QTD", "VALOR", "CONTAS", "STATUS",
                                     "CATEGORIA_ATUAL", "POSSIVEL_PF", "EXEMPLO_TEXTO_ORIGINAL"])
    g = pend.groupby("TEXTO_REGRA").agg(
        QTD=("VALOR_MOEDA", "size"),
        VALOR=("VALOR_MOEDA", "sum"),
        CONTAS=("CONTA_CADASTRO", lambda s: " | ".join(sorted(set(s.dropna())))),
        STATUS=("STATUS", lambda s: " | ".join(sorted(set(s)))),
        CATEGORIA_ATUAL=("CATEGORIA", lambda s: " | ".join(sorted({str(x) for x in s.dropna()}))),
        POSSIVEL_PF=("POSSIVEL_PF", "max"),
        EXEMPLO_TEXTO_ORIGINAL=("TEXTO_TRANSACAO", "first"),
    ).reset_index()
    g = g[g["VALOR"].abs() >= valor_minimo]
    g.insert(0, "CATEGORIA", "")
    g.insert(0, "NOME_CONTA", "")
    g = g.rename(columns={"TEXTO_REGRA": "TEXTO"})
    g = g[["TEXTO", "NOME_CONTA", "CATEGORIA"] + [c for c in g.columns if c not in ("TEXTO", "NOME_CONTA", "CATEGORIA")]]
    g["_ABS"] = g["VALOR"].abs()
    return g.sort_values("_ABS", ascending=False).drop(columns="_ABS")


def _meses_ate(df, competencia, n):
    meses = sorted(m for m in df["MES"].dropna().unique() if m <= competencia)
    return meses[-n:]


def aba_comparativo_categorias(todos: pd.DataFrame, competencia: str, n: int) -> pd.DataFrame:
    meses = _meses_ate(todos, competencia, n)
    tmp = todos[todos["MES"].isin(meses)].assign(
        CATEGORIA=lambda d: d["CATEGORIA"].fillna(SEM_CLASSIFICACAO),
        GRUPO_ESPERADO=lambda d: d["GRUPO_ESPERADO"].fillna(SEM_CLASSIFICACAO),
    )
    if tmp.empty:
        return pd.DataFrame()
    tab = tmp.pivot_table(index=["GRUPO_ATUAL", "GRUPO_ESPERADO", "CATEGORIA"], columns="MES",
                          values="VALOR_MOEDA", aggfunc="sum", fill_value=0)
    tab = tab.reindex(columns=meses, fill_value=0)
    if len(meses) >= 2:
        ant, atual = meses[-2], meses[-1]
        tab["VARIACAO_MES_ANTERIOR"] = tab[atual] - tab[ant]
        tab["VARIACAO_%"] = tab["VARIACAO_MES_ANTERIOR"] / tab[ant].where(tab[ant] != 0)
    return tab.reset_index()


def aba_comparativo_status(todos: pd.DataFrame, competencia: str, n: int) -> pd.DataFrame:
    meses = _meses_ate(todos, competencia, n)
    tmp = todos[todos["MES"].isin(meses)]
    if tmp.empty:
        return pd.DataFrame()
    linhas = []
    for mes in meses:
        d = tmp[tmp["MES"] == mes]
        total = d["VALOR_MOEDA"].sum()
        linha = {"MES": mes, "TOTAL": total}
        for st in STATUS_ORDEM:
            linha[st] = d.loc[d["STATUS"] == st, "VALOR_MOEDA"].sum()
        linha["Combustível em Manutenção"] = d.loc[d["COMBUSTIVEL_EM_MANUTENCAO"], "VALOR_MOEDA"].sum()
        linha["Manutenção fora da conta"] = d.loc[d["MANUTENCAO_FORA_DA_CONTA"], "VALOR_MOEDA"].sum()
        linha["Não manutenção em Manutenção"] = d.loc[d["NAO_MANUTENCAO_EM_MANUTENCAO"], "VALOR_MOEDA"].sum()
        linha["%_CLASSIFICADO"] = _pct(d.loc[d["CATEGORIA"].notna(), "VALOR_MOEDA"].sum(), total)
        linhas.append(linha)
    return pd.DataFrame(linhas)


def aba_validacao(razao_total: float, razao_qtd: int, todos: pd.DataFrame, fora: pd.DataFrame,
                  excluidos: pd.DataFrame, mes: pd.DataFrame, regras: Regras, competencia: str) -> pd.DataFrame:
    linhas = [
        ("Razão completo (todas as contas, todos os meses)", razao_total, razao_qtd),
        ("Contas cadastradas na aba CONTAS (todos os meses)", *_soma(todos)),
        ("Excluído por premissa (tipo de lançamento/transação)", *_soma(excluidos)),
        ("Contas fora do escopo (não cadastradas)", *_soma(fora)),
        (f"Analisado na competência {competencia}", *_soma(mes)),
        ("Lançamentos sem data válida (sem mês)", *_soma(todos, todos["MES"].isna())),
    ]
    df = pd.DataFrame(linhas, columns=["ITEM", "VALOR", "QTD"])

    usadas = set(todos["CHAVE_CONTA"])
    sem_lanc = regras.contas[~regras.contas["CHAVE_CONTA"].isin(usadas)]["NOME_CONTA"].tolist()
    extra = pd.DataFrame({"ITEM": [f"Conta cadastrada sem nenhum lançamento no razão: {c}" for c in sem_lanc]})
    return pd.concat([df, extra], ignore_index=True)


def aba_contas_fora(fora: pd.DataFrame) -> pd.DataFrame:
    if fora.empty:
        return pd.DataFrame(columns=["NOME_CONTA", "VALOR", "QTD"])
    g = fora.groupby("NOME_CONTA", dropna=False).agg(VALOR=("VALOR_MOEDA", "sum"), QTD=("VALOR_MOEDA", "size")).reset_index()
    return g.sort_values("VALOR", ascending=False)


def aba_atividades_base(base_lookup: pd.DataFrame, regras: Regras) -> pd.DataFrame:
    if base_lookup.empty or base_lookup["BASE_ATIVIDADE"].dropna().empty:
        return pd.DataFrame(columns=["ATIVIDADE", "QTD_TEXTOS_NA_BASE", "CATEGORIA_MAPEADA"])
    g = base_lookup.groupby("BASE_ATIVIDADE").size().rename("QTD_TEXTOS_NA_BASE").reset_index()
    g = g.rename(columns={"BASE_ATIVIDADE": "ATIVIDADE"})
    g["CATEGORIA_MAPEADA"] = g["ATIVIDADE"].map(lambda a: regras.mapa_atividade.get(normaliza_chave(a), ""))
    return g.sort_values("QTD_TEXTOS_NA_BASE", ascending=False)


# ---------------------------------------------------------------------------
# Exportação
# ---------------------------------------------------------------------------

_CABECALHO = PatternFill("solid", fgColor="1F4E78")
_FONTE_CAB = Font(bold=True, color="FFFFFF")
_PALAVRAS_VALOR = ("VALOR", "TOTAL", "VARIACAO_MES")


def _formata_aba(ws, df: pd.DataFrame, startrow: int = 0):
    for col_idx, nome in enumerate(df.columns, start=1):
        cel = ws.cell(row=startrow + 1, column=col_idx)
        cel.fill, cel.font = _CABECALHO, _FONTE_CAB
        cel.alignment = Alignment(wrap_text=True, vertical="center")

        nome_str = str(nome)
        serie = df[nome]
        eh_pct = nome_str.startswith("%") or nome_str.endswith("%") or "_%" in nome_str
        eh_num = pd.api.types.is_numeric_dtype(serie) and not pd.api.types.is_bool_dtype(serie)
        formato = None
        if eh_pct:
            formato = "0.0%"
        elif eh_num and not nome_str.startswith(("QTD", "PRIORIDADE", "COD_")):
            formato = "#,##0.00"
        elif pd.api.types.is_datetime64_any_dtype(serie):
            formato = "DD/MM/YYYY"
        if formato:
            for linha in range(startrow + 2, startrow + 2 + len(df)):
                ws.cell(row=linha, column=col_idx).number_format = formato

        amostra = max((len(str(v)) for v in serie.head(200)), default=0)
        largura = min(max(len(nome_str), int(amostra or 0)) + 2, 60)
        letra = get_column_letter(col_idx)
        ws.column_dimensions[letra].width = max(ws.column_dimensions[letra].width or 0, largura)


def exporta_excel(caminho, abas: dict):
    """abas: {nome_aba: DataFrame ou lista de (titulo, DataFrame)}"""
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome, conteudo in abas.items():
            blocos = conteudo if isinstance(conteudo, list) else [(None, conteudo)]
            linha = 0
            for titulo, df in blocos:
                if df is None:
                    continue
                if titulo:
                    pd.DataFrame().to_excel(writer, sheet_name=nome, startrow=linha)
                    ws = writer.sheets[nome]
                    ws.cell(row=linha + 1, column=1, value=titulo).font = Font(bold=True, size=12)
                    linha += 1
                df.to_excel(writer, sheet_name=nome, index=False, startrow=linha)
                ws = writer.sheets[nome]
                _formata_aba(ws, df, startrow=linha)
                if len(blocos) == 1:
                    ws.freeze_panes = ws.cell(row=linha + 2, column=1)
                    if len(df.columns):
                        ws.auto_filter.ref = f"A{linha + 1}:{get_column_letter(len(df.columns))}{linha + 1 + len(df)}"
                linha += len(df) + 3
