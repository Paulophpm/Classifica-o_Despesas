"""
Análise mensal de naturezas logísticas do razão de despesas.

Uso:
    python analise_mensal.py --competencia 2026-10
    python analise_mensal.py                       # usa o último mês do razão

Veja `python analise_mensal.py --help` para todos os parâmetros.
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from classificacao import relatorios as rel
from classificacao.classificador import classifica, prepara_base, prepara_razao
from classificacao.regras import ErroRegras, carrega_regras


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Classificação de naturezas logísticas do razão de despesas.")
    p.add_argument("--competencia", help="Mês analisado no formato AAAA-MM. Padrão: último mês presente no razão.")
    p.add_argument("--razao", default="data/razao.xlsx", help="Arquivo do razão (padrão: data/razao.xlsx)")
    p.add_argument("--aba", default=None, help="Aba do razão (padrão: primeira aba)")
    p.add_argument("--base", default="data/Transacoes_base.xlsx", help="Base de referência do PROCV")
    p.add_argument("--regras", default="config/regras.xlsx", help="Planilha de regras e premissas")
    p.add_argument("--saida", default="output", help="Pasta de saída")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        regras = carrega_regras(args.regras)
    except ErroRegras as e:
        print(f"ERRO na planilha de regras ({args.regras}):\n{e}")
        return 1

    print(f"Lendo razão: {args.razao}")
    razao = pd.read_excel(args.razao, sheet_name=args.aba if args.aba is not None else 0)
    razao_qtd = len(razao)

    base = None
    if args.base and Path(args.base).exists():
        print(f"Lendo base de referência: {args.base}")
        base = pd.read_excel(args.base)
    else:
        print(f"AVISO: base de referência não encontrada ({args.base}); seguindo só com regras.")

    analise, fora, excluidos = prepara_razao(razao, regras)
    razao_total = float(pd.concat([analise["VALOR_MOEDA"], fora["VALOR_MOEDA"], excluidos["VALOR_MOEDA"]]).sum())
    base_lookup, conflitos_base = prepara_base(base)
    todos = classifica(analise, regras, base_lookup)

    meses = sorted(todos["MES"].dropna().unique())
    if not meses:
        print("ERRO: nenhum lançamento com data válida nas contas cadastradas.")
        return 1
    competencia = args.competencia or meses[-1]
    mes = todos[todos["MES"] == competencia]
    if mes.empty:
        print(f"ERRO: nenhum lançamento na competência {competencia}. Meses disponíveis: {', '.join(meses)}")
        return 1

    n_meses = regras.premissa_int("MESES_COMPARATIVO")
    premissas_usadas = pd.DataFrame(sorted(regras.premissas.items()), columns=["PARAMETRO", "VALOR"])
    premissas_usadas["VALOR"] = premissas_usadas["VALOR"].astype(str)

    abas = {
        "RESUMO": [
            (f"Resumo da competência {competencia}", rel.aba_resumo(mes, competencia)),
            ("Valor por grupo atual (conta) e status", rel.aba_resumo_por_grupo(mes)),
        ],
        "MATRIZ_GRUPOS": rel.aba_matriz_grupos(mes),
        "MATRIZ_CATEGORIAS": rel.aba_matriz_categorias(mes),
        "CATEGORIA_POR_CONTA": rel.aba_categoria_por_conta(mes),
        "DIVERGENCIAS": rel.aba_divergencias(mes),
        "PENDENTES": rel.aba_pendentes(mes, regras.premissa_float("VALOR_MINIMO_PENDENTE")),
        "COMPARATIVO_STATUS": rel.aba_comparativo_status(todos, competencia, n_meses),
        "COMPARATIVO_CATEGORIAS": rel.aba_comparativo_categorias(todos, competencia, n_meses),
        "BASE_CLASSIFICADA": rel.colunas_detalhe(mes),
        "CONFLITOS_BASE": conflitos_base,
        "ATIVIDADES_BASE": rel.aba_atividades_base(base_lookup, regras),
        "CONTAS_FORA_ESCOPO": rel.aba_contas_fora(fora),
        "VALIDACAO": rel.aba_validacao(razao_total, razao_qtd, todos, fora, excluidos, mes, regras, competencia),
        "PREMISSAS_USADAS": premissas_usadas,
    }

    Path(args.saida).mkdir(parents=True, exist_ok=True)
    caminho = Path(args.saida) / f"analise_logistica_{competencia}.xlsx"
    rel.exporta_excel(caminho, abas)

    resumo = abas["RESUMO"][0][1]
    print(f"\nCompetência {competencia}")
    for _, linha in resumo.iterrows():
        print(f"  {linha['INDICADOR']:<50} {linha['VALOR']:>18,.2f}  ({linha['%_DO_TOTAL']:6.1%})  {linha['QTD_LANCAMENTOS']:>7} lanç.")
    print(f"\nArquivo gerado: {caminho}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
