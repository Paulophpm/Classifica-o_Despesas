import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

import analise_mensal  # noqa: E402
from classificacao.classificador import classifica, prepara_base, prepara_razao  # noqa: E402
from classificacao.regras import carrega_regras  # noqa: E402
from classificacao.relatorios import exporta_excel  # noqa: E402
from criar_planilha_regras import monta_abas  # noqa: E402


def _regras(tmp_path, excecoes=None):
    abas = monta_abas()
    if excecoes:
        abas["EXCECOES"] = pd.DataFrame(excecoes, columns=["TEXTO", "NOME_CONTA", "CATEGORIA", "OBS"])
    caminho = tmp_path / "regras.xlsx"
    exporta_excel(caminho, abas)
    return carrega_regras(caminho)


def _razao(linhas):
    return pd.DataFrame(
        [{"DATA_TRANSACAO": data, "NOME_CONTA": conta, "TEXTO_TRANSACAO": texto, "VALOR_MOEDA": valor}
         for data, conta, texto, valor in linhas]
    )


def _classifica(tmp_path, linhas, base=None, excecoes=None):
    regras = _regras(tmp_path, excecoes)
    analise, fora, _ = prepara_razao(_razao(linhas), regras)
    lookup, _ = prepara_base(base)
    return classifica(analise, regras, lookup).set_index("TEXTO_TRANSACAO"), fora


def test_combustivel_em_manutencao(tmp_path):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Pecas e Acessorios", "COMPRA DIESEL S10 AUTO POSTO X", 500)])
    linha = df.loc["COMPRA DIESEL S10 AUTO POSTO X"]
    assert linha["CATEGORIA"] == "Combustível"
    assert linha["STATUS"] == "Divergente"
    assert linha["COMBUSTIVEL_EM_MANUTENCAO"]


def test_palavra_inteira_motor_nao_bate_motorista(tmp_path):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Servicos para Logistica e Frota", "DIARIA MOTORISTA ENTREGA", 100)])
    assert df.loc["DIARIA MOTORISTA ENTREGA", "CATEGORIA"] != "Mão de obra terceirizada (oficina)"


def test_nota_nao_derruba_manutencao(tmp_path):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Servicos para Logistica e Frota", "TROCA DE OLEO CONFORME NOTA 123", 80)])
    linha = df.loc["TROCA DE OLEO CONFORME NOTA 123"]
    assert linha["CATEGORIA"] == "Troca de óleo (serviço)"
    assert linha["STATUS"] == "OK"


def test_manutencao_fora_da_conta(tmp_path):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Combustiveis e Lubrificantes", "PNEU 295/80 R22", 2000)])
    linha = df.loc["PNEU 295/80 R22"]
    assert linha["CATEGORIA"] == "Pneus"
    assert linha["MANUTENCAO_FORA_DA_CONTA"]


def test_pessoa_fisica_vai_para_revisao_e_excecao_resolve(tmp_path):
    linhas = [("2026-10-05", "Servicos para Logistica e Frota", "RICARDO SOUZA", 300)]
    df, _ = _classifica(tmp_path, linhas)
    assert df.loc["RICARDO SOUZA", "STATUS"] == "Revisar"
    assert df.loc["RICARDO SOUZA", "POSSIVEL_PF"]

    excecao = [("Ricardo Souza", "", "Mão de obra pessoa física", "mecânico de Campinas")]
    df, _ = _classifica(tmp_path, linhas, excecoes=excecao)
    assert df.loc["RICARDO SOUZA", "CATEGORIA"] == "Mão de obra pessoa física"
    assert df.loc["RICARDO SOUZA", "STATUS"] == "OK"
    assert df.loc["RICARDO SOUZA", "ORIGEM_CLASSIFICACAO"] == "Exceção manual"


def test_estorno_e_ajuste(tmp_path):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Pecas e Acessorios", "ESTORNO PECAS NF 10", -50)])
    assert df.loc["ESTORNO PECAS NF 10", "STATUS"] == "Ajuste contábil"


def test_base_procv_sim_nao_e_regra_especifica(tmp_path):
    base = pd.DataFrame({"TRANSACAO": ["FORNECEDOR XPTO", "FILTRO DE AR"], "MANUTENCAO": ["Sim", "Não"],
                         "ATIVIDADE": [None, None]})
    df, _ = _classifica(tmp_path, [
        ("2026-10-05", "Pecas e Acessorios", "FORNECEDOR XPTO", 10),
        ("2026-10-05", "Pecas e Acessorios", "FILTRO DE AR", 20),
    ], base=base)
    assert df.loc["FORNECEDOR XPTO", "CATEGORIA"] == "Manutenção - não detalhado"
    assert df.loc["FORNECEDOR XPTO", "CONFIANCA"] == "Baixa"
    # regra específica vence o Sim/Não da base e o conflito é sinalizado
    assert df.loc["FILTRO DE AR", "CATEGORIA"] == "Peças e acessórios"
    assert df.loc["FILTRO DE AR", "CONFLITO_BASE_x_REGRA"]


def test_padrao_da_conta_e_nome_com_acento(tmp_path):
    df, fora = _classifica(tmp_path, [
        ("2026-10-05", "Pedagios", "FATURA 9981", 70),          # sem acento no razão
        ("2026-10-05", "Energia Eletrica", "CONTA DE LUZ", 999),  # fora do escopo
    ])
    assert df.loc["FATURA 9981", "CATEGORIA"] == "Pedágio"
    assert df.loc["FATURA 9981", "ORIGEM_CLASSIFICACAO"] == "Padrão da conta"
    assert len(fora) == 1


def test_fluxo_completo_gera_excel(tmp_path):
    exporta_excel(tmp_path / "regras.xlsx", monta_abas())
    _razao([
        ("2026-09-10", "Pecas e Acessorios", "AUTO PECAS BRASIL LTDA", 100),
        ("2026-10-10", "Pecas e Acessorios", "AUTO PECAS BRASIL LTDA", 150),
        ("2026-10-11", "Pecas e Acessorios", "ABASTECIMENTO CAMINHAO", 400),
        ("2026-10-12", "Multas", "AIT 123", 200),
        ("2026-10-13", "Servicos para Logistica e Frota", "JOSE SILVA", 300),
        ("2026-10-14", "Servicos para Logistica e Frota", "XYZ QWERTY", 50),
    ]).to_excel(tmp_path / "razao.xlsx", index=False)

    rc = analise_mensal.main([
        "--razao", str(tmp_path / "razao.xlsx"), "--regras", str(tmp_path / "regras.xlsx"),
        "--base", str(tmp_path / "nao_existe.xlsx"), "--saida", str(tmp_path / "out"), "--competencia", "2026-10",
    ])
    assert rc == 0
    saida = pd.read_excel(tmp_path / "out" / "analise_logistica_2026-10.xlsx", sheet_name=None)
    assert {"RESUMO", "DIVERGENCIAS", "PENDENTES", "MATRIZ_CATEGORIAS", "COMPARATIVO_STATUS"} <= set(saida)
    assert len(saida["BASE_CLASSIFICADA"]) == 5
    assert "JOSE SILVA" in set(saida["PENDENTES"]["TEXTO"])
    assert "XYZ QWERTY" in set(saida["PENDENTES"]["TEXTO"])


def test_planilha_com_erro_avisa_claramente(tmp_path):
    abas = monta_abas()
    abas["REGRAS"].loc[0, "CATEGORIA"] = "Categoria que não existe"
    exporta_excel(tmp_path / "regras.xlsx", abas)
    with pytest.raises(ValueError, match="não está cadastrada"):
        carrega_regras(tmp_path / "regras.xlsx")


@pytest.mark.parametrize("texto, categoria", [
    ("FILTRO DE OLEO", "Peças e acessórios"),
    ("OLEO LUBRIFICANTE 15W40", "Lubrificantes"),
    ("MULTA DETRAN", "Multa de trânsito"),
    ("LICENCIAMENTO DETRAN", "Licenciamento"),
    ("BICO INJETOR DIESEL", "Peças e acessórios"),
])
def test_regras_especificas(tmp_path, texto, categoria):
    df, _ = _classifica(tmp_path, [("2026-10-05", "Pecas e Acessorios", texto, 10)])
    assert df.loc[texto, "CATEGORIA"] == categoria
