"""
Gera o MODELO da planilha de regras (config/regras.xlsx).

Rode uma vez para criar o arquivo. Depois disso, a planilha passa a ser a
fonte da verdade: ajuste regras, contas, categorias e premissas direto no
Excel. Este script NÃO sobrescreve uma planilha existente (use --forcar).

As regras iniciais foram migradas de legado/config.py, com estes ajustes:
  - em vez de Sim/Não, cada termo aponta para uma CATEGORIA;
  - combustível/ticket deixaram de ser "manutenção" (agora são Combustível);
  - termos que batiam em palavras demais (car, desc, nota, contas, titulo,
    boleto, prestacao, equipamentos) foram removidos;
  - a busca passou a ser por palavra inteira (PNEU não bate em PNEUMATICO,
    MOTOR não bate em MOTORISTA), salvo quando TIPO_MATCH = prefixo/contem.
"""

import argparse
from pathlib import Path

import pandas as pd

from classificacao.relatorios import exporta_excel

MANUT = "Manutenção"
COMB = "Combustíveis e Lubrificantes"
SEG = "Seguros"
MAT = "Materiais de Logística e Assist. Técnica"
FRETE = "Fretes"
MULTA = "Multas"
PED = "Pedágios"
IPVA = "IPVA e Emplacamento"
OUTRAS = "Outras despesas (fora da logística)"
AJUSTE = "Ajuste contábil"
VERIF = "A verificar"

# ---------------------------------------------------------------------------
PREMISSAS = [
    ("CAMPOS_TEXTO", "TEXTO_TRANSACAO", "Colunas do razão usadas nas regras, separadas por ';' (ex.: TEXTO_TRANSACAO;TEXTO_LOG)."),
    ("GRUPO_MANUTENCAO", MANUT, "Nome do grupo que representa Manutenção (usado nos indicadores)."),
    ("GRUPO_COMBUSTIVEL", COMB, "Nome do grupo que representa Combustível (indicador 'Combustível em Manutenção')."),
    ("PRIORIDADE_EXCECAO", 0, "Exceções manuais vencem tudo."),
    ("PRIORIDADE_BASE_ATIVIDADE", 50, "Base PROCV quando a ATIVIDADE tem de-para na aba MAPA_ATIVIDADE."),
    ("PRIORIDADE_BASE_SIM_NAO", 500, "Base PROCV quando só há Sim/Não. Regras com prioridade menor vencem a base."),
    ("PRIORIDADE_PF", 990, "Pessoa física detectada (só quando nada antes classificou)."),
    ("PRIORIDADE_PADRAO_CONTA", 995, "Categoria padrão da conta (última tentativa)."),
    ("DETECTAR_PF", "SIM", "SIM/NAO - marca possíveis pessoas físicas pela aba PESSOA_FISICA."),
    ("APLICAR_PADRAO_CONTA", "SIM", "SIM/NAO - usa CATEGORIA_PADRAO da aba CONTAS quando nada mais classificou."),
    ("EXCLUIR_TIPO_LANCAMENTO", "", "Valores de TIPO_LANCAMENTO a excluir da análise, separados por ';'."),
    ("EXCLUIR_TIPO_TRANSACAO", "", "Valores de TIPO_TRANSACAO a excluir da análise, separados por ';'."),
    ("VALOR_MINIMO_PENDENTE", 0, "Só lista na aba PENDENTES textos com |valor| >= este."),
    ("MESES_COMPARATIVO", 6, "Quantidade de meses nas abas de comparativo."),
]

# (NOME_CONTA, GRUPO_ATUAL, CATEGORIA_PADRAO, OBS)
CONTAS = [
    ("Servicos para Logistica e Frota", MANUT, "", "Conta de manutenção - sem padrão, para enxergar o detalhe"),
    ("Pecas e Acessorios", MANUT, "", "Conta de manutenção"),
    ("Pecas e Acessorios Logistica e Frota", MANUT, "", "Conta de manutenção"),
    ("Pecas e Acessorios para Logistica e Frota", MANUT, "", "Mesmo nome com 'para' (usado no código antigo)"),
    ("Combustiveis e Lubrificantes", COMB, "Combustível", ""),
    ("Seguros", SEG, "Seguro de veículo", ""),
    ("Seguro Veiculo do socio", SEG, "Seguro veículo de sócio", ""),
    ("Materiais de Logistica e Assistencia Tecnica", MAT, "", ""),
    ("Fretes", FRETE, "Frete", ""),
    ("Multas", MULTA, "Multa de trânsito", ""),
    ("Pedágios", PED, "Pedágio", ""),
    ("IPVA", IPVA, "IPVA", ""),
    ("IPVA  Veiculo do socio", IPVA, "IPVA veículo de sócio", ""),
    ("Emplacamento", IPVA, "Emplacamento", ""),
]

# (CATEGORIA, GRUPO_ESPERADO, DESCRICAO)
CATEGORIAS = [
    # Manutenção
    ("Peças e acessórios", MANUT, "Peças em geral (filtros, freio, suspensão, correias...)"),
    ("Pneus", MANUT, "Pneus, câmaras, recapagem"),
    ("Serviço de pneus / borracharia", MANUT, "Borracharia, alinhamento, balanceamento"),
    ("Mão de obra terceirizada (oficina)", MANUT, "Oficinas, mecânica, revisão, retífica"),
    ("Mão de obra pessoa física", MANUT, "Mecânico autônomo / PF identificado"),
    ("Elétrica automotiva", MANUT, "Auto elétrica, injeção, bateria"),
    ("Funilaria e pintura", MANUT, ""),
    ("Lavagem", MANUT, ""),
    ("Guincho / socorro", MANUT, ""),
    ("Tacógrafo", MANUT, "Aferição de tacógrafo"),
    ("Troca de óleo (serviço)", MANUT, "Serviço de troca; o óleo em si é Lubrificantes"),
    ("Manutenção - genérico veículo", MANUT, "Texto só cita veículo/caminhão, sem detalhe"),
    # Combustíveis e Lubrificantes
    ("Combustível", COMB, "Diesel, gasolina, etanol, GNV, postos, cartão combustível"),
    ("Arla 32", COMB, ""),
    ("Lubrificantes", COMB, "Óleo e lubrificantes (produto)"),
    # Outros grupos logísticos
    ("Frete", FRETE, ""),
    ("Chapa (carga e descarga)", FRETE, "Ajudante avulso de carga/descarga - CONFIRMAR grupo"),
    ("Multa de trânsito", MULTA, ""),
    ("Pedágio", PED, ""),
    ("IPVA", IPVA, ""),
    ("IPVA veículo de sócio", IPVA, ""),
    ("Licenciamento", IPVA, "Licenciamento, DPVAT, taxas Detran"),
    ("Emplacamento", IPVA, ""),
    ("Despachante", IPVA, ""),
    ("Seguro de veículo", SEG, ""),
    ("Seguro veículo de sócio", SEG, ""),
    ("Sinistro / franquia", SEG, ""),
    ("Rastreamento / telemetria", MAT, ""),
    ("Equipamentos de movimentação", MAT, "Empilhadeiras, paleteiras"),
    ("Balanças e pesagem", MAT, ""),
    ("EPI e uniformes", MAT, "CONFIRMAR grupo"),
    # Fora da logística
    ("Locação", OUTRAS, "Locação de veículos/equipamentos"),
    ("TI e telecom", OUTRAS, ""),
    ("Serviços administrativos", OUTRAS, "Consultoria, contabilidade, jurídico, marketing, gráfica"),
    ("Alimentação e viagem", OUTRAS, ""),
    ("Predial e construção", OUTRAS, "Construção, limpeza, vigilância, material elétrico, tintas"),
    ("Benefícios", OUTRAS, ""),
    ("Outros fornecedores (não logístico)", OUTRAS, "Fornecedores herdados como 'Não' do código antigo"),
    # Ajustes e revisão
    ("Estorno / ajuste contábil", AJUSTE, "Estorno, provisão, reclassificação, saldo, duplicidade"),
    ("Reembolso a colaborador", VERIF, "Pode ser despesa legítima paga por motorista"),
    ("Despesas gerais (sem detalhe)", VERIF, "Ex.: 'despesas semanais', 'despesas operacionais'"),
]

# Regras: (PRIORIDADE, TERMO, TIPO_MATCH, CATEGORIA, SOMENTE_GRUPO_ATUAL, OBS)
P_AJUSTE, P_COMB, P_FORTE, P_ESPEC, P_OUTRAS, P_FRACA, P_GENERICA = 10, 20, 30, 200, 300, 800, 900


def _r(prio, cat, termos, tipo="palavra", somente="", obs=""):
    return [(prio, t, tipo, cat, somente, obs) for t in termos]


REGRAS = (
    # --- Ajustes contábeis (vencem tudo) ---
    _r(P_AJUSTE, "Estorno / ajuste contábil", [
        "estorno", "estornos", "provisao", "reversao de provisao", "reclassificacao", "reeclassificacao", "reclass",
        "ajuste", "ajuste de fechamento", "fechamento de estoque", "saldo inicial", "duplicidade", "erro lancamento",
        "baixa", "transferencia", "transferecia",
    ])
    # --- Combustível (vence a base PROCV, que marcava como manutenção) ---
    + _r(P_COMB, "Arla 32", ["arla", "arla 32"])
    + _r(P_COMB, "Combustível", [
        "combustivel", "combustiveis", "posto de combustivel", "diesel", "oleo diesel", "gasolina", "etanol", "gnv",
        "abastecimento", "abastec", "ticket log", "ticket car", "auto posto", "posto",
    ])
    + _r(P_COMB, "Combustível", ["ticket"], somente=COMB, obs="'ticket' só vira combustível dentro da conta de combustível")
    + _r(P_COMB, "Lubrificantes", ["lubrificante", "lubrificantes", "oleo", "oleos", "graxa", "fluido de freio", "aditivo"],
         obs="CONFIRMAR premissa: óleo é Lubrificantes (Comb. e Lub.), não Manutenção")
    + _r(P_COMB - 1, "Troca de óleo (serviço)", ["troca de oleo", "troca oleo"])
    + _r(P_COMB - 1, "Peças e acessórios", ["filtro de oleo", "filtro do oleo", "filtro oleo", "retentor de oleo",
                                            "bomba de oleo", "filtro de combustivel", "filtro combustivel",
                                            "bomba de combustivel", "tanque de combustivel", "bico injetor"],
         obs="Peças que citam óleo/combustível no nome")
    # --- Outros grupos específicos ---
    + _r(P_FORTE - 1, "Multa de trânsito", ["multa", "multas", "infracao", "autuacao", "notificacao de autuacao"])
    + _r(P_FORTE, "Pedágio", ["pedagio", "pedagios", "sem parar", "conectcar", "veloe", "vale pedagio", "move mais"])
    + _r(P_FORTE, "IPVA", ["ipva"])
    + _r(P_FORTE, "Licenciamento", ["licenciamento", "dpvat", "crlv", "detran", "taxa de licenciamento"])
    + _r(P_FORTE, "Emplacamento", ["emplacamento", "lacre de placa"])
    + _r(P_FORTE, "Despachante", ["despachante", "despachente"])
    + _r(P_FORTE, "Sinistro / franquia", ["sinistro", "sinistro placa", "franquia"])
    + _r(P_FORTE, "Seguro de veículo", ["seguro", "seguros", "seguradora", "apolice"])
    + _r(P_FORTE, "Frete", ["frete", "fretes", "transportadora", "cte", "conhecimento de transporte"])
    + _r(P_FORTE, "Chapa (carga e descarga)", ["chapa", "chapas", "carga e descarga", "descarga"])
    + _r(P_FORTE, "Reembolso a colaborador", ["reembolso", "ressarcimento"])
    # --- Manutenção detalhada ---
    + _r(P_ESPEC, "Pneus", ["pneu", "pneus", "camara de ar", "recapagem", "recauchutagem", "recapadora"])
    + _r(P_ESPEC, "Serviço de pneus / borracharia", [
        "borracharia", "borracheiro", "vulcanizadora", "alinhamento", "balanceamento", "geometria"])
    + _r(P_ESPEC, "Funilaria e pintura", ["funilaria", "lanternagem", "pintura automotiva", "martelinho"])
    + _r(P_ESPEC, "Lavagem", ["lavagem", "lava jato", "lavajato", "lava rapido", "higienizacao"])
    + _r(P_ESPEC, "Guincho / socorro", ["guincho", "reboque", "socorro mecanico", "auto socorro"])
    + _r(P_ESPEC, "Tacógrafo", ["tacografo", "tacografos", "cronotacografo"])
    + _r(P_ESPEC, "Elétrica automotiva", [
        "auto eletrica", "eletrica automotiva", "injecao eletronica", "bateria", "baterias", "bateria automotiva",
        "alternador", "motor de partida"])
    + _r(P_ESPEC, "Peças e acessórios", [
        "pecas", "peca", "auto pecas", "autopecas", "de pecas", "filtro", "filtros", "freio", "freios", "lona de freio",
        "pastilha", "disco de freio", "suspensao", "embreagem", "radiador", "escapamento", "rolamento", "correia",
        "amortecedor", "molas", "mola", "molas automotivas", "valvula", "retentor", "cardan", "cruzeta", "junta homocinetica", "junta do cabecote",
        "mangueira", "farol", "lampada", "retrovisor", "para choque", "parabrisa", "vidracaria automotiva",
        "acessorios automotivos", "equipamentos automotivos", "rodas", "roda", "cambio", "borracha", "borrachas"])
    + _r(P_ESPEC, "Mão de obra terceirizada (oficina)", [
        "oficina", "mecanica", "mecanico", "mao de obra", "servico mecanico", "revisao", "revisao veicular",
        "manutencao", "manutencao veicular", "retifica", "motor", "motores", "auto center", "autocenter",
        "chaveiro automotivo", "ar condicionado automotivo", "limpeza de bico", "limpeza de bicos"])
    + _r(P_ESPEC, "Rastreamento / telemetria", ["rastreador", "rastreamento", "telemetria", "rastreio"])
    + _r(P_ESPEC, "Equipamentos de movimentação", ["empilhadeira", "empilhadeiras", "paleteira", "transpaleteira"])
    + _r(P_ESPEC, "Balanças e pesagem", ["balanca", "balancas", "pesagem", "toledo"])
    + _r(P_ESPEC, "EPI e uniformes", ["epi", "epis", "uniforme", "uniformes"])
    # --- Fora da logística ---
    + _r(P_OUTRAS, "Locação", ["locacao", "locacao de equipamentos", "asa rent", "locadora", "aluguel"])
    + _r(P_OUTRAS, "TI e telecom", ["informatica", "informata", "tecnologia", "software", "internet", "telefonia",
                                     "telecomunicacoes"])
    + _r(P_OUTRAS, "Serviços administrativos", ["consultoria", "contabilidade", "advocacia", "juridico", "marketing",
                                                 "publicidade", "grafica", "papelaria", "escritorio", "treinamento",
                                                 "educacao", "recursos humanos", "midia digital", "midia digitais"])
    + _r(P_OUTRAS, "Alimentação e viagem", ["alimentacao", "alimentos", "restaurante", "lanchonete", "hotel",
                                             "hospedagem", "viagem", "passagem aerea"])
    + _r(P_OUTRAS, "Predial e construção", ["construcao civil", "material de construcao", "contrucao", "limpeza",
                                             "limpeza predial", "seguranca patrimonial", "vigilancia", "material eletrico",
                                             "materiais eletricos", "tintas", "refrigeracao", "refigreracao"])
    + _r(P_OUTRAS, "Benefícios", ["swile", "vale refeicao", "vale alimentacao"])
    + _r(P_OUTRAS, "Outros fornecedores (não logístico)", ["fusion", "magazine luiza", "metalurgica", "estacionamento"],
         obs="Herdado do código antigo como 'Não' - revisar")
    + _r(P_OUTRAS, "Despesas gerais (sem detalhe)", ["despesas operacionais", "despesas administrativas",
                                                      "despesas vendedores", "despesas semanais"])
    # --- Fracas / genéricas ---
    + _r(P_FRACA, "Emplacamento", ["placa", "placas"], obs="Fraca: placa costuma só identificar o veículo")
    + _r(P_FRACA, "Mão de obra terceirizada (oficina)", ["ar condicionado"], somente=MANUT)
    + _r(P_GENERICA, "Manutenção - genérico veículo", [
        "veiculo", "veiculos", "veicular", "automotivo", "automotiva", "automoveis", "carro", "carros", "caminhao",
        "caminhoes", "onibus", "truck", "mobi", "mobbi", "strada", "frota"])
)

# Exemplos para PESSOA_FISICA (amplie com os nomes que aparecem no seu razão)
NOMES = """JOSE JOAO ANTONIO FRANCISCO CARLOS PAULO PEDRO LUCAS LUIZ LUIS MARCOS GABRIEL RAFAEL DANIEL MARCELO BRUNO
EDUARDO FELIPE RAIMUNDO RODRIGO MANOEL MANUEL MATEUS MATHEUS ANDRE FERNANDO FABIO LEONARDO GUSTAVO GUILHERME LEANDRO
TIAGO THIAGO ANDERSON RICARDO MARCIO JORGE SEBASTIAO ALEXANDRE ROBERTO EDSON DIEGO VITOR VICTOR SERGIO CLAUDIO
JULIO RENATO JEFERSON JEFFERSON WELLINGTON ADRIANO VALDIR ROGERIO SILVIO MARIO ALEX WAGNER GERALDO ADEMIR JAIR
MARIA ANA FRANCISCA ANTONIA ADRIANA JULIANA MARCIA FERNANDA PATRICIA ALINE SANDRA CAMILA AMANDA BRUNA JESSICA
LETICIA JULIA LUCIANA VANESSA MARIANA GABRIELA VERA VITORIA LARISSA CLAUDIA BEATRIZ RITA LUANA SONIA RENATA ELIANE""".split()
SOBRENOMES = """SILVA SANTOS OLIVEIRA SOUZA SOUSA RODRIGUES FERREIRA ALVES PEREIRA LIMA GOMES COSTA RIBEIRO MARTINS
CARVALHO ALMEIDA LOPES SOARES FERNANDES VIEIRA BARBOSA ROCHA DIAS NASCIMENTO ANDRADE MOREIRA NUNES MARQUES MACHADO
MENDES FREITAS CARDOSO RAMOS GONCALVES SANTANA TEIXEIRA ARAUJO PINTO CORREIA CAVALCANTE BATISTA MONTEIRO MOURA
CAMPOS BORGES REIS MIRANDA MOTA FARIAS PIRES CASTRO JESUS SALES BEZERRA""".split()
EMPRESA = """LTDA ME EPP EIRELI SA S A CIA COMERCIO COMERCIAL SERVICOS INDUSTRIA DISTRIBUIDORA TRANSPORTES LOGISTICA
AUTO POSTO PECAS OFICINA MECANICA CENTER GRUPO HOLDING MOTORS VEICULOS""".split()


def monta_abas():
    regras = pd.DataFrame(REGRAS, columns=["PRIORIDADE", "TERMO", "TIPO_MATCH", "CATEGORIA", "SOMENTE_GRUPO_ATUAL", "OBS"])
    regras.insert(0, "ATIVO", "SIM")
    pf = pd.DataFrame(
        [(n, "NOME") for n in NOMES] + [(s, "SOBRENOME") for s in SOBRENOMES] + [(e, "EMPRESA") for e in EMPRESA],
        columns=["TERMO", "TIPO"],
    )
    return {
        "PREMISSAS": pd.DataFrame(PREMISSAS, columns=["PARAMETRO", "VALOR", "DESCRICAO"]),
        "CONTAS": pd.DataFrame(CONTAS, columns=["NOME_CONTA", "GRUPO_ATUAL", "CATEGORIA_PADRAO", "OBS"]),
        "CATEGORIAS": pd.DataFrame(CATEGORIAS, columns=["CATEGORIA", "GRUPO_ESPERADO", "DESCRICAO"]),
        "REGRAS": regras,
        "EXCECOES": pd.DataFrame(columns=["TEXTO", "NOME_CONTA", "CATEGORIA", "OBS"]),
        "MAPA_ATIVIDADE": pd.DataFrame(columns=["ATIVIDADE", "CATEGORIA"]),
        "PESSOA_FISICA": pf,
    }


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--destino", default="config/regras.xlsx")
    p.add_argument("--forcar", action="store_true", help="Sobrescreve a planilha se já existir")
    args = p.parse_args()

    destino = Path(args.destino)
    if destino.exists() and not args.forcar:
        print(f"{destino} já existe - nada foi alterado (use --forcar para sobrescrever).")
        return
    destino.parent.mkdir(parents=True, exist_ok=True)
    exporta_excel(destino, monta_abas())
    print(f"Planilha de regras criada em {destino}")


if __name__ == "__main__":
    main()
