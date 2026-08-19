import datetime
import pandas as pd
import re
import unicodedata

#importação dos arquivos e da base de dados
razao = pd.read_excel("razao.xlsx", sheet_name="2026")
base = pd.read_excel("Transacoes_base.xlsx")

Total_contas= razao['VALOR_MOEDA'].sum()

# Selecionar apenas as colunas que serão utilizadas na análise
razao = razao[['COD_ESTABELECIMENTO','DATA_TRANSACAO','NOME_CONTA','TEXTO_TRANSACAO','VALOR_MOEDA']]

#adicionar uma nova coluna para classificação de manutenção, copia da texto transação
razao['TEXTO_MANUTENCAO'] = razao['TEXTO_TRANSACAO']

#criação de uma novas colunas
razao['TIPO_ATIVIDADE'] = None

#Transformar as colunas VALOR_MOEDA e DATA_TRANSACAO para os tipos corretos
razao['VALOR_MOEDA'] = razao['VALOR_MOEDA'].astype('Float64')
razao['DATA_TRANSACAO'] = pd.to_datetime(razao['DATA_TRANSACAO'], errors='coerce')

#criação de uma coluna para o mês e ano da transação, no formato YYYY-MM
razao['MES'] = razao['DATA_TRANSACAO'].dt.strftime('%Y-%m')

# filtrar as colunas apenas da coluna "NOME_CONTA" que contenham  as contas: Serviços para Logística e Frota, Peças e Assessorios, Peças e Assesorios para Logistica e Frota
razao = razao[razao['NOME_CONTA'].str.contains('Servicos para Logistica e Frota|Pecas e Acessorios|Pecas e Acessorios para Logistica e Frota', case=False, na=False)]

Total_contas_manutencao = razao['VALOR_MOEDA'].sum()
                     
# limpeza do texto da coluna de transação das duas bases para melhoria da classificação das transações
def limpa_texto(text):
    if pd.isna(text):
        return text
    text = str(text)
    text = unicodedata.normalize('NFKD', text).encode('ASCII', errors='ignore').decode('ASCII')
    text = re.sub(r'[0-9\-]', ' ', text) # retirar números e hífens
    text = re.sub(r'[^A-Za-z\s]', ' ', text) # retirar outros caracteres especiais
    text = re.sub(r'\s+', ' ', text).strip() # compactar espaços
    text = re.sub(r'\bS\s+NF\b', '', text) # retirar o termo "S NF"
    return text.upper()

razao['TEXTO_MANUTENCAO'] = razao['TEXTO_MANUTENCAO'].apply(limpa_texto)
base['TRANSACAO'] = base['TRANSACAO'].apply(limpa_texto)

# função que faz um PROC VLOOKUP (procv) da coluna texto de razao na coluna TRANSACAO da base
# e preenche a classificação MANUTENCAO na tabela razao

def preenche_manutencao_procv(razao, base,
                               razao_texto_col='TEXTO_MANUTENCAO',
                               base_trans_col='TRANSACAO',
                               manut_col='MANUTENCAO'):
    base_copy = base.copy()
    if 'TRANSACAO' in base_copy.columns:
        base_copy = base_copy.rename(columns={'TRANSACAO': 'TRANSACAO'})
    if 'MANUTENCAO' in base_copy.columns:
        base_copy = base_copy.rename(columns={'MANUTENCAO': 'MANUTENCAO'})

    if base_trans_col not in base_copy.columns:
        raise KeyError(f'Coluna de transação não encontrada em base: {base_trans_col}')
    if manut_col not in base_copy.columns:
        raise KeyError(f'Coluna de manutenção não encontrada em base: {manut_col}')

    base_copy[base_trans_col] = base_copy[base_trans_col].astype(str).apply(limpa_texto)
    resultado = razao.copy()
    resultado[razao_texto_col] = resultado[razao_texto_col].astype(str).apply(limpa_texto)

    base_unique = base_copy[[base_trans_col, manut_col]].drop_duplicates(subset=[base_trans_col])
    resultado = resultado.merge(
        base_unique,
        left_on=razao_texto_col,
        right_on=base_trans_col,
        how='left'
    )

    resultado['MANUTENCAO'] = resultado[manut_col]
    resultado = resultado.drop(columns=[base_trans_col])
    return resultado

# retirada de valores duplicados da base, caso existam
#base = base.drop_duplicates(subset=['TRANSACAO'], keep=False)
nomes_razao = set(razao['TEXTO_MANUTENCAO'].dropna().unique())
nomes_base = set(base['TRANSACAO'].dropna().unique())

# preenche a classificação MANUTENCAO da planilha razao usando PROC V na base
razao = preenche_manutencao_procv(razao, base)

# ---------------------------------------------------------------------------
# CLASSIFICAÇÃO POR PALAVRA-CHAVE (novo passo complementar/prioritário)
# ---------------------------------------------------------------------------
# Dicionário simples: palavra-chave (case-insensitive) -> classificação a aplicar.
# Basta adicionar novas entradas aqui para cobrir novos casos.
# IMPORTANTE: a função limpa_texto() remove acentos do texto antes de comparar,
# então as palavras-chave abaixo também devem estar SEM acento (ex: 'pecas',
# não 'peças'). A busca já é case-insensitive no código, então não é preciso
# cadastrar a mesma palavra em maiúscula e minúscula.
REGRAS_PALAVRA_CHAVE = {
    # --- Manutenção automotiva -> "Sim" ---
    'automotivo': 'Sim',
    'automotiva': 'Sim',
    'automoveis': 'Sim',
    'oficina': 'Sim',
    'mecanica': 'Sim',
    'mecanico': 'Sim',
    'auto pecas': 'Sim',
    'autopecas': 'Sim',
    'auto center': 'Sim',
    'autocenter': 'Sim',
    'funilaria': 'Sim',
    'pintura automotiva': 'Sim',
    'borracharia': 'Sim',
    'borrachas': 'Sim',
    'borracha': 'Sim',
    'pneu': 'Sim',
    'pneus': 'Sim',
    'freio': 'Sim',
    'freios': 'Sim',
    'suspensao': 'Sim',
    'suspensão': 'Sim',
    'embreagem': 'Sim',
    'radiador': 'Sim',
    'injecao eletronica': 'Sim',
    'alinhamento': 'Sim',
    'balanceamento': 'Sim',
    'lavagem': 'Sim',
    'lava jato': 'Sim',
    'lavajato': 'Sim',
    'guincho': 'Sim',
    'socorro mecanico': 'Sim',
    'retifica': 'Sim',
    'auto eletrica': 'Sim',
    'veiculo': 'Sim',
    'veiculos': 'Sim',
    'carro': 'Sim',
    'carros': 'Sim',
    'caminhao': 'Sim',
    'onibus': 'Sim',
    'escapamento': 'Sim',
    'cambio': 'Sim',
    'rolamento': 'Sim',
    'correia': 'Sim',
    'amortecedor': 'Sim',
    'lubrificante': 'Sim',
    'troca de oleo': 'Sim',
    'oleo': 'Sim',
    'filtro': 'Sim',
    'acessorios automotivos': 'Sim',
    'de pecas': 'Sim',
    'pecas': 'Sim',
    'ticket': 'Sim',
    'posto de combustivel': 'Sim',
    'combustivel': 'Sim',
    'bateria automotiva': 'Sim',
    'vidracaria automotiva': 'Sim',
    'chaveiro automotivo': 'Sim',
    'revisao veicular': 'Sim',
    'manutencao veicular': 'Sim',
    'molas automotivas': 'Sim',
    'molas': 'Sim',
    'equipamentos automotivos': 'Sim',
    'equipamentos': 'Sim',
    'veicular': 'Sim',
    'car': 'Sim',
    'caminhoes': 'Sim',
    'truck': 'Sim',
    'tacografos': 'Sim',
    'ar condicionado': 'Sim',
    'vulcanizadora': 'Sim',
    'oleos': 'Sim',
    'borracheiro': 'Sim',
    'borracharia': 'Sim',
    'motor': 'Sim',
    'valvula': 'Sim',
    'manutencao': 'Sim',
    'Mobi': 'Sim',
    'mobbi': 'Sim',
    'strada': 'Sim',
    'rodas': 'Sim',
    'motores': 'Sim',
    'Motores': 'Sim',
    'motores': 'Sim',
    'LAVA RAPIDO': 'Sim',


    # --- Não manutenção -> "Não" ---
    'chapa': 'Não',
    'fusion': 'Não',
    'informata': 'Não',
    'informatica': 'Não',
    'consultoria': 'Não',
    'tecnologia': 'Não',
    'papelaria': 'Não',
    'escritorio': 'Não',
    'contabilidade': 'Não',
    'advocacia': 'Não',
    'juridico': 'Não',
    'marketing': 'Não',
    'publicidade': 'Não',
    'alimentacao': 'Não',
    'alimentos': 'Não',
    'restaurante': 'Não',
    'lanchonete': 'Não',
    'hotel': 'Não',
    'hospedagem': 'Não',
    'viagem': 'Não',
    'passagem aerea': 'Não',
    'telefonia': 'Não',
    'internet': 'Não',
    'software': 'Não',
    'construcao civil': 'Não',
    'material de construcao': 'Não',
    'limpeza predial': 'Não',
    'seguranca patrimonial': 'Não',
    'vigilancia': 'Não',
    'recursos humanos': 'Não',
    'treinamento': 'Não',
    'educacao': 'Não',
    'grafica': 'Não',
    'uniformes': 'Não',
    'nota': 'Não',
    'tintas': 'Não',
    'magazine luiza': 'Não',
    'contrução': 'Não',
    'refigreracao': 'Não',
    'despachente': 'Não',
    'balança': 'Não',
    'balancas': 'Não',
    'balanças': 'Não',
    'swile': 'Não',
    'toledo': 'Não',
    'telecomunicações': 'Não',
    'telecomunicacoes': 'Não',
    'locação de equipamentos': 'Não',
    'locação': 'Não',
    'empilhadeira': 'Não',
    'empilhadeiras': 'Não',
    'placas': 'Não',
    'reclassificação': 'Não',
    'prestação': 'Não',
    'despesas operacionais': 'Não',
    'despesas administrativas': 'Não',
    'despesas vendedores': 'Não',
    'despesas semanais': 'Não',
    'material eletrico': 'Não',
    'materiais eletricos': 'Não',
    'material de construcao': 'Não',
    'EPI': 'Não',
    'Despachante': 'Não',
    'BALANCAS': 'Não',
    'BALANCA': 'Não',
    'N Faturar': 'Não',
    'Reclassificação': 'Não',
    'Reclassificacao': 'Não',
    'contas': 'Não',
    'transferecia': 'Não',
    'REEMBOLSO': 'Não',
    'Reembolso': 'Não',
    'Pesagem': 'Não',
    'N FATURA': 'Não',
    'Midia digitais': 'Não',
    'Midia digital': 'Não',
    'RESSARCIMENTO': 'Não',
    'duplicidade': 'Não',
    'edenred': 'Não',
    'reembolso': 'Não',
    'boleto': 'Não',
    'estacionamento': 'Não',
    'estorno': 'Não',
    'Estorno': 'Não',
    'estornos': 'Não',
    'provisao': 'Não',
    'Provisão': 'Não',
    'PROVISAO': 'Não',
    'LIMPEZA': 'Não',
    'fechamento':'Não',
    'fechamento de estoque':'Não',
    'ajuste de fechamento':'Não',
    'titulo':'Não',
    'ajuste':'Não',
    'multa':'Não',
    'desconto':'Não',
    'desc':'Não',
    'saldo':'Não',
    'BAIXA':'Não',
    'baixa':'Não',
    'Baixa':'Não',
    'saldo inicial':'Não',
    'locacao':'Não',
    'SINISTRO PLACA': 'Não',
    'Sinistro': 'Não',
    'Placa': 'Não',
    'erro lançamento': 'Não',
    'ERRO LANCAMENTO': 'Não',
    'Reeclassificacao': 'Não',
    'ASA RENT': 'Não',
    'metalurgica': 'Não',


    # 'outra palavra': 'Sim',
}

def classifica_por_palavra_chave(razao, regras=REGRAS_PALAVRA_CHAVE,
                                  texto_col='TEXTO_MANUTENCAO',
                                  manut_col='MANUTENCAO',
                                  prioridade=True):
    """
    Aplica regras de classificação por palavra-chave sobre a coluna de texto limpo.

    - regras: dict {palavra_chave: classificacao}, busca case-insensitive por
      substring dentro do texto limpo (TEXTO_MANUTENCAO).
    - prioridade=True: a regra de palavra-chave tem prioridade sobre o resultado
      do PROCV, ou seja, sobrescreve MANUTENCAO sempre que a palavra-chave for
      encontrada (útil para exceções conhecidas, como "automotiva" -> "Não").
    - prioridade=False: a regra de palavra-chave só é aplicada nos registros que
      ainda estão sem classificação (MANUTENCAO nulo/NaN) após o PROCV.
    """
    resultado = razao.copy()
    texto_lower = resultado[texto_col].astype(str).str.lower()

    for palavra, classificacao in regras.items():
        contem_palavra = texto_lower.str.contains(palavra.lower(), na=False, regex=False)

        if prioridade:
            alvo = contem_palavra
        else:
            alvo = contem_palavra & resultado[manut_col].isna()

        resultado.loc[alvo, manut_col] = classificacao

    return resultado

# aplica a classificação por palavra-chave (prioridade sobre o PROCV para as
# palavras-chave cadastradas; ajuste prioridade=False para só preencher os
# registros ainda não classificados)
razao = classifica_por_palavra_chave(razao, REGRAS_PALAVRA_CHAVE, prioridade=True)

# Resumo mensal
resumo = (
    razao
    .groupby('MES')
    .agg(
        Nao_Manutencao=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'] == 'Não'].sum()),
        Manutencao=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'] == 'Sim'].sum()),
        Nan=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'].isna()].sum()),
        Valor_Total=('VALOR_MOEDA', 'sum')
    )
    .reset_index()
    .cummax()
)

resumo

percentual = (
    razao
    .groupby('MES')
    .agg(
        Nao_manutencao_percentual=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'] == 'Não'].sum() / x.sum() * 100),
        Manutencao_percentual=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'] == 'Sim'].sum() / x.sum() * 100),
        Nan_percentual=('VALOR_MOEDA', lambda x: x[razao['MANUTENCAO'].isna()].sum() / x.sum() * 100),
        total_percentual=('VALOR_MOEDA', lambda x: x.sum() / x.sum() * 100)
    )
    .reset_index()
)

#print(razao.head(20))
print(resumo)
print(percentual)

razao.to_excel('espurgos_2026.xlsx', index=False)

#base.to_excel('base_tratada.xlsx', index=False)