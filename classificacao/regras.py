"""
Leitura e validação da planilha de regras (config/regras.xlsx).

A planilha é a "fonte da verdade" das premissas: contas analisadas, grupos,
categorias, palavras-chave, exceções manuais e detecção de pessoa física.
Nada disso fica fixo no código.
"""

from dataclasses import dataclass, field

import pandas as pd

from .texto import limpa_texto_regra, normaliza_chave

# Grupos com tratamento especial no status
GRUPO_AJUSTE = "Ajuste contábil"
GRUPO_VERIFICAR = "A verificar"
GRUPO_NAO_MANUT = "Não manutenção (a detalhar)"

# Categorias usadas quando a base PROCV só diz Sim/Não (ATIVIDADE sem de-para)
CATEGORIA_BASE_SIM = "Manutenção - não detalhado"
CATEGORIA_BASE_NAO = "Não manutenção - não detalhado"
CATEGORIA_PF = "Pessoa física (a identificar)"

TIPOS_MATCH = {"palavra", "prefixo", "contem", "regex"}

ABAS_OBRIGATORIAS = {
    "PREMISSAS": ["PARAMETRO", "VALOR"],
    "CONTAS": ["NOME_CONTA", "GRUPO_ATUAL", "CATEGORIA_PADRAO"],
    "CATEGORIAS": ["CATEGORIA", "GRUPO_ESPERADO"],
    "REGRAS": ["ATIVO", "PRIORIDADE", "TERMO", "TIPO_MATCH", "CATEGORIA", "SOMENTE_GRUPO_ATUAL"],
    "EXCECOES": ["TEXTO", "NOME_CONTA", "CATEGORIA"],
    "MAPA_ATIVIDADE": ["ATIVIDADE", "CATEGORIA"],
    "PESSOA_FISICA": ["TERMO", "TIPO"],
}

PREMISSAS_PADRAO = {
    "CAMPOS_TEXTO": "TEXTO_TRANSACAO",
    "GRUPO_MANUTENCAO": "Manutenção",
    "GRUPO_COMBUSTIVEL": "Combustíveis e Lubrificantes",
    "PRIORIDADE_EXCECAO": 0,
    "PRIORIDADE_BASE_ATIVIDADE": 50,
    "PRIORIDADE_BASE_SIM_NAO": 500,
    "PRIORIDADE_PF": 990,
    "PRIORIDADE_PADRAO_CONTA": 995,
    "DETECTAR_PF": "SIM",
    "APLICAR_PADRAO_CONTA": "SIM",
    "EXCLUIR_TIPO_LANCAMENTO": "",
    "EXCLUIR_TIPO_TRANSACAO": "",
    "VALOR_MINIMO_PENDENTE": 0,
    "MESES_COMPARATIVO": 6,
}


class ErroRegras(ValueError):
    pass


def _texto(v) -> str:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return str(v).strip()


def _sim(v) -> bool:
    return normaliza_chave(v) in {"sim", "s", "1", "true", "x", "verdadeiro"}


def _lista(v) -> list[str]:
    return [p.strip() for p in _texto(v).replace(",", ";").split(";") if p.strip()]


@dataclass
class Regras:
    premissas: dict
    contas: pd.DataFrame          # CHAVE_CONTA, NOME_CONTA, GRUPO_ATUAL, CATEGORIA_PADRAO
    categorias: dict              # categoria -> grupo esperado
    regras: pd.DataFrame          # regras ativas, já normalizadas e ordenadas
    excecoes: pd.DataFrame        # TEXTO_REGRA, CHAVE_CONTA, CATEGORIA
    mapa_atividade: dict          # chave atividade -> categoria
    pf_nomes: set = field(default_factory=set)
    pf_sobrenomes: set = field(default_factory=set)
    pf_empresa: list = field(default_factory=list)

    # atalhos
    @property
    def grupo_manutencao(self) -> str:
        return self.premissas["GRUPO_MANUTENCAO"]

    @property
    def grupo_combustivel(self) -> str:
        return self.premissas["GRUPO_COMBUSTIVEL"]

    def premissa_int(self, nome) -> int:
        try:
            return int(float(self.premissas[nome]))
        except (TypeError, ValueError):
            raise ErroRegras(f"Premissa {nome} deveria ser número, veio: {self.premissas[nome]!r}")

    def premissa_float(self, nome) -> float:
        try:
            return float(self.premissas[nome])
        except (TypeError, ValueError):
            raise ErroRegras(f"Premissa {nome} deveria ser número, veio: {self.premissas[nome]!r}")

    def premissa_sim(self, nome) -> bool:
        return _sim(self.premissas[nome])

    def premissa_lista(self, nome) -> list[str]:
        return _lista(self.premissas[nome])


def carrega_regras(caminho) -> Regras:
    try:
        abas = pd.read_excel(caminho, sheet_name=None, dtype=object)
    except FileNotFoundError:
        raise ErroRegras(
            f"Planilha de regras não encontrada: {caminho}. "
            "Rode `python criar_planilha_regras.py` para gerar o modelo."
        )

    erros = []
    for aba, colunas in ABAS_OBRIGATORIAS.items():
        if aba not in abas:
            erros.append(f"Aba '{aba}' não existe na planilha de regras.")
            continue
        df = abas[aba]
        df.columns = [str(c).strip().upper() for c in df.columns]
        faltando = [c for c in colunas if c not in df.columns]
        if faltando:
            erros.append(f"Aba '{aba}': faltam as colunas {faltando}.")
    if erros:
        raise ErroRegras("\n".join(erros))

    # --- PREMISSAS ---
    premissas = dict(PREMISSAS_PADRAO)
    for _, linha in abas["PREMISSAS"].iterrows():
        nome = _texto(linha["PARAMETRO"]).upper()
        if nome:
            premissas[nome] = linha["VALOR"] if not pd.isna(linha["VALOR"]) else ""

    # --- CATEGORIAS ---
    categorias = {}
    for _, linha in abas["CATEGORIAS"].iterrows():
        cat, grupo = _texto(linha["CATEGORIA"]), _texto(linha["GRUPO_ESPERADO"])
        if not cat:
            continue
        if not grupo:
            erros.append(f"CATEGORIAS: categoria '{cat}' está sem GRUPO_ESPERADO.")
        if cat in categorias and categorias[cat] != grupo:
            erros.append(f"CATEGORIAS: categoria '{cat}' aparece com dois grupos diferentes.")
        categorias[cat] = grupo
    # categorias internas (sempre existem)
    categorias.setdefault(CATEGORIA_BASE_SIM, premissas["GRUPO_MANUTENCAO"])
    categorias.setdefault(CATEGORIA_BASE_NAO, GRUPO_NAO_MANUT)
    categorias.setdefault(CATEGORIA_PF, GRUPO_VERIFICAR)

    def valida_categoria(cat, onde):
        if cat not in categorias:
            erros.append(f"{onde}: categoria '{cat}' não está cadastrada na aba CATEGORIAS.")

    # --- CONTAS ---
    contas = abas["CONTAS"].copy()
    contas["NOME_CONTA"] = contas["NOME_CONTA"].map(_texto)
    contas = contas[contas["NOME_CONTA"] != ""]
    contas["GRUPO_ATUAL"] = contas["GRUPO_ATUAL"].map(_texto)
    contas["CATEGORIA_PADRAO"] = contas["CATEGORIA_PADRAO"].map(_texto)
    contas["CHAVE_CONTA"] = contas["NOME_CONTA"].map(normaliza_chave)
    duplicadas = contas[contas["CHAVE_CONTA"].duplicated()]["NOME_CONTA"].tolist()
    if duplicadas:
        erros.append(f"CONTAS: contas repetidas: {duplicadas}")
    for _, linha in contas.iterrows():
        if not linha["GRUPO_ATUAL"]:
            erros.append(f"CONTAS: conta '{linha['NOME_CONTA']}' está sem GRUPO_ATUAL.")
        if linha["CATEGORIA_PADRAO"]:
            valida_categoria(linha["CATEGORIA_PADRAO"], f"CONTAS ({linha['NOME_CONTA']})")
    contas = contas[["CHAVE_CONTA", "NOME_CONTA", "GRUPO_ATUAL", "CATEGORIA_PADRAO"]]

    # --- REGRAS ---
    regras = abas["REGRAS"].copy()
    regras = regras[regras["TERMO"].map(_texto) != ""]
    regras = regras[regras["ATIVO"].map(_sim)].copy()
    regras["TERMO_ORIGINAL"] = regras["TERMO"].map(_texto)
    regras["TIPO_MATCH"] = regras["TIPO_MATCH"].map(lambda v: normaliza_chave(v) or "palavra")
    regras["CATEGORIA"] = regras["CATEGORIA"].map(_texto)
    regras["SOMENTE_GRUPO_ATUAL"] = regras["SOMENTE_GRUPO_ATUAL"].map(_lista)
    regras["PRIORIDADE"] = pd.to_numeric(regras["PRIORIDADE"], errors="coerce")
    for idx, linha in regras.iterrows():
        onde = f"REGRAS (termo '{linha['TERMO_ORIGINAL']}')"
        if pd.isna(linha["PRIORIDADE"]):
            erros.append(f"{onde}: PRIORIDADE vazia ou não numérica.")
        if linha["TIPO_MATCH"] not in TIPOS_MATCH:
            erros.append(f"{onde}: TIPO_MATCH '{linha['TIPO_MATCH']}' inválido (use {sorted(TIPOS_MATCH)}).")
        valida_categoria(linha["CATEGORIA"], onde)
    # regex usa o termo como está; os demais tipos comparam com o texto limpo
    regras["TERMO"] = [
        t if tipo == "regex" else limpa_texto_regra(t)
        for t, tipo in zip(regras["TERMO_ORIGINAL"], regras["TIPO_MATCH"])
    ]
    # dentro da mesma prioridade, o termo mais longo (mais específico) vem primeiro
    regras["_TAM"] = regras["TERMO"].str.len()
    regras = regras.sort_values(["PRIORIDADE", "_TAM"], ascending=[True, False]).drop(columns="_TAM")
    regras = regras.reset_index(drop=True)

    # --- EXCECOES ---
    exc = abas["EXCECOES"].copy()
    exc["TEXTO_REGRA"] = exc["TEXTO"].map(limpa_texto_regra)
    exc = exc[exc["TEXTO_REGRA"] != ""].copy()
    exc["CHAVE_CONTA"] = exc["NOME_CONTA"].map(normaliza_chave)
    exc["CATEGORIA"] = exc["CATEGORIA"].map(_texto)
    for _, linha in exc.iterrows():
        valida_categoria(linha["CATEGORIA"], f"EXCECOES (texto '{linha['TEXTO']}')")
    exc = exc[["TEXTO_REGRA", "CHAVE_CONTA", "CATEGORIA"]].drop_duplicates(
        subset=["TEXTO_REGRA", "CHAVE_CONTA"], keep="last"
    )

    # --- MAPA_ATIVIDADE ---
    mapa = {}
    for _, linha in abas["MAPA_ATIVIDADE"].iterrows():
        atv, cat = normaliza_chave(linha["ATIVIDADE"]), _texto(linha["CATEGORIA"])
        if atv and cat:
            valida_categoria(cat, f"MAPA_ATIVIDADE ('{linha['ATIVIDADE']}')")
            mapa[atv] = cat

    # --- PESSOA_FISICA ---
    pf = abas["PESSOA_FISICA"]
    nomes, sobrenomes, empresa = set(), set(), []
    for _, linha in pf.iterrows():
        termo, tipo = limpa_texto_regra(linha["TERMO"]), normaliza_chave(linha["TIPO"])
        if not termo:
            continue
        if tipo == "nome":
            nomes.add(termo)
        elif tipo == "sobrenome":
            sobrenomes.add(termo)
        elif tipo == "empresa":
            empresa.append(termo)
        else:
            erros.append(f"PESSOA_FISICA: tipo '{linha['TIPO']}' inválido (use NOME, SOBRENOME ou EMPRESA).")

    if erros:
        raise ErroRegras("Problemas na planilha de regras:\n- " + "\n- ".join(erros))

    return Regras(
        premissas=premissas,
        contas=contas.reset_index(drop=True),
        categorias=categorias,
        regras=regras,
        excecoes=exc.reset_index(drop=True),
        mapa_atividade=mapa,
        pf_nomes=nomes,
        pf_sobrenomes=sobrenomes,
        pf_empresa=empresa,
    )
