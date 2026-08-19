# Classificação de Manutenção x Não Manutenção — Frota

Script em Python que classifica lançamentos do razão contábil (contas de
**Serviços para Logística e Frota**, **Peças e Acessórios** e **Peças e
Acessórios para Logística e Frota**) em **Manutenção** ou **Não Manutenção**,
combinando um cruzamento (PROCV) contra uma base de referência com regras de
palavra-chave configuráveis.

Hoje o script classifica automaticamente cerca de **80% das transações**; o
restante costuma ser lançamentos em nome de pessoa física ou fornecedores
que não dá para identificar só pelo texto da transação.

## Como funciona

1. **Leitura**: importa o razão contábil (`data/razao.xlsx`, aba `2026`) e a
   base de referência (`data/Transacoes_base.xlsx`).
2. **Filtro**: mantém apenas os lançamentos das contas de manutenção de frota.
3. **Limpeza de texto**: remove acentos, números e caracteres especiais do
   texto da transação, para melhorar o cruzamento.
4. **PROCV (merge)**: cruza o texto limpo de cada transação com a base de
   referência para preencher a classificação `MANUTENCAO` (Sim/Não).
5. **Regras de palavra-chave**: aplica um dicionário de palavras-chave
   (`src/config.py`) como camada complementar e prioritária — cobre casos que
   o PROCV não resolve e permite corrigir exceções sem mexer no código.
6. **Resumo mensal**: agrega os valores por mês, com totais e percentuais de
   Manutenção, Não Manutenção e não classificados.
7. **Exportação**: salva o resultado detalhado em `output/espurgos_2026.xlsx`.

## Estrutura do repositório

```
classificacao-manutencao-frota/
├── data/                   # arquivos de entrada (não versionados)
│   └── .gitkeep
├── output/                 # arquivos gerados (não versionados)
│   └── .gitkeep
├── src/
│   ├── main.py              # fluxo principal (leitura, limpeza, PROCV, resumo, exportação)
│   └── config.py            # dicionário de regras de palavra-chave
├── requirements.txt
├── .gitignore
└── README.md
```

## Pré-requisitos

- Python 3.10+
- Bibliotecas listadas em `requirements.txt`

## Instalação

```bash
git clone https://github.com/<seu-usuario>/classificacao-manutencao-frota.git
cd classificacao-manutencao-frota
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

1. Coloque os arquivos de entrada em `data/`:
   - `razao.xlsx` (com a aba `2026`)
   - `Transacoes_base.xlsx`
2. Rode o script:

```bash
python src/main.py
```

3. O resultado detalhado sai em `output/espurgos_2026.xlsx`, e o resumo
   mensal (valores e percentuais) é impresso no terminal.

## Ajustando as regras de classificação

As palavras-chave usadas na classificação ficam em `src/config.py`, no
dicionário `REGRAS_PALAVRA_CHAVE`. Cada entrada é `"palavra": "Sim"` (é
manutenção) ou `"palavra": "Não"` (não é manutenção). A busca:

- ignora acentos e maiúsculas/minúsculas (o texto já é limpo antes de comparar);
- é por substring, então `"pecas"` cobre `"AUTO PECAS"`, `"PECAS LTDA"` etc.;
- por padrão tem **prioridade sobre o PROCV** — ou seja, se uma palavra-chave
  bater, ela sobrescreve o que veio da base de referência. Isso é útil para
  tratar exceções conhecidas sem precisar editar a base.

Para adicionar um novo caso, basta incluir uma linha no dicionário — não é
necessário alterar `main.py`.

## Dados sensíveis

Os arquivos de entrada e saída contêm informações contábeis internas e
**não são versionados** (ver `.gitignore`). Ao clonar o repositório, gere ou
copie seus próprios arquivos para `data/` antes de rodar o script.
