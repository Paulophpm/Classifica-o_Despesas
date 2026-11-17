# Classificação de Naturezas Logísticas — Razão de Despesas

Classifica cada lançamento das contas logísticas do razão em uma **categoria**
(Peças, Pneus, Mão de obra terceirizada, Combustível, Frete, Pedágio...) e
compara o **grupo esperado** dessa categoria com o **grupo atual** (a conta
onde o lançamento caiu). Assim a análise mensal responde:

- quanto de **combustível está caindo em manutenção**;
- quanto de **manutenção está fora das contas de manutenção**;
- quanto do que está em manutenção **não é manutenção**;
- dentro de cada grupo, **quanto é de cada categoria** (ex.: mão de obra x peças x frete dentro de Manutenção).

## Como rodar (todo mês)

```bash
pip install -r requirements.txt
python criar_planilha_regras.py          # só na primeira vez: cria config/regras.xlsx
python analise_mensal.py --competencia 2026-10 --aba 2026
```

Entradas padrão (podem ser trocadas por parâmetro, veja `--help`):

| Arquivo | Conteúdo |
|---|---|
| `data/razao.xlsx` | Razão exportado do sistema (colunas obrigatórias: `DATA_TRANSACAO`, `NOME_CONTA`, `TEXTO_TRANSACAO`, `VALOR_MOEDA`) |
| `data/Transacoes_base.xlsx` | Base de referência do PROCV (`TRANSACAO`, `MANUTENCAO`, `ATIVIDADE`) — opcional |
| `config/regras.xlsx` | Premissas e regras (versionada no Git) |

Saída: `output/analise_logistica_AAAA-MM.xlsx`. Sem `--competencia`, usa o último mês do razão.

## Planilha de regras (`config/regras.xlsx`)

Tudo que é premissa fica aqui — **não é preciso mexer no código**.

| Aba | Para que serve |
|---|---|
| `PREMISSAS` | Parâmetros gerais: campos de texto usados, prioridades das camadas, exclusões por tipo de lançamento, valor mínimo de pendência, meses no comparativo |
| `CONTAS` | Contas analisadas → `GRUPO_ATUAL`. `CATEGORIA_PADRAO` (opcional) é usada quando nada mais classificou — ex.: tudo em *Pedágios* é Pedágio. As contas de manutenção ficam **sem** padrão de propósito |
| `CATEGORIAS` | Cada categoria → `GRUPO_ESPERADO` |
| `REGRAS` | Palavras-chave → categoria, com `PRIORIDADE` (menor vence), `TIPO_MATCH` e restrição opcional por grupo atual |
| `EXCECOES` | Classificação manual por texto (e opcionalmente por conta). É onde entram as pessoas físicas já identificadas — vale para todos os meses seguintes |
| `MAPA_ATIVIDADE` | De-para da coluna `ATIVIDADE` da base PROCV para as categorias |
| `PESSOA_FISICA` | Listas de NOME / SOBRENOME / EMPRESA para a detecção de PF |

### Camadas de classificação

A primeira camada que reconhecer o lançamento define a categoria:

| Prioridade | Camada | Confiança |
|---|---|---|
| 0 | Exceções manuais | Alta |
| 10–49 | Regras fortes (ajustes contábeis, combustível, multas, pedágio, IPVA, seguros, frete) | Alta |
| 50 | Base PROCV com `ATIVIDADE` mapeada | Alta |
| 200–499 | Regras específicas (peças, pneus, mão de obra, lavagem...) e não-logísticas | Média |
| 500 | Base PROCV só com Sim/Não | Baixa |
| 800–900 | Regras fracas/genéricas ("placa", "caminhão", "veículo") | Baixa |
| 990 | Pessoa física detectada → status **Revisar** | Baixa |
| 995 | Categoria padrão da conta | Baixa |

### Tipos de match

- `palavra` (padrão): palavra/expressão inteira — `MOTOR` não bate em `MOTORISTA`, `CAR` não bate em `RICARDO`.
- `prefixo`: começo de palavra — `PNEU` bate em `PNEUS`.
- `contem`: qualquer trecho do texto (comportamento do código antigo).
- `regex`: expressão regular.

O texto é comparado sem acento e em maiúsculas, **mantendo números** (dá para usar `ARLA 32`, `DIESEL S10`).

## Status

| Status | Significado |
|---|---|
| OK | Grupo esperado = grupo da conta |
| Divergente | Caiu em conta de outro grupo (a coluna `DIVERGENCIA` diz qual) |
| Revisar | Pessoa física, reembolso, despesas sem detalhe |
| Não classificado | Nenhuma camada reconheceu |
| Ajuste contábil | Estorno, provisão, reclassificação etc. |

## Abas do arquivo de saída

`RESUMO` · `MATRIZ_GRUPOS` (grupo atual × esperado) · `MATRIZ_CATEGORIAS` (categoria × grupo atual) ·
`CATEGORIA_POR_CONTA` · `DIVERGENCIAS` · `PENDENTES` · `COMPARATIVO_STATUS` · `COMPARATIVO_CATEGORIAS` ·
`BASE_CLASSIFICADA` · `CONFLITOS_BASE` · `ATIVIDADES_BASE` · `CONTAS_FORA_ESCOPO` · `VALIDACAO` · `PREMISSAS_USADAS`

### Ciclo mensal sugerido

1. Rodar a análise do mês.
2. Conferir `VALIDACAO` (totais batem com o razão? apareceu conta nova em `CONTAS_FORA_ESCOPO`?).
3. Trabalhar a aba `PENDENTES`: ela já vem agrupada por texto e no mesmo formato da aba `EXCECOES`
   (`TEXTO`, `NOME_CONTA`, `CATEGORIA`) — preencha a categoria e cole em `EXCECOES`.
   Para um termo que se repete, prefira criar uma regra em `REGRAS`.
4. Rodar de novo e analisar `RESUMO`, `MATRIZ_CATEGORIAS` e `DIVERGENCIAS`.
5. `PREMISSAS_USADAS` registra os parâmetros daquela rodada.

## Estrutura

```
analise_mensal.py            # ponto de entrada
criar_planilha_regras.py     # gera o modelo de config/regras.xlsx
classificacao/
    texto.py                 # limpeza de texto
    regras.py                # leitura e validação da planilha de regras
    classificador.py         # preparação do razão + camadas de classificação + status
    relatorios.py            # abas de saída e exportação
config/regras.xlsx
tests/                       # pytest
legado/                      # código anterior (Sim/Não de manutenção)
```

## Testes

```bash
python -m pytest
```

## Dados sensíveis

`data/` e `output/` não são versionados (ver `.gitignore`).
