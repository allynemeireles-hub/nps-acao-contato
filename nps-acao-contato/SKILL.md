---
name: nps-acao-contato
description: >
  Gera planilha de detratores NPS MetLife × iFood para ação de contato, com filtro de data de corte.
  Use quando a usuária mencionar: "ação de contato", "reclamações NPS", "detratores para contato",
  "planilha de contato NPS", "puxar detratores", "lista de detratores", "exportar reclamações NPS".
tags:
  - audience:allyne
  - domain:nps
  - domain:metlife
---

# Skill: nps-acao-contato

## Propósito

Gera uma planilha Excel com detratores NPS (nota 0–6) que possuem comentários negativos a partir de
uma data de corte informada. Cruza com Databricks PII para trazer telefone de contato.

## Acionamento

Quando a Allyne pedir:
- "ação de contato dos detratores"
- "reclamações NPS"
- "detratores para contato"
- "planilha de contato NPS"
- "puxar detratores"
- "lista de detratores"
- "exportar reclamações NPS"
- qualquer variação de "gerar planilha de detratores" + data

## Fluxo obrigatório

### 1. Extrair data de corte

Identificar a data no pedido da usuária. Exemplos:
- "a partir de 22/06" → `22/06/2026`
- "desde 01/07" → `01/07/2026`
- "do mês de junho" → `01/06/2026`

Se a data não for informada, perguntar antes de executar.

### 2. Executar o script

```bash
cd /workspace
python3 skills/nps-acao-contato/scripts/gerar_planilha.py --data-corte DD/MM/YYYY
```
Substituir `DD/MM/YYYY` pela data de corte extraída.

### 3. Retornar resultado

Após execução bem-sucedida, reportar:
- Caminho do arquivo gerado: `data/nps-sinistros/reclamacoes_MMDD_YYYY.xlsx`
- Número de detratores filtrados
- Período coberto (data mínima e máxima)
- Quantos têm telefone disponível

## Fontes de dados

| Fonte | Detalhe |
|---|---|
| Google Sheets NPS | ID `1vy_33ywbjRY9dYxo_e5jOeNSBSlJofrFkGXAr6xKL-Q`, aba `NPS base full Metlife` |
| Databricks PII | `pii.pii_data_raw.driver_data` — colunas `cpf` e `phone` |

## Colunas de saída (ordem exata)

`nota_nps` | `sinistro` | `distribuido` | `cobertura` | `status` | `mot_recusa` |
`q2_comentario` | `q2_motivo` | `driver_uuid` | `nome` | `cpf` | `telefone`

## Critérios de filtro

1. `distribuido` >= data de corte
2. `nota_nps` entre 0 e 6 (detratores)
3. Tem conteúdo em `q1_comentario` OU `q2_comentario`
4. Comentário contém ao menos uma keyword negativa (case-insensitive):
   recusado, negado, indeferido, não recebi, nao recebi, demora, sem resposta, desistir,
   péssimo, pessimo, horrível, horrivel, decepcionante, abandona, não funciona, nao funciona,
   problema, injusto, mentira, enganado, cancelado, não pago, nao pago, difícil, dificil,
   ruim, prejudic
5. **Fallback:** se não houver casos com keyword, incluir todos detratores com comentário

## Output

- Planilha: `data/nps-sinistros/reclamacoes_<MMDD>_<YYYY>.xlsx`
- Metadados: `data/nps-sinistros/reclamacoes_<MMDD>_<YYYY>_meta.json`

```json
{
  "total_detratores_filtrados": 12,
  "com_telefone": 9,
  "data_minima": "22/06/2026",
  "data_maxima": "30/06/2026",
  "gerado_em": "2026-07-02T14:30:00+00:00",
  "keyword_filter_applied": true
}
```

## Dependências técnicas

- Python 3 + `openpyxl` + `requests` (pré-instalados)
- Acesso ao Google Sheets via MCP (proxy injeta auth)
- Acesso ao Databricks via HTTP SQL API (proxy injeta auth)
- Cache NPS: `data/nps-painel/nps_data_cache.json` (usado como fallback se Sheets indisponível)

## Notas

- O script segue a arquitetura ToqanClaw: sem hardcode de credenciais, sem headers manuais.
- Se o Databricks PII não retornar dados, a coluna `telefone` fica vazia mas o arquivo é gerado.
- Ver `data/nps-sinistros/gen_reclamacoes_jun22.py` para referência de implementação original.
