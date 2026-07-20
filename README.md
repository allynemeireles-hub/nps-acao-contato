# nps-acao-contato

Skill ToqanClaw para gerar planilha de detratores NPS MetLife × iFood para ação de contato.

## O que faz

Filtra detratores NPS (nota 0–6) com comentários negativos a partir de uma data de corte, cruza com Databricks PII para trazer telefone de contato, e gera um arquivo Excel com 12 colunas:

`nota_nps | sinistro | distribuido | cobertura | status | mot_recusa | q2_comentario | q2_motivo | driver_uuid | nome | cpf | telefone`

## Instalação

```bash
npx skills add allynemeireles-hub/nps-acao-contato --skill nps-acao-contato --agent '*' --yes --copy
```


## Uso

A skill é acionada quando Allyne mencionar:
- "ação de contato dos detratores"
- "reclamações NPS"
- "detratores para contato"
- "planilha de contato NPS"
- "puxar detratores"
- "lista de detratores"
- "exportar reclamações NPS"

Basta informar a data de corte (ex: "desde 06/07/2026").


## Estrutura

```
nps-acao-contato/
├── SKILL.md                      # Definição da skill (ToqanClaw Skills CLI)
├── scripts/
│   └── gerar_planilha.py         # Script principal
└── README.md
```

## Fontes de dados

| Fonte | Detalhe |
|---|---|
| Google Sheets NPS | ID `1vy_33ywbjRY9dYxo_e5jOeNSBSlJofrFkGXAr6xKL-Q`, aba `NPS base full Metlife` |
| Databricks PII | `pii.pii_data_raw.driver_data` — colunas `cpf` e `phone` |

## Output

- Planilha: `data/nps-sinistros/reclamacoes_<MMDD>_<YYYY>.xlsx`
- Metadados: `data/nps-sinistros/reclamacoes_<MMDD>_<YYYY>_meta.json`
