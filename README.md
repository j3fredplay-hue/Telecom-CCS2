# Telecom-CCS2 — Limpeza de ONUs x IXC

`olt_ixc.py` cruza a lista de ONUs de uma OLT com os relatórios exportados do IXC e gera, por OLT:

- `<OLT>_limpeza.xlsx`: resumo, ONUs a **remover** (contrato desativado), ONUs a **conferir** (sem número, número inexistente no IXC, ambíguas) e a lista completa.
- `<OLT>_ativos.xlsx`: relatório para engenharia, com resumo por PON (cliente final x rede neutra, online x offline), ONUs por cliente/provedor e a lista de ativos.

## Uso

```bash
python3 olt_ixc.py dados/olts/OLT1-SHOPMILLENNIUM.txt \
  --logins dados/ixc/Relatorio_com_Login.xlsx \
  --contratos dados/ixc/Relatorio_de_contrato_do_IXC.xlsx \
  --sgp dados/ixc/Relatorio_Upwire_SGP.xlsx \
  --saida saida/
```

Formatos de OLT suportados: Datacom (saída de `show interface gpon 1/1/ onu`). VSOL e Intelbras EPON: a fazer.

## Regras

- O número após `CTR-` é buscado primeiro como **login PPPoE** e depois como **ID de contrato**.
- Ativo = status de acesso Ativo, Bloqueio Automático/Manual ou Financeiro em atraso.
- Rede neutra = plano de venda 67 (Circuito Dedicado); "provável" quando o mesmo número aparece em 3 ou mais ONUs.
- Parceiro Upwire (`CTR-201939-...`, contrato MARCELO SILVA DO CARMO): o status vem do relatório do SGP do parceiro (`--sgp`), pelo nº do contrato SGP (`CTR-201939-11202`) ou, sem número, pelo nome (`CTR-201939-MONICA`). Cancelado = remover, Suspenso = ativo, Inativo = pré-contrato. Nome ambíguo ou cliente com outro contrato ativo vai para "Conferir".
- Números do sistema antigo da CCS Brasil (`2019xxx`–`2026xxx`) são buscados como login no IXC; os que não têm login são mapeados em `NUMEROS_ANTIGOS` no script (ex.: `2023716` → contrato 538, ID HOUSE).
- Senhas e outros dados sensíveis do IXC nunca vão para a saída.

As pastas `dados/` e `saida/` ficam fora do git porque têm dados de clientes.
