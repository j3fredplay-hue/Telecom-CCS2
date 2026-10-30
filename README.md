# Telecom-CCS2 — Limpeza de ONUs x IXC (CCS Brasil)

`olt_ixc.py` cruza as ONUs de cada OLT com o IXC (CCS Brasil) e com o SGP da Upwire, e gera por OLT:

- `saida/<OLT>_limpeza.xlsx`: Resumo, **Remover**, **Conferir** e Todas.
- `saida/<OLT>_ativos.xlsx`: Resumo por PON, ONUs por cliente/provedor, lista de ativos e aba **POP** (equipamentos de `pops.json`).

## Como continuar em uma sessão nova

1. Anexe o arquivo `dados_backup.zip` (relatórios do IXC/SGP e as OLTs já processadas) e diga:
   > Continue o trabalho da branch `claude/olt-active-inactive-clients-4czwvq`: leia o README, descompacte o dados_backup.zip em `dados/` e processe a OLT anexa.
2. Anexe a OLT em `.txt` e informe nome da OLT + equipamentos do POP.
3. Responda às dúvidas de forma curta (ex.: "1 sim, 2 remover").

Se os relatórios do IXC/SGP forem atualizados, anexe os novos no lugar dos do zip.

## Comando

```bash
python3 olt_ixc.py dados/olts/<OLT>.txt \
  --logins dados/ixc/Relatorio_com_Login.xlsx \
  --contratos dados/ixc/Relatorio_de_contrato_do_IXC.xlsx \
  --sgp dados/ixc/Relatorio_Upwire_SGP.xlsx --saida saida
```
Depois rodar `recalc.py` (skill xlsx) nos Excel gerados para calcular as fórmulas.

## Formatos de OLT

- **Datacom**: `show interface gpon 1/1/ onu`.
- **Parks**: `show interface gponX/Y onu` **e** `show interface gponX/Y onu status` no mesmo `.txt` (status ACTIVE = online, INVALID = offline; também lê o sinal óptico).
- **VSOL (EPON)**: um `.txt` por OLT, com o nome da OLT no nome do arquivo; aba "ONU list" da web copiada para o bloco de notas (colunas separadas por TAB). O MAC entra no lugar do serial (coluna "Serial/MAC"). A aba "OPM Diag" pode vir logo abaixo no mesmo `.txt`: o RX Power vira o sinal óptico.

## Regras combinadas com o usuário

**Identificação**
- Número após `CTR-` (aceita `C-TR-`, `CRT-`, `CTR.`, `UPWIRE-CTR-`): primeiro como **login PPPoE** no IXC, depois como **ID de contrato**.
- Números do sistema antigo da CCS (`2019xxx`–`2026xxx`) sem login no IXC: mapa `NUMEROS_ANTIGOS` (2023716 → ID HOUSE 538; 2025928 → VISION 364).
- **Upwire** (provedor parceiro, no IXC é o contrato 388 MARCELO SILVA DO CARMO; nº antigo `201939`): `CTR-201939-<contrato SGP>` ou `CTR-201939-<NOME>` → status do cliente no **SGP**. Número de 5 dígitos sem o 201939 (ex.: `CTR-11267`, `11234-nome`) também é procurado no SGP.
- Descrição só com nome/login (padrão ex-Nova Solutec, ex.: `joedizalima`): busca como login no IXC e no SGP; depois pelo nome (primeiro nome + sobrenomes/iniciais). Mapa `NOMES_CONHECIDOS` para casos confirmados (BOBS, CPDE, STAFF, ID HOUSE, NERDS, FECHACOM…).
- Rede neutra: plano 67 ou contrato na lista `PARCEIROS_EXTRA` (relatório "Porta - Rede Neutra" do IXC).

**Classificação**
- Ativo = Ativo, Bloqueio Automático/Manual, Financeiro em atraso (IXC); Ativo/Suspenso (SGP). Inativo no SGP = pré-contrato (mantém).
- Desativado (IXC) / Cancelado (SGP) = **remover**.
- ONU de **parceiro (rede neutra)** nunca vai para remoção: o controle é do parceiro (contrato ativo = Ativo; senão, Conferir). Exceção: clientes da **Upwire** (contrato 388), controlados pela CCS via SGP; e ONU com o **mesmo serial/MAC online** em outra ONU (nesta ou em outra OLT): a offline vai para remoção.
- Mesmo cliente com um contrato ativo e outro cancelado = ativo.
- Cliente **não encontrado ou ambíguo**: offline = **remover**; online = **conferir**.
- ONU offline cujo mesmo login/cliente (ou serial, em outra OLT) está online em outra ONU = ONU antiga, **remover**.
- Decisões pontuais: `DECISOES_MANUAIS` (por serial + descrição) e `REMOVER_SEM_CONTRATO` (Tintão).
- Os arquivos não mencionam o engenheiro.

## Dados

`dados/` e `saida/` ficam fora do git (dados de clientes; senhas do IXC nunca vão para a saída). `pops.json` (equipamentos dos POPs) fica no repositório.
