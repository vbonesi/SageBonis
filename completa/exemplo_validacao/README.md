# Exemplo de validação — extração reversa de IEDs + comando analógico

`SageBonis_exemplo_validacao.ods` é uma planilha pequena, já processada
(`importar_dats` + `extrair_pontos` já rodados), pra você conferir visualmente
as duas funcionalidades mais recentes:

- **Extração reversa de IEDs** (aba `IEDs`) — reconstrói `Protocolo`/`Direcao`/
  `CNF.CONFIG`/limites a partir de `LSC`/`CNF`/`CXU`/`UTR`/`ENU`/`TAC`/`MUL`/`ENM`
  já importados.
- **Comando para pontos analógicos (setpoint)** (aba `PontoAnalogico`) — coluna
  `Comando`/`ID_Fisico_Comando`/`KCONV_Comando`/`LMI1C..LMS2C`.

## Onde olhar

- **Aba `IEDs`** — 27 linhas, cobrindo **4 protocolos** reconhecidos numa base
  real: DNP3 (aquisição `JDM` + distribuição `JDM_LSC`/`COR_LSC`), 61850 (7
  IEDs), MODBUS (`JDM4S8`) e SNMP (16 dispositivos). Confira `Protocolo`,
  `Direcao`, e os campos de `CNF.CONFIG` reconstruídos (ex.: `PlPr/LiPr/PlRe/LiRe`
  pro DNP3, `ApTitle/OPMSK` pro 61850, `VERSAO/HOST/COMMUNITY` pro SNMP).
- **Aba `PontoAnalogico`** — procure a linha `ID_Logico = JDM:REGU-STPS`
  ("Valor Scan de Regulacao da Barra 69KV-JDM"): `Comando = S`,
  `ID_Fisico_Comando = JDM_CDNP_2_CSTP_0`, `LMI1C = 680`, `LMS1C = 715` — um
  setpoint de tensão real, com limites numéricos de verdade (não um valor
  fabricado).
- **Aba `ComandoAvulso`** — os demais `CGS` que não são o setpoint acima (17
  linhas), incluindo comandos ligados a um ponto genérico compartilhado.

## De onde veio

Fragmento **real** (não sintético) extraído da base `jdm` (CHESF) do acervo
`~/Drive/Projetos/SkillSAGE/10_extraidas/jdm` — mantém tudo que já está no
nível raiz da base original (`bd/dados/*.dat`, sem seguir os ~57 `#include` de
instalação) mais só **3 subpastas** escolhidas por conterem cada protocolo:
`SNMP/`, `RDP/` (61850) e `coringa4S8/` (MODBUS). O conteúdo em si é cópia
literal — a única mudança deliberada foi reativar o `#include` de
`coringa4S8` (estava comentado/desativado na base de origem — é o único
exemplo do acervo com `TTP=TMBUS`, o valor que bate com o protocolo MODBUS tal
como modelado aqui; os outros 2 exemplos reais disponíveis usam `TTP=SMBUS`,
uma variante ainda não coberta).

`dados/` tem o fragmento `.dat` puro (útil se quiser reimportar do zero ou
conferir contra o original). Script de poda: pergunte ao Claude pelo histórico
desta sessão, ou refaça manualmente comparando `dados/*.dat` com a base
original.

## O que NÃO validar aqui

Esta planilha **não** é uma cópia da base de produção real do usuário
(`completa/SageBonis.ods`) — as abas de entidade foram completamente
substituídas por este fragmento pequeno ao importar. Não usar para nada além
de conferir visualmente as duas funcionalidades acima.
