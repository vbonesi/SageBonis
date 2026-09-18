# Exemplo de validação — extração reversa de IEDs + comando analógico

> **Codificação:** os arquivos `.dat` desta pasta estão em ISO-8859-1 (`latin-1`),
> como esperado pelo SAGE. Visualizadores que presumem UTF-8 podem mostrar acentos
> como `�`, embora os bytes estejam íntegros. Para inspecioná-los em terminal UTF-8,
> use `iconv -f latin1 -t utf8 arquivo.dat`.

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
  real (anonimizada): DNP3 (aquisição `EX1` + distribuição `EX1_LSC`/`COR_LSC`),
  61850 (7 IEDs), MODBUS (`EX14S8`) e SNMP (16 dispositivos). Confira `Protocolo`,
  `Direcao`, e os campos de `CNF.CONFIG` reconstruídos (ex.: `PlPr/LiPr/PlRe/LiRe`
  pro DNP3, `ApTitle/OPMSK` pro 61850, `VERSAO/HOST/COMMUNITY` pro SNMP).
- **Aba `PontoAnalogico`** — procure a linha `ID_Logico = EX1:REGU-STPS`
  ("Valor Scan de Regulacao da Barra 69KV-EX1"): `Comando = S`,
  `ID_Fisico_Comando = EX1_CDNP_2_CSTP_0`, `LMI1C = 680`, `LMS1C = 715` — os
  limites numéricos (`LMI1C`/`LMS1C`) vieram de um setpoint real, só o
  identificador da subestação/IED foi trocado (não é um valor fabricado do
  zero).
- **Aba `ComandoAvulso`** — os demais `CGS` que não são o setpoint acima (17
  linhas), incluindo comandos ligados a um ponto genérico compartilhado.

## De onde veio

Fragmento de uma base real de cliente. Este repositório é público, então o
código da subestação, um sufixo de projeto, o nome de um IED e os 16 IPs
internos reais foram trocados por tokens/faixa genéricos (`192.0.2.0/24`,
reservada pra documentação pela RFC 5737). Fora esses tokens, mantém a estrutura da base
original: tudo que já estava no nível raiz (`bd/dados/*.dat`, sem seguir os
~57 `#include` de instalação) mais só **3 subpastas** escolhidas por
conterem cada protocolo: `SNMP/`, `RDP/` (61850) e `coringa4S8/` (MODBUS). A
única mudança de conteúdo além da anonimização foi reativar o `#include` de
`coringa4S8` (estava comentado/desativado na base de origem — é o único
exemplo do acervo com `TTP=TMBUS`, o valor que bate com o protocolo MODBUS
tal como modelado aqui; os outros 2 exemplos reais disponíveis usam
`TTP=SMBUS`, uma variante ainda não coberta).

`dados/` tem o fragmento `.dat` puro, já anonimizado (útil se quiser
reimportar do zero). Não reintroduzir dado real de cliente aqui — qualquer
atualização futura desta pasta deve passar pela mesma anonimização antes do
commit.

## O que NÃO validar aqui

Esta planilha **não** é uma cópia da base de produção real do usuário
(`completa/SageBonis.ods`) — as abas de entidade foram completamente
substituídas por este fragmento pequeno ao importar. Não usar para nada além
de conferir visualmente as duas funcionalidades acima.
