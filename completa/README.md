# SageBonis — Trilha Completa (em desenvolvimento)

Esta é a variante **Completa** do SageBonis: um fork da planilha/macro
[Simples](../README.md) que mantém todo o import/export e vai acumulando recursos
avançados. A estratégia das duas trilhas está em [PLANEJAMENTO.md](../PLANEJAMENTO.md).

> A planilha **Simples** (na raiz do repo) continua sendo a recomendada para quem
> quer só importar/exportar base rápido. Use a Completa se quiser os recursos abaixo.

## Recursos além da Simples

Todas as funções abaixo estão no menu **SageBonis** e na toolbar do documento (junto
com as da Simples, separadas por um divisor). As abas novas seguem a mesma paleta
quente já usada na planilha (entidades de dado em vermelhos saturados; cada aba de
config/meta com seu próprio tom pastel, como `Geral`/`Cores`/`MaisUsadas`/
`EntidadeAtributoValor` já faziam): **salmão pastel** = config (você preenche),
**terroso** = relatório (a ferramenta escreve). A aba `Análise` foge disso de
propósito — usa vermelho/verde dinâmico (erro/limpo), informação mais útil que uma
cor fixa. (Só vale para abas criadas daqui
pra frente; `VerificacaoRefs`, criada antes dessa convenção existir, fica sem cor.)

### Verificador de base (linter) — `verificar_base`
Roda checagens de integridade na base e escreve um relatório na aba **`Análise`**
(criada automaticamente), com colunas `Severidade | Entidade | Linha | Atributo |
Valor | Descrição` e os erros no topo. A cor da aba fica **vermelha** se houver erro
e **verde** se estiver limpa. É **read-only** sobre as entidades — nunca altera
pontos, então não afeta a exportação.

Checagens atuais:
- **ID vazio** em ponto ativo (`Gera = x`);
- **ID duplicado** dentro da mesma entidade;
- **ID acima do tamanho conhecido** para a entidade (aviso; ver `LIMITES_TAMANHO_ID` no código —
  restrição documentada pelo projeto SkillSAGE, ainda não conferida contra o manual CEPEL);
- **Valor fora do domínio conhecido** — reaproveita a aba já existente `EntidadeAtributoValor`
  (`Entidade | Atributo | Valor1 | Valor2 | ...`) para avisar quando um atributo tem um valor
  que não está entre os valores válidos catalogados (ex.: `cgs.TPCTL` só aceita
  `AFIC/CSAC/CSCD/DFIC`);
- **Integridade referencial cruzada** — dirigida pela aba de config `VerificacaoRefs`.

#### Configurando a integridade referencial (`VerificacaoRefs`)
Na primeira execução do verificador, a aba `VerificacaoRefs` é criada com **81 regras
pré-populadas** (todas **inativas**), curadas a partir do levantamento de entidades e
relacionamentos do projeto SkillSAGE (base de conhecimento SAGE do autor; grafo de FK derivado
dos manuais CEPEL + bases reais). Cada linha é uma regra "o atributo X da entidade de origem
deve existir como atributo Y na(s) entidade(s) de destino":

| EntidadeOrigem | AtributoOrigem | EntidadeDestino | AtributoDestino | Ativa |
|----------------|----------------|-----------------|-----------------|-------|
| PDS | TAC | TAC | ID | N |
| PDD | PDS | PDS | ID | N |
| PDF | PNT | PDS\|PDD | ID | N |

A coluna `EntidadeDestino` aceita **múltiplas entidades separadas por `\|`** para modelar FK
"ambígua" do SAGE — ex.: `PDF.PNT` pode ser um `PDS` (ponto em aquisição) ou um `PDD` (ponto em
distribuição); a checagem valida contra a **união** dos IDs das duas. Isso elimina uma classe
inteira de falso-positivo que existia antes (quando só dava para modelar um destino por regra).

Ajuste as regras conforme o padrão da **sua** base e mude `Ativa` para `S` nas que
quiser ligar. Só regras ativas são checadas. (Por isso a 1ª execução não gera erros
de referência — você ativa o que faz sentido.)

> ⚠️ **Base parcial:** ao checar relacionamentos numa base **incompleta**, ainda é normal
> aparecerem `ERRO`s de "Referência não encontrada" quando o destino está numa parte da base
> que não foi importada. Não é bug: numa base completa esses casos somem. Ative os
> relacionamentos preferencialmente ao verificar a base inteira.
>
> ⚠️ **Planilhas já existentes:** se a sua `VerificacaoRefs` já foi criada por uma versão
> anterior desta ferramenta (3 regras de exemplo), ela **não é atualizada automaticamente** —
> a aba só é (re)criada quando não existe, para nunca sobrescrever edições suas. Para herdar as
> 81 regras curadas, apague a aba `VerificacaoRefs` e rode `verificar_base` de novo (isso inclui
> o `SageBonis.ods` deste próprio repositório, que ainda tem as 3 regras antigas).

### Unificação de pontos (fan-out) — `unificar_pontos`
Gera as entidades relacionadas (PDF/PDS/PDD, PAF/PAS/PAD, CGF/CGS) a partir de uma
definição única de ponto físico, em vez de escrever cada entidade manualmente. Escreve
direto nas abas de entidade existentes por **upsert** (casa por ID): regenerar depois
de ajustar a config atualiza as linhas, não duplica, e preserva colunas/linhas que não
vêm daqui (ex.: pontos importados de `.dat` reais na mesma aba).

Abas de config (criadas automaticamente, vazias, na 1ª execução):

- **`PontoDigital`** / **`PontoAnalogico`** — uma linha por **origem física** de um
  ponto (`ID_Logico, ID_Fisico, NOME, NV2, KCONV(s), TAC, OCR, Gera`). O `ID_Fisico` é
  quem você decide — endereçamento por protocolo (61850, DNP3, 101/104...) é escopo do
  futuro assistente de protocolo/IED; aqui só entra o fan-out a partir dele.
  - **Redundância**: várias linhas com o **mesmo `ID_Logico`** = múltiplas origens do
    mesmo ponto lógico. 1 origem → PDF/PAF direto (`TPFIL=NLFL`); 2+ origens → um
    PDF/PAF por origem + `RFC` em cadeia (fan-in "ou válido") + PDS/PAS com
    `TPFIL=FIL5`. Não assume nenhuma topologia fixa de IED (P/D/virtual, bit OPMSK
    12...) — só o número de origens que você declarar.
  - **Comando**: coluna `Comando=S` gera CGF/CGS com o **mesmo ID** do PDS/PAS (regra
    fixa do SAGE: comando e status compartilham o ID). Se 2+ origens tiverem
    `Comando=S`, gera um CGF por origem, todos referenciando o mesmo CGS.
    Em `PontoAnalogico` o comando é um **setpoint**: colunas extras
    `ID_Fisico_Comando`/`KCONV_Comando` (endereço físico do comando, mesma convenção
    do `PontoDigital`) e `LMI1C/LMI2C/LMS1C/LMS2C` (limites inferior/superior do
    comando, direto pro `CGS`). Achado real (`CGS.TIPO=PAS`) confirmado em 6 bases
    independentes, com 2 variantes: comando de TAP (2 estados, tipo Aumentar/
    Diminuir — limites ficam vazios) e setpoint numérico de verdade (limites
    preenchidos, ex. `tucurui`/`jdm`). Modelado como um único design genérico pras
    duas variantes — a distinção semântica exata entre `LMI1C` e `LMI2C` não foi
    confirmada o bastante pra virar 2 campos diferentes (ver `PLANEJAMENTO.md`).
- **`ComandoAvulso`** — comandos **sem** ponto de status próprio (ex.: um `COM_SAGE`
  genérico ligado a um TAC local, como algumas bases já usam). Cada linha tem seu
  próprio `ID` de CGS/CGF; várias linhas podem repetir o mesmo `TAC`/`PAC` (o ponto
  genérico) — é justamente o caso de vários comandos ligados ao mesmo ponto. Também
  aceita `LMI1C/LMI2C/LMS1C/LMS2C` (caso avulso analógico, ex. um setpoint dummy sem
  status próprio — achado real `ur_mir`); ficam vazios/sem uso no caso avulso digital.
- **`CanaisDistribuicao`** — um canal de saída por linha (`Nome, TDD, Metodo
  [Prefixo/Sufixo/Substituir/Explicito], Valor1, Valor2, Ativo`). Substitui os "4
  slots fixos" que macros de referência (GE) hard-codificam por quantos canais
  fizerem sentido. **`Explicito`**: achado real (algumas bases distribuem com um ID
  totalmente independente do `ID_Logico`, sem prefixo/sufixo nenhum) — nesse caso o
  ID final vem de `IDExplicito` em `DistribuicaoPontos`, não de uma transformação.
- **`DistribuicaoPontos`** — liga um `ID_Logico` a 1+ canais (`ID_Logico, Canal,
  IDExplicito, Ativo`). `IDExplicito` só é lido quando o canal usa
  `Metodo=Explicito` (fica vazio/ignorado nos demais casos). Um ponto sem nenhuma
  linha aqui simplesmente não gera distribuição.

#### Extração reversa (bases já existentes) — `extrair_pontos`
`unificar_pontos` sozinho só ajuda com **pontos novos**: uma base real já importada
(PDF/PDS/PAF/PAS/CGF/CGS vindos de `.dat`) não aparece em `PontoDigital`/
`PontoAnalogico` automaticamente. `extrair_pontos` é o espelho — lê as entidades já
existentes e povoa (upsert, mesma regra de não duplicar) as 5 abas de config:

- PDS/PAS com origens (PDF/PAF cujo `PNT` aponta pra eles) → linhas em
  `PontoDigital`/`PontoAnalogico` (1 origem = 1 linha; N origens = redundância);
- PDS/PAS **sem nenhum** PDF/PAF correspondente (tipicamente pontos calculados,
  RCA/TCL) **não são extraídos** — não têm componente físico, então ficam fora do
  escopo de `PontoDigital`/`PontoAnalogico` (que descrevem pontos com origem física);
  o ponto em si continua intocado na aba original. Uma versão anterior tentava
  preservá-los com um `ID_Fisico` fictício, mas isso gerava um PDF/PAF fantasma toda
  vez que `unificar_pontos` rodava de novo;
- CGS cujo **`PAC`** aponta pro `ID` de um PDS/PAS → marca `Comando=S` na origem
  correspondente (em `PontoAnalogico`, também traz `ID_Fisico_Comando`/
  `KCONV_Comando` do CGF ligado e os 4 limites `LMI1C/LMI2C/LMS1C/LMS2C` direto
  do CGS). **`CGS.PAC` é o FK de verdade aqui, não `CGS.ID`** — achado real (base
  jdm/CHESF) tem um CGS com `ID` totalmente independente do ponto
  (`ID=JDM:REGU:STPC`, `PAC=JDM:REGU-STPS`) — o forward desta planilha sempre usa
  o mesmo `ID` nos dois por convenção própria (e também seta `PAC` corretamente),
  mas bases de terceiros nem sempre seguem essa convenção;
- CGS cujo `PAC` **não** aponta pra nenhum PDS/PAS extraído → vai pra
  `ComandoAvulso` (com os 4 limites também, quando presentes no CGS);
- PDD/PAD → um canal por `TDD` distinto em `CanaisDistribuicao` (com **Método
  inferido automaticamente** quando é Prefixo/Sufixo simples — compara o ID original
  com o transformado; quando não dá pra inferir com confiança, assume `Explicito`
  e preenche `IDExplicito` com o ID observado, em vez de arriscar uma
  transformação que pode não valer pros outros pontos do mesmo canal) + a ligação
  em `DistribuicaoPontos`.

Rode `extrair_pontos` **antes** de estender uma base já existente pelo modelo
unificado (ex.: adicionar uma 2ª origem redundante a um ponto que já existe). É uma
reconstrução de melhor esforço, não um inverso perfeito — confira o resultado antes
de confiar cegamente, principalmente o Método inferido dos canais.

### Ganhos rápidos

**Troca de ID global — `trocar_id_global`**
Aba `TrocaId` (`IDAntigo | IDNovo | Ativa`, criada vazia na 1ª execução): renomeia um
ID onde ele é a própria chave e propaga a troca por **toda** coluna que o referencia,
usando o mesmo grafo de relacionamentos do verificador (`REGRAS_REFS_PADRAO`) — ex.:
renomear um `PDS` atualiza `PDD.PDS`, `RCA.PARC`, etc. automaticamente. Suporta várias
trocas em lote numa só execução (linhas processadas em ordem — uma troca pode
encadear na seguinte). ID não encontrado ou ambíguo (existe em mais de uma entidade)
não altera nada e fica registrado como tal. Relatório em `RelatorioTrocaId`.

**Estatística — `estatistica_base`**
Conta linhas totais e ativas (`Gera = x`) por entidade, com uma linha `TOTAL` no
fim. Relatório em `Estatística`.

**Gestão de includes — `gerir_includes`**
Aba `SubstituirIncludes` (`Buscar | Substituir | Ativa`, criada vazia na 1ª execução):
aplica as substituições ativas no caminho de **toda** linha de include (`Gera = i` ou
`u`), em qualquer entidade — corrige em lote sem precisar editar célula por célula.
Sempre escreve o estado atual de todos os includes (já com as trocas aplicadas) em
`RelatorioIncludes`, então também serve só pra listar (deixe `SubstituirIncludes`
vazia/inativa).

### Assistente de protocolo/IED — `gerar_ied`
Cria a infraestrutura de canal de um IED (LSC/CNF/CXU/UTR/ENU + `TAC` na aquisição
ou `TDD` na distribuição, mais os NV1/NV2 "grupo" padrão do protocolo) a partir da
aba `IEDs` (criada vazia na 1ª execução). Depois de rodar, você referencia os NV2
criados aqui nas abas `PontoDigital`/`PontoAnalogico`/`ComandoAvulso` (Unificação de
Pontos) para gerar os pontos individuais — esta parte só monta a "casca" onde os
pontos vão morar, não cria PDF/PDS/PAF/PAS/CGF/CGS.

**Protocolos disponíveis**: **104**, **101**, **DNP3**, **MODBUS**, **61850**,
**SNMP** e **ICCP**. 104, 101 e **DNP3** confirmados contra bases reais nos
dois sentidos (aquisição **e** distribuição — a distribuição de DNP3 só foi
confirmada depois, contra uma base real diferente da usada pra aquisição, ver
seção própria abaixo); MODBUS confirmado contra base real só na **aquisição**
— a distribuição foi extrapolada por consistência (mesmo formato "stripped"
do 104/101) e pelo código de referência da macro GE, sem uma base real de
distribuição disponível pra validar (a extrapolação de DNP3 seguia esse mesmo
caminho até uma base real ser encontrada — ver nota); 61850 é bidirecional por
natureza (ver seção própria abaixo) e foi confirmado contra 12 IEDs reais;
SNMP é só aquisição por natureza (protocolo de monitoramento, não modela
comando — ver seção própria) e foi confirmado contra 2 bases reais
independentes; ICCP é bidirecional por natureza igual ao 61850 (ver seção
própria) e foi confirmado contra o **manual oficial** do protocolo (não contra
uma base real — o acervo de referência não tem nenhuma, ver nota na seção do
ICCP). **103** ficou de fora por ora — não há nenhuma base real com IEC 103 no
acervo de referência, só documentação de manual. **OPC UA** e **C37.118**
ficam de fora por ora, sem base real nem documentação suficiente disponível
para validar. Retomar qualquer um dos dois se aparecer uma base real ou um
manual equivalente ao que resolveu o ICCP.

101 e DNP3 tipicamente rodam por serial — configure a entrada correspondente em
`tsr.conf` (`config/<base>/sys/tsr.conf`, transportador `iec1s`/`iec2s`/`iec2t`
para 101, `iec3s` para DNP3 serial) à parte; é um arquivo de sistema fora do
modelo desta planilha, `gerar_ied` não mexe nele. MODBUS também pode rodar por
serial (RTU) ou TCP — mesma ressalva, ajuste `tsr.conf` à parte se for serial.

Colunas da aba `IEDs`: `ID, Protocolo, Direcao (Aquisicao/Distribuicao), Nome, GSD,
INS, MAP, NSRV1, NSRV2, PlPr, LiPr, PlRe, LiRe, IGNERS, SINCR, INVAL, TZBR, DnpLvl,
PROTO, ApTitle, AeQ, PS, SS, TS, IDAD, KEEP, NREP, TOUT, MPDU, OPMSK, GOOSE, VERSAO,
HOST, COMMUNITY, VERBD, NSERV1, NSERV2, IDIG, IANL, IDIS, T2V, BLC3, AQANL, AQPOL,
AQTOT, INTGR, NFAIL, SFAIL, FAILP, FAILR, NTENT, RESPT, TDESC, TRANS, VLUTR,
Redundante, Gera`. `IGNERS/SINCR/INVAL` só valem para 104/101; `TZBR/DnpLvl` só
para DNP3; `PROTO` só para MODBUS; `ApTitle` até `GOOSE` só para 61850 (e
parcialmente reaproveitado por ICCP, ver abaixo); `VERSAO/HOST/COMMUNITY` só
para SNMP; `VERBD` até `BLC3` só para ICCP (cada protocolo usa seu próprio
conjunto de campos extras no `CONFIG` do CNF — colunas que não se aplicam ao
protocolo escolhido ficam simplesmente sem uso; 61850 e ICCP também não usam
`AQANL` até `VLUTR`, ver seções próprias). `INS` (a instalação/estação a que o
`TAC` pertence), `HOST` (o IP do equipamento monitorado via SNMP) e `VERBD` (o
identificador do Acordo Bilateral vigente do ICCP) não têm default: são
específicos do site. A maioria dos outros campos tem um default sensato (ex.:
`MAP=GERAL`, `NSRV1/NSRV2=localhost`, `PROTO=BIN`, `VERSAO=2c`,
`COMMUNITY=public`, `ApTitle=1 1 10 / 1 1 10`) — só preencha o que quiser
mudar; a célula sempre vence o default. `Redundante=S` cria `UTR` em par
(PRI/REV); `ENU` sempre vem em par (redundância de rede), mesmo com um `UTR`
só, igual ao observado na base real (não vale para 61850/ICCP, que não usam
`UTR`/`ENU`; para ICCP, `Redundante=S` cria um 2º servidor `ENM`/`NSERV2` no
lugar).

> ℹ️ Se você já rodou `gerar_ied` (ou qualquer outra função que crie uma aba de
> config) antes de atualizar o SageBonis para uma versão com colunas novas, não
> precisa recriar a aba: rodar a macro de novo adiciona ao final só as colunas
> que estiverem faltando, sem tocar no que já existe.

> ⚠️ **Simplificação assumida**: a base real de aquisição separa digital e
> analógico em NV1 distintos; aqui juntamos tudo num único NV1 de leitura (mais um
> de comando) para manter a aba simples. Reorganize manualmente se precisar
> replicar exatamente um padrão com mais grupos.

**A distribuição de DNP3 tem 3 diferenças reais da aquisição** (confirmado
contra uma base real diferente da usada pra aquisição — não é só o formato
"stripped" do 104/101, extrapolar por analogia teria dado errado):
- `LSC.TTP` sai **`UDPF3`** na distribuição, não o `IEC3S` da aquisição (mesmo
  `TCV=CNVH` nos dois lados).
- `CNF.CONFIG` **também** tem `TZBR`/`DnpLvl` na distribuição — diferente do
  104/101, confirmados sem esses extras do lado distribuição.
- Sai **1 `TDD`** só (mesmo `ID` do `LSC`), sem o split `_DIG`/`_ANA` do
  104/101/MODBUS.
- O grupo de comando da distribuição roteia **`CDUP` e `CSIM` juntos** — a
  aquisição só tem `CDUP` (confirmado, sem `CSIM`).

**MODBUS é o mais diferente dos quatro protocolos "clássicos"**: seu grupo de
leitura não usa ASIM/ADUP/APFL como a família 60870/DNP3 — usa `ALAT` (digital
via registrador latch), `AANL` (analógico via holding register) e `ASTP`
(analógico via input register, FC4). A aquisição real também gera **2 registros
de `TAC`** (não 1): um `TPAQS=ASAC` normal e um `TPAQS=AFIL` à parte, específico
para analógicos com conversão float (2 registradores 16-bit compondo 1
IEEE-754) — `gerar_ied` já cria os dois automaticamente. MODBUS tem mais
variantes reais de registrador (faixas 0x/1x/2x/3x/4x) que não tentamos cobrir
todas; ajuste manualmente se seu MDB usar um range diferente. `grupos_comando`
(CDUP) é mantido por consistência com os outros protocolos, mas **não foi
confirmado** contra base real — a base disponível era só de medição/cálculo,
sem comando.

**61850 tem uma casca inteiramente diferente dos outros quatro** (confirmado
contra 12 IEDs reais de uma base didática de referência, 100% consistentes) —
a associação MMS é **bidirecional por natureza**, então:
- **1 linha na aba `IEDs` já é o IED completo** — `Direcao` não é usado (a
  mesma conexão faz aquisição e distribuição juntas). `LSC.TIPO` sai sempre
  `AD` (nunca `AA`/`DD`), e `TAC` **e** `TDD` são sempre criados os dois, com o
  mesmo `ID` do `LSC`.
- **Sem `CXU`/`UTR`/`ENU`** — a base real não usa essa camada para 61850 (a
  própria associação MMS já é a "conexão"; não há remota serial pra modelar).
  `Redundante` não tem efeito aqui. Usa **`MUL`+`ENM`** no lugar (achado real —
  a planilha do usuário já tinha 90 `MUL`/180 `ENM` de 61850 pré-existentes ao
  implementar o ICCP, confirmando que 61850 também usa essa camada): 1 `MUL`
  por canal, com `ID` igual ao `CNF` (diferente do ICCP, que usa um sufixo
  `_AQ` — 61850 não tem a noção de "domain name" separado por direção), e
  **sempre 2 `ENM`** (confirmado 90/90 na base real — não é condicional a
  `Redundante`, mais parecido com o `ENU` sempre-em-par dos protocolos
  clássicos).
- `TN1` sai fixo `NLN1` (não varia por papel, ao contrário dos outros
  protocolos); os grupos de leitura/comando usam `ADAQ` (digital), `AAAQ`
  (analógico) e `CSIM` — **comando simples**, não `CDUP` (comando duplo) como
  os outros 4 protocolos.
- `CNF.CONFIG` usa campos de associação MMS totalmente diferentes — nada de
  `PlPr/LiPr/PlRe/LiRe`: `ApTitle, AeQ, PS, SS, TS, IDAD, KEEP, NREP, TOUT,
  MPDU, OPMSK, GOOSE`. A maioria tem default sensato confirmado contra a base
  real (`OPMSK=228521`, o valor mais comum do acervo; `GOOSE=0`); `ApTitle`
  (endereçamento MMS do IED) não tem default — é específico do site.
- **Redundância "IED virtual"** (2 IEDs físicos + 1 virtual que assume o
  controle via bit 12 do `OPMSK`) **não é automatizada** — não cabe no modelo
  de 1-linha-por-IED desta aba. Se precisar desse padrão, crie as 3 linhas
  manualmente (2 físicas, sem bit 12; 1 virtual, com `OPMSK` ajustado para
  incluir o bit 12 — ex. `OPMSK=8010`) e ajuste os `NV1`/`NV2` gerados à mão.

**SNMP volta a caber no caminho padrão** (ao contrário do 61850) — tem
`CXU`/`UTR`/`ENU` normalmente e `LSC.TIPO` segue `Direcao` como nos 4
protocolos "clássicos" — mas com 3 diferenças pontuais, confirmadas contra 2
bases reais independentes (100% consistentes entre si):
- `CNF.CONFIG` não tem `PlPr/LiPr/PlRe/LiRe` — usa `VERSAO` (SNMP versão 2c),
  `HOST` (IP do equipamento monitorado, sem default) e `COMMUNITY` (default
  `public`, o community string padrão do próprio SNMP).
- `TN1` sai fixo `SNM1` (sem prefixo `A`/`C`/`D`/`O`), e **não existe grupo de
  comando** — SNMP é só monitoramento (0/13 exemplos reais com algum ponto de
  comando); só o grupo de leitura é criado, com 1 único tipo confirmado (`ASIM`
  digital — status ligado/desligado de equipamento de rede/TI). Se precisar de
  um valor analógico via SNMP (ex.: um MIB do tipo `Gauge32`), adicione esse
  grupo manualmente — não confirmado contra base real.
- `UTR.ENUTR` sai `1` no `PRI` / `0` no `REV`, diferente do `9` fixo usado
  pelos outros 4 protocolos (também confirmado nas 2 bases).

SNMP é só aquisição por escopo — os 2 exemplos reais confirmam isso (só
`TIPO=AA`); é um protocolo de monitoramento, não modela distribuição/comando.

**ICCP/TASE.2 é o único protocolo confirmado só contra o manual oficial**
(`SAGE_ManCfg_Anx15_ICCP_rev21.pdf`, CEPEL), não contra uma base real — o
acervo de referência (SkillSAGE) não tem nenhuma disponível: a "referência"
que existia lá (biblioteca de protocolos) na verdade contém dados no padrão
104, e a única base real marcada como "ICCP" no inventário não tem nenhum
traço do protocolo em nenhuma entidade. O manual, porém, é completo e
determinístico o suficiente (é a fonte primária CEPEL/SAGE) pra implementar
com confiança. Pontos de atenção:
- Existem **dois mecanismos** de ICCP no SAGE, bem diferentes: o **conversor
  "iccp"** (fino, escolhe exatamente quais pontos vão pra qual centro remoto —
  é o que `gerar_ied` modela) e o **servidor "SICCP"** (genérico, expõe
  automaticamente TODOS os pontos PDS/PAS/PTS/CGS pra qualquer VCC autorizado,
  **sem nenhuma configuração de entidades** — é só um arquivo de sistema,
  `siccp.cnf`, mesma categoria do `tsr.conf`, fora do modelo desta planilha).
  Use SICCP quando não precisar escolher pontos por centro remoto (a maioria
  dos casos reais, a julgar pela ausência de bases com o conversor fino no
  acervo); use "iccp" (via `gerar_ied`) quando precisar dessa granularidade.
- Como o 61850, é bidirecional por natureza (`LSC.TIPO="AD"` sempre,
  `Direcao` não é usado) e usa `TN1` fixo `NLN1`.
- **Sem `CXU`/`UTR`/`ENU`/`TAC`/`TDD`** — usa duas entidades novas: `MUL`
  ("Multiligação com Centro de Controle Remoto", domain name da direção de
  aquisição) e `ENM` ("Enlace de multiligação", servidor principal/reserva do
  centro remoto — `Redundante=S` cria o 2º).
- **Um único NV1 reúne até 8 tipos de NV2** simultaneamente — aquisição
  (`ADAQ`/`AAAQ`/`ATTA`/`CSIM`) **e** distribuição (`DDAQ`/`DAAQ`/`DTTA`/`CDUP`)
  podem coexistir no mesmo canal MMS, já que o mesmo VCC pode aquisitar e
  distribuir ao mesmo tempo — diferente de todos os outros protocolos, que
  fazem só um dos dois por linha da aba `IEDs`.
- `CNF.CONFIG` reaproveita `ApTitle/AeQ/PS/SS/TS` do 61850 (mesmo formato e
  default), mas troca os campos obrigatórios: `IDIG/IANL/IDIS/TOUT/MPDU/T2V/
  OPMSK/BLC3` (confirmado no manual, incluindo os nomes das 8 variáveis).
  `OPMSK` usa default **0** aqui — diferente do 228521 do 61850, mesma coluna
  da planilha. `IDIG/IANL/IDIS` (temporizadores de integridade) não têm
  default — o valor certo depende de o VCC remoto suportar o bloco 2 do ICCP.

#### Extração reversa de IEDs (bases já existentes) — `extrair_pontos`
`gerar_ied` sozinho só ajuda com **canais novos**: um IED já importado de uma
base real (`LSC`/`CNF`/`CXU`/`UTR`/`ENU`/`TAC`/`MUL`/`ENM` vindos de `.dat`) não
aparece na aba `IEDs` automaticamente. `extrair_pontos` (o mesmo comando que já
reconstrói `PontoDigital`/`PontoAnalogico`/`ComandoAvulso`, ver seção anterior)
também reconstrói `IEDs` a partir de cada `LSC` já existente:

- **1 `LSC` = 1 linha de `IEDs`** — aquisição e distribuição de um protocolo
  "clássico" são `LSC` distintos (nunca fundidos numa linha só), exatamente
  como a própria geração também trata cada um como independente.
- **`Protocolo`** vem de `(LSC.TCV, LSC.TTP)` comparado contra cada entrada
  conhecida (incluindo o `ttp_distribuicao` do DNP3 — um `LSC.TTP=UDPF3` já
  identifica "DNP3 distribuição" sem precisar olhar mais nada). Um `LSC` cujo
  `TCV`/`TTP` não bate com nenhum protocolo conhecido (ainda não modelado, ex.
  103/OPC UA/C37.118) é **ignorado silenciosamente** — mesmo espírito do
  Método não reconhecido em `CanaisDistribuicao`.
- **`Direcao`** vem de `LSC.TIPO` (`AA`→Aquisicao/`DD`→Distribuicao) pros
  protocolos "clássicos"; fica vazia para 61850/ICCP (bidirecional, `Direcao`
  não se aplica, mesma regra do forward).
- **`CNF.CONFIG`** é parseado de volta campo a campo, sabendo lidar com valor
  multi-token (ex. `ApTitle= 1 1 10 / 1 1 10` do 61850/ICCP) porque procura o
  próximo campo **esperado** pro protocolo/direção em questão, não o próximo
  espaço em branco.
- **`INS`** vem do `TAC` ligado ao `LSC` (quando existe — só na aquisição dos
  protocolos "clássicos", ou sempre em 61850); **`AQANL`/`AQPOL`/`AQTOT`/
  `INTGR`/`NFAIL`/`SFAIL`/`FAILP`/`FAILR`** vêm do `CXU`; **`NTENT`/`RESPT`**
  do `UTR` (`PRI`); **`TDESC`/`TRANS`/`VLUTR`** do `ENU` (`PRI`).
  `Redundante` é **inferido**, nunca lido de um campo próprio (não existe um
  campo assim em nenhuma entidade): 2 `UTR` (`PRI`+`REV`) para os protocolos
  "clássicos"/SNMP, 2 `ENM` do mesmo `MUL` para ICCP, e **sem efeito nenhum**
  para 61850 (mesma regra do forward — fica de fora).

Rode `extrair_pontos` **antes** de estender uma base já existente com um novo
canal do mesmo protocolo, pelas mesmas razões da extração de pontos: é uma
reconstrução de melhor esforço, não um inverso perfeito.

> 📎 **Exemplo pra conferir visualmente**: [`exemplo_validacao/`](exemplo_validacao/)
> tem uma planilha pequena, já processada, com um fragmento real (DNP3+61850+
> MODBUS+SNMP, mais um setpoint analógico de verdade) — ver o README lá dentro.

## Instalação e uso
Igual à Simples: abra `SageBonis.ods` e habilite as macros do documento (a macro vem
embutida). Atribua as funções `verificar_base`, `unificar_pontos`, `extrair_pontos`,
`trocar_id_global`, `estatistica_base`, `gerir_includes` e `gerar_ied` a botões ou
atalhos, como as demais.

## Sincronizar a macro com o .ods
A partir da raiz do repositório:

```bash
python sync_macro.py inject  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
python sync_macro.py status  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
```

## Testes (`completa/tests/`)
```bash
python completa/tests/run_all.py            # roda tudo
python completa/tests/run_all.py --sem-uno   # só os smoke tests (sem soffice)
```
- **Smoke tests em memória** (`smoke_test_*.py`) — lógica pura, sem LibreOffice, rodam em
  segundos: `smoke_test_ied.py` (assistente de protocolo + extração reversa de IEDs, ~160
  checks, incluindo round-trip forward→reverso pelos 7 protocolos), `smoke_test_unificacao.py`,
  `smoke_test_extracao.py`, `smoke_test_ganhos_rapidos.py`. Importam `ImportadorSAGE.py` direto
  (`importlib`) e chamam as funções `_gerar_*`/`_extrair_*`/etc. isoladas de qualquer UNO.
- **Teste UNO real** (`teste_uno_protocolos.py`) — sobe um `soffice --headless`, copia o
  `SageBonis.ods` real pra um arquivo descartável (`/tmp`), injeta a macro atual, roda os 7
  protocolos via macro de verdade e confirma que o upsert não mexeu nos dados reais
  pré-existentes (ex.: `mul`/`enm` de 61850). **Nunca escreve no `.ods` rastreado.** Requer
  `soffice` no `PATH`. O ciclo de vida (subir/derrubar processo, profile, cópia temporária) é
  todo administrado por `uno_harness.py` (`class TesteUno`, use como *context manager*, aceita
  `ods_origem`/`py_origem` pra apontar pra outra planilha/macro).
- **Teste de paridade import/export** (`teste_paridade_import_export.py`) — critério de
  maturidade da convergência (`PLANEJAMENTO.md`): gera uma base `.dat` sintética pequena
  (autocontida, sem depender de nenhum caminho externo), importa numa cópia do `SageBonis.ods`
  em branco da raiz rodando o `ImportadorSAGE.py` da Simples, e noutra cópia do MESMO arquivo em
  branco rodando o `ImportadorSAGE.py` da Completa — depois exporta as duas de volta e faz diff
  byte-a-byte. Falha se qualquer `.dat` divergir entre as trilhas.
- Ao adicionar um protocolo/recurso novo, estenda o smoke test correspondente e, se a mudança
  envolver criação de aba/entidade nova, adicione um caso em `teste_uno_protocolos.py`. Se
  mexer no núcleo de import/export compartilhado com a Simples, rode
  `teste_paridade_import_export.py` antes de commitar.

## Status
🚧 Em desenvolvimento, mas os 3 itens do roadmap original (verificador, unificação
de pontos, assistente de protocolo/IED) já foram entregues. Do assistente de
protocolo/IED, restam só protocolos sem base real (nem manual completo, no
caso de OPC UA/C37.118) disponível pra validar (103, OPC UA, C37.118 — ver
[PLANEJAMENTO.md](../PLANEJAMENTO.md)), retomados se aparecer uma base ou
documentação equivalente ao que resolveu o ICCP.
