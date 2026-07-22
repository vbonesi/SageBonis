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
  - **Comando** (só `PontoDigital`): coluna `Comando=S` gera CGF/CGS com o **mesmo ID**
    do PDS (regra fixa do SAGE: comando e status compartilham o ID). Se 2+ origens
    tiverem `Comando=S`, gera um CGF por origem, todos referenciando o mesmo CGS.
- **`ComandoAvulso`** — comandos **sem** ponto de status próprio (ex.: um `COM_SAGE`
  genérico ligado a um TAC local, como algumas bases já usam). Cada linha tem seu
  próprio `ID` de CGS/CGF; várias linhas podem repetir o mesmo `TAC`/`PAC` (o ponto
  genérico) — é justamente o caso de vários comandos ligados ao mesmo ponto.
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
- CGS com o mesmo ID de um PDS/PAS → marca `Comando=S` na origem correspondente;
- CGS **sem** PDS/PAS correspondente → vai pra `ComandoAvulso`;
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

**Protocolos disponíveis**: **104**, **101**, **DNP3**, **MODBUS** e **61850**.
104 e 101 confirmados contra bases reais nos dois sentidos (aquisição **e**
distribuição); DNP3 e MODBUS confirmados contra base real só na **aquisição** —
a distribuição de ambos foi extrapolada por consistência (mesmo formato
"stripped" do 104/101) e pelo código de referência da macro GE, sem uma base
real de distribuição disponível pra validar; 61850 é bidirecional por natureza
(ver seção própria abaixo) e foi confirmado contra 12 IEDs reais. **103** ficou
de fora por ora — não há nenhuma base real com IEC 103 no acervo de referência,
só documentação de manual. Próximos da lista: SNMP, ICCP/SICCP (este com
aquisição **e** distribuição — ICCP funciona nos dois sentidos entre centros de
controle). OPC UA e C37.118 ficam de fora por ora, sem base real disponível
para validar.

101 e DNP3 tipicamente rodam por serial — configure a entrada correspondente em
`tsr.conf` (`config/<base>/sys/tsr.conf`, transportador `iec1s`/`iec2s`/`iec2t`
para 101, `iec3s` para DNP3 serial) à parte; é um arquivo de sistema fora do
modelo desta planilha, `gerar_ied` não mexe nele. MODBUS também pode rodar por
serial (RTU) ou TCP — mesma ressalva, ajuste `tsr.conf` à parte se for serial.

Colunas da aba `IEDs`: `ID, Protocolo, Direcao (Aquisicao/Distribuicao), Nome, GSD,
INS, MAP, NSRV1, NSRV2, PlPr, LiPr, PlRe, LiRe, IGNERS, SINCR, INVAL, TZBR, DnpLvl,
PROTO, ApTitle, AeQ, PS, SS, TS, IDAD, KEEP, NREP, TOUT, MPDU, OPMSK, GOOSE, AQANL,
AQPOL, AQTOT, INTGR, NFAIL, SFAIL, FAILP, FAILR, NTENT, RESPT, TDESC, TRANS, VLUTR,
Redundante, Gera`. `IGNERS/SINCR/INVAL` só valem para 104/101; `TZBR/DnpLvl` só
para DNP3; `PROTO` só para MODBUS; `ApTitle` até `GOOSE` só para 61850 (cada
protocolo usa seu próprio conjunto de campos extras no `CONFIG` do CNF —
colunas que não se aplicam ao protocolo escolhido ficam simplesmente sem uso;
61850 também não usa `AQANL` até `VLUTR`, ver seção própria). `INS` (a
instalação/estação a que o `TAC` pertence — entidade própria, referenciada por
`TAC.INS`) não tem default: é específico do site, preencha o código já usado no
restante da sua base. A maioria dos outros campos tem um default sensato (ex.:
`MAP=GERAL`, `NSRV1/NSRV2=localhost`, `PROTO=BIN`) — só preencha o que quiser
mudar; a célula sempre vence o default. `Redundante=S` cria `UTR` em par
(PRI/REV); `ENU` sempre vem em par (redundância de rede), mesmo com um `UTR` só,
igual ao observado na base real (não vale para 61850, que não usa `UTR`/`ENU`).

> ⚠️ **Simplificação assumida**: a base real de aquisição separa digital e
> analógico em NV1 distintos; aqui juntamos tudo num único NV1 de leitura (mais um
> de comando) para manter a aba simples. Reorganize manualmente se precisar
> replicar exatamente um padrão com mais grupos.

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
  `Redundante` não tem efeito aqui.
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

## Status
🚧 Em desenvolvimento. Único item do roadmap ainda pendente (ver
[PLANEJAMENTO.md](../PLANEJAMENTO.md)): assistente de protocolo/IED.
