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

**Protocolos disponíveis**: **104** e **101**, confirmados contra bases reais
(aquisição **e** distribuição). Próximos da lista: 103, DNP3, MODBUS, 61850, SNMP,
ICCP/SICCP (OPC UA e C37.118 ficam de fora por ora, sem base real disponível para
validar). O 101 tipicamente roda por serial — configure a entrada correspondente em
`tsr.conf` (`config/<base>/sys/tsr.conf`, transportador `iec1s`/`iec2s`/`iec2t`) à
parte; é um arquivo de sistema fora do modelo desta planilha, `gerar_ied` não mexe
nele.

Colunas da aba `IEDs`: `ID, Protocolo, Direcao (Aquisicao/Distribuicao), Nome, GSD,
MAP, NSRV1, NSRV2, PlPr, LiPr, PlRe, LiRe, IGNERS, SINCR, INVAL, AQANL, AQPOL, AQTOT,
INTGR, NFAIL, SFAIL, FAILP, FAILR, NTENT, RESPT, TDESC, TRANS, VLUTR, Redundante,
Gera`. A maioria tem um default sensato (ex.: `MAP=GERAL`, `NSRV1/NSRV2=localhost`) —
só preencha o que quiser mudar; a célula sempre vence o default. `Redundante=S` cria
`UTR` em par (PRI/REV); `ENU` sempre vem em par (redundância de rede), mesmo com um
`UTR` só, igual ao observado na base real.

> ⚠️ **Simplificação assumida**: a base real de aquisição separa digital e
> analógico em NV1 distintos; aqui juntamos tudo num único NV1 de leitura (mais um
> de comando) para manter a aba simples. Reorganize manualmente se precisar
> replicar exatamente um padrão com mais grupos.

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
