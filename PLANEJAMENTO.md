# Planejamento — SageBonis

Documento vivo de estratégia e roadmap. Não descreve o que já existe (isso está no
[README](README.md) e no [CLAUDE.md](CLAUDE.md)), e sim **para onde o projeto vai**.

## Filosofia: duas trilhas no mesmo repositório

O SageBonis vai manter **duas variantes** convivendo no mesmo repo, porque atendem
perfis de uso diferentes:

- **Trilha Simples (atual)** — a planilha e a macro como estão hoje. Foco em
  **importar/exportar base rápido**, sem recursos que atrapalhem alterações
  pontuais. É a preferida de quem (inclusive o autor) valoriza agilidade e baixo
  atrito. **Princípio inegociável: ela permanece simples e estável.** Nada de
  empilhar funcionalidade aqui.
- **Trilha Completa (nova, forkada)** — um fork da macro + uma nova planilha, com
  os recursos avançados (verificação de base, unificação de pontos, assistentes de
  protocolo). Melhor para **novos desenvolvedores / configuração de uma SE do zero**,
  onde estrutura e validações compensam o atrito extra.

### Meta de longo prazo: convergência
O objetivo final é chegar a **uma única planilha que trabalhe das duas formas**
(um "modo simples" e um "modo completo" na mesma base) e então **unificar**. Mas a
unificação só acontece **quando a trilha Completa atingir um bom nível de
maturidade** (ver [Critérios de maturidade](#critérios-de-maturidade-para-convergir)).
Até lá, as duas evoluem em paralelo e de forma independente.

## Estrutura proposta no repositório

Como a trilha Simples não pode ser perturbada, a proposta de menor impacto é manter
a Simples na raiz (como está) e criar uma pasta para a Completa:

```
/                         raiz = Trilha Simples (intocada)
  ImportadorSAGE.py
  SageBonis.ods
  README.md
  sync_macro.py           utilitário compartilhado (já aceita --ods/--py)
  PLANEJAMENTO.md         este documento
  completa/               Trilha Completa (em desenvolvimento)
    ImportadorSAGE.py     fork do script da raiz
    SageBonis.ods         nova planilha, mais estruturada
    README.md             específico da variante
```

Nomes a definir (sugestões): variante = "SageBonis Pro" / "Completa" / "Plus".
Decisão de layout (raiz vs. subpasta `simples/`) fica em aberto — a recomendação é
**não mover a Simples** por enquanto, para zero risco.

> O `sync_macro.py` já serve as duas: `python sync_macro.py inject --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py`.

## Funcionalidades a portar para a Trilha Completa

Ideias extraídas de duas macros VBA de planilhas que servem ao mesmo fim (ver
[Referências](#referências)). **Não são para copiar** — são para **re-implementar
as ideias** em Python, dirigidas por configuração (ver [Princípios](#princípios-de-design)).

### 1. Verificador / linter de base  ·  origem: Eletronorte‑2  ·  prioridade 🥇
Aba "Análise" que roda checagens de **integridade referencial cruzada** e reporta
em log com severidade (Erro/OK/Info), entidade, linha, atributo, valor, descrição.

**Entregue** (estado atual e detalhes em [completa/README.md](completa/README.md)): ID
duplicado/vazio; referências cruzadas (agora com **81 regras pré-populadas**, curadas via
SkillSAGE, e suporte a FK "ambígua" — múltiplas entidades de destino por regra); comprimento
de ID acima do limite conhecido da entidade; valor fora do domínio conhecido (reaproveita a
aba `EntidadeAtributoValor`, que já existia mas estava sem consumidor).

**Ainda pendente** desta frente: prefixo do ID = sigla da SE (precisa de config indicando
qual é a sigla "correta" para a base carregada — não é universal como as demais checagens).

**Por que primeiro:** maior ROI, menor risco (read‑only, não toca no formato de
exportação), e o Python leva vantagem real sobre o VBA (índices `dict`/`set` com
busca O(1) no lugar dos loops O(n²) do original). Entregável incremental: uma
checagem por vez.

### 2. Unificação / distribuição de pontos  ·  origem: GE  ·  prioridade 🥈
Tabela‑mestre **Digital / Analógico / Comando** onde cada ponto físico é definido
**uma vez**, com marcação de a quais "distribuições" (canais/protocolos) pertence.
Um gerador faz o **fan‑out** para as entidades relacionadas (Digital → PDS/PDF/PDD;
Analógico → PAS/PAF/PAD; Comando → CGS/CGF), aplicando naming por protocolo
(101/104/DNP3) e método de ID configurável (Prefixo/Sufixo/Substituir).

**Entregue** (estado atual e detalhes em [completa/README.md](completa/README.md)):
abas `PontoDigital`/`PontoAnalogico` (uma linha por origem física; múltiplas linhas
com o mesmo ID lógico = redundância genérica, resolvida por `RFC` em cadeia, sem
assumir a topologia fixa P/D/virtual do pySAGE); `Comando=S` gera CGF/CGS com o mesmo
ID do PDS; `ComandoAvulso` cobre comandos sem status próprio (ponto genérico tipo
`COM_SAGE`, vários comandos no mesmo TAC/PAC); `CanaisDistribuicao` +
`DistribuicaoPontos` substituem os "4 slots fixos" da macro GE por N canais
configuráveis, com Método (Prefixo/Sufixo/Substituir) por canal. Escrita por upsert
(casa por ID nas abas de entidade já existentes; regenerar não duplica). Também
entregue: `extrair_pontos`, o espelho reverso — reconstrói as 5 abas de config a
partir de uma base **já importada** (com inferência automática de Método por
Prefixo/Sufixo), fechando o ciclo pra bases reais existentes, não só pontos novos.

**Ainda pendente** desta frente: endereçamento automático do `ID_Fisico`/`NV2` por
protocolo (101/104/DNP3/61850) — hoje o usuário informa esses campos prontos; migrar
para o **Assistente de Protocolo/IED** (item 3), que já nasce dependente disso.

**Entregue — comando para pontos analógicos (setpoints)**: `PontoAnalogico` ganhou
`Comando`/`ID_Fisico_Comando`/`KCONV_Comando` (mesma convenção do `PontoDigital`) mais
`LMI1C/LMI2C/LMS1C/LMS2C` (limites inferior/superior do comando, direto no `CGS`).
`ComandoAvulso` também ganhou os 4 campos de limite, cobrindo o caso avulso analógico
(achado real `PAC=MC_DUMMY_SAGE_ANA` em `ur_mir`). Achado real que motivou
(`CGS.TIPO=PAS`) confirmado em **6 bases reais independentes**, com dois padrões
distintos — comando de TAP (2 estados, tipo `AUMD`=Aumentar/Diminuir do `TCTL`, em
`conv_iccp104/GRD`, `padrao_copel`, `siemens_ds_din`, `jdm`) e setpoint numérico com
limites (`tucurui`, provavelmente CAG/despacho de geração; `jdm`
`PAC=JDM:REGU-STPS`). **Nuance não resolvida**: modelado como um único design
genérico (mesmas 4 colunas de limite pras duas variantes) — a semântica exata de
`LMI1C` vs `LMI2C` (normal vs emergência? ou outra coisa?) não foi confirmada a
fundo antes de implementar; se aparecer um caso real que não encaixe no genérico,
pode precisar de um passe de revisão.

**Por que importa:** é exatamente o item de roadmap "unificar abas de entidades em
grupos compactos". A GE é o **blueprint pronto** dele. Esforço médio‑alto.

### 3. Assistente de protocolo / IED  ·  origem: GE  ·  prioridade 🥉
A partir de **IED + protocolo + direção (aquisição/distribuição)**, gera o esqueleto
padrão de infraestrutura de canal (LSC/CNF/CXU/UTR/ENU/TAC-ou-TDD/NV1/NV2), que
depois alimenta a Unificação de Pontos (item 2) via os NV2 criados.

**Escopo final acordado** (ordem de implementação): **104 → 101 → 103 → DNP3 →
MODBUS → 61850 → SNMP → ICCP/SICCP**, com aquisição **e** distribuição para
104/101/DNP3 **e ICCP/SICCP** (ICCP funciona nos dois sentidos entre centros de
controle; os demais só aquisição, seguindo o que a própria macro GE de
referência já limitava). Fora do escopo por ora: **103** (sem nenhuma base real
no acervo, só documentação de manual — diferente de OPC UA/C37.118, que nem
documentação tem) e **OPC UA**/**C37.118** — sem base real nem manual
suficientemente completo disponível; retomar qualquer um se aparecer uma base
ou documentação equivalente ao que resolveu o ICCP (ver abaixo).

**Entregue**: **104**, **101**, **DNP3** e **MODBUS**, confirmados contra bases
reais (`conv_iccp104`/GRD para 104, aquisição e distribuição; base do próprio
usuário — SE Miracema/`neoenergia` — para 101, aquisição e distribuição;
`ctl_dnp_mdb`/DJ9E539 para DNP3 aquisição e `mdb_alat_calc`/MDB1 para MODBUS
aquisição, na entrega inicial — DNP3 teve a distribuição confirmada depois
contra outra base real, ver nota mais abaixo; MODBUS segue com a distribuição
extrapolada por consistência, sem base real disponível). `PARAMS_PROTOCOLO`
generalizado para cobrir as diferenças reais
entre protocolos: DNP3 usa sufixo "DNP" no TN1 (não "DNP3"), TN2 analógico
"AANL" (não "APFL"), e campos `TZBR`/`DnpLvl` no `CNF.CONFIG` em vez de
`IGNERS`/`SINCR`/`INVAL` (nessa ordem: depois de PlPr/LiPr/PlRe/LiRe, ao
contrário do 104/101). MODBUS foi além disso: seu grupo de leitura inteiro é
outro (`ALAT`/`AANL`/`ASTP`, sem `ASIM`/`ADUP`) — daí a generalização de
`tn2_analogico` (1 campo) para `grupos_leitura`/`grupos_comando` (listas
completas de TN2/TPPNT/descrição); e sua aquisição real gera **2** registros de
`TAC` (TPAQS `ASAC` + `AFIL`, este para analógicos com conversão float), não 1 —
generalizado via `params["tacs"]`. Também foi adicionado o campo `INS`
(instalação/estação do `TAC`, campo `TAC.INS` já existia como FK no verificador
mas não era gerado) a **todos** os protocolos, não só MODBUS — confirmado
presente em 3 das 4 bases reais (101/DNP3/MODBUS), fidelidade que faltava desde
o 104. `gerar_ied` cria a infraestrutura; validado via UNO real para os 4
protocolos, incluindo integração ponta-a-ponta com `unificar_pontos` (NV2 criado
aqui → consumido por um ponto novo em `PontoDigital`) e regressão dos 3
protocolos anteriores após a generalização. 101 e DNP3 tipicamente precisam de
uma entrada em `tsr.conf` (serial), configurada à parte — fora do modelo desta
planilha; MODBUS também pode rodar serial ou TCP, mesma ressalva se for serial.
Detalhes em [completa/README.md](completa/README.md).

Também entregue: **61850**, confirmado contra 12 IEDs reais de uma base
didática de referência (100% consistentes). É o mais diferente de todos —
diferente o bastante que ganhou um caminho próprio (`_gerar_infra_ied_61850`),
não o genérico dos outros 4: a associação MMS é bidirecional por natureza, então
1 linha na aba `IEDs` já é o IED completo (`Direcao` não é usado), `LSC.TIPO`
sai sempre `AD` (nunca `AA`/`DD`), e `TAC`+`TDD` saem sempre os dois juntos, com
o mesmo `ID`. Não usa `CXU`/`UTR`/`ENU` (a base real não usa essa camada — a
própria associação MMS já é a "conexão"), então `Redundante` não tem efeito.
`TN1` sai fixo `NLN1` (não varia por papel); os grupos usam `ADAQ`/`AAAQ`/`CSIM`
— **comando simples**, não `CDUP` (comando duplo) como os outros 4. `CNF.CONFIG`
usa campos de associação MMS totalmente diferentes (`ApTitle/AeQ/PS/SS/TS/IDAD/
KEEP/NREP/TOUT/MPDU/OPMSK/GOOSE`, nada de `PlPr/LiPr/PlRe/LiRe`), com defaults
confirmados contra a base real (`OPMSK=228521`, o mais comum do acervo;
`GOOSE=0`). A redundância real de 61850 (**IED virtual**: 2 físicos + 1 virtual
que assume controle via bit 12 do `OPMSK`) não foi automatizada — não cabe no
modelo de 1-linha-por-IED desta aba; documentado como criação manual das 3
linhas. Validado via UNO real, incluindo a mesma integração ponta-a-ponta com
`unificar_pontos`.

Também entregue: **SNMP**, confirmado contra **2 bases reais independentes**
(100% consistentes entre si). Ao contrário do 61850, coube no caminho padrão de
`_gerar_infra_ied` (tem `CXU`/`UTR`/`ENU`, `LSC.TIPO` segue `Direcao`) — só
precisou de 3 pontos de extensão novos em vez de uma função à parte:
`cnf_campos` (substitui PlPr/LiPr/PlRe/LiRe inteiramente por `VERSAO`/`HOST`/
`COMMUNITY` — SNMP não é um protocolo de enlace mestre/escravo), `tn1_fixo`
(TN1 sempre `SNM1`, sem prefixo A/C/D/O) e `enutr_por_ordem` (`UTR.ENUTR` sai
`1`/`0` no PRI/REV, não o `9` fixo dos outros 4). SNMP também não tem grupo de
comando (`grupos_comando=()` — confirmado 0/13 exemplos reais; é protocolo só
de monitoramento) — o laço de geração de NV1/NV2 foi ajustado pra pular grupos
vazios inteiramente, em vez de criar um NV1 sem NV2 dentro.

**Bug real encontrado e corrigido durante o teste UNO** (não pelos testes de
lógica pura): `_garantir_aba_config` só criava uma aba de config se ela **não
existisse ainda** — mas a aba `IEDs`, já criada (vazia) na planilha real antes
de SNMP existir, não ganhava as 3 colunas novas ao rodar `gerar_ied` de novo,
porque a função simplesmente retornava sem checar se o cabeçalho já existente
estava desatualizado. Corrigido para comparar o cabeçalho atual com o canônico
e completar as colunas que faltarem (sem tocar nas já existentes nem nos
dados) — vale pra qualquer `CABECALHOS_*` que cresça no futuro, não só pro
SNMP. Reaplicado na planilha real depois do fix (mesmo processo de
ensaio-antes-de-aplicar do item anterior).

**ICCP investigado no SkillSAGE primeiro — sem base real** (2 pistas seguidas
até o fim): `30_base_mestre/biblioteca/protocolos/iccp/` (referência
supostamente curada) tem a pasta `distribuicao/` **vazia**, e a `aquisicao/`
na verdade contém dados no **padrão 104**, não nenhum campo de ICCP; `fin_ems`
(única base real do acervo listada "DNP+ICCP") não tem nenhum traço de ICCP em
nenhum `.dat`. A própria documentação curada do SkillSAGE confirma: *"No
acervo NÃO há base ICCP. O `conv_iccp104` é uma ferramenta que converte
ICCP → 104 (...) não ICCP."*

**Mas o usuário apontou o manual oficial** (`SAGE_ManCfg_Anx15_ICCP_rev21.pdf`,
CEPEL, em `Drive/SAGE/Manuais/`) como fonte suficiente — diferente de
103/OPC UA/C37.118 (que não têm nem base real nem manual completo o bastante),
o Anx15 é completo e determinístico (é literalmente de onde uma base real
viria). **Entregue com base nele**: existem **dois mecanismos** de ICCP no
SAGE, confirmados no manual e cruzados com um achado real (ver abaixo) — o
**conversor "iccp"** (fino, LSC/CNF/MUL/ENM/NV1/NV2 por centro remoto, é o que
`gerar_ied` modela) e o **servidor "SICCP"** (genérico, expõe TUDO
automaticamente via um único arquivo de sistema `siccp.cnf`, **sem nenhuma
entidade** — por isso nenhuma base real do acervo tinha entidades de ICCP: a
maioria dos deployments reais usa o SICCP genérico, não o conversor fino).

`gerar_ied` para "iccp": igual ao 61850, é bidirecional por natureza
(`LSC.TIPO="AD"`, `Direcao` não usado, `TN1` fixo `NLN1`), mas **sem**
`CXU`/`UTR`/`ENU`/`TAC`/`TDD` — usa `MUL` ("Multiligação com Centro de
Controle Remoto") + `ENM` ("Enlace de multiligação", servidor principal/
reserva) no lugar. Diferente do 61850, **um único NV1 reúne até 8 tipos de
NV2** — aquisição (`ADAQ`/`AAAQ`/`ATTA`/`CSIM`) **e** distribuição
(`DDAQ`/`DAAQ`/`DTTA`/`CDUP`) podem coexistir no mesmo canal MMS bidirecional.
`CNF.CONFIG` reaproveita `ApTitle`/`AeQ`/`PS`/`SS`/`TS` do 61850 (mesmo
default), mas troca os campos obrigatórios: `IDIG`/`IANL`/`IDIS`/`TOUT`/
`MPDU`/`T2V`/`OPMSK`/`BLC3` — `OPMSK` usa default **0** aqui (não 228521 do
61850, mesma coluna da planilha, resolvido por protocolo dentro da função
dedicada em vez de `_DEFAULTS_IED`, que é compartilhado por nome de campo).

**Achado real durante o teste UNO**: a planilha real do usuário **já tinha**
as abas `mul`/`enm` populadas — não de ICCP, mas do **61850** (o `CNF.CONFIG`
correspondente bate exatamente com o formato ApTitle/.../OPMSK/GOOSE já
confirmado pro 61850). Isso confirma o schema de `MUL`/`ENM` contra uma base
real (mesmas colunas ID/CNF/GSD/ORDEM e ID/MUL/ORDEM do manual) — mas também
revela que **o `_gerar_infra_ied_61850` já entregue NÃO gera `MUL`/`ENM`**,
uma lacuna real no protocolo já implementado (a base real de referência usada
pro 61850, `par`/CTEEP, não tinha CXU/UTR/ENU nem eu cheguei a procurar
MUL/ENM nela na hora). Não corrigido nesta rodada — ficou fora do pedido desta
vez (só ICCP/DNP3), fica como pendência pra próxima sessão se o usuário quiser.

Validado: smoke test em memória (105 checks, incluindo regressão dos 6
protocolos anteriores) + teste ponta-a-ponta via UNO real, incluindo a
integração com `unificar_pontos` e confirmação de que o upsert em `mul`/`enm`
só ADICIONOU linhas (as ~90/~180 linhas reais de 61850 ficaram intactas).

**DNP3-distribuição corrigida com base real** (usuário apontou
`Drive/Projetos/_scada/DNP3-MDB.zip`, base "ctl"/`COGTXA21` — a mesma base do
`ctl_dnp_mdb`/`DJ9E539` já usado pra aquisição, só que com o backup completo
incluindo o lado distribuição). A entrega original do DNP3 tinha a
distribuição **extrapolada por consistência** com 104/101 (mesmo formato
"stripped") — a base real revelou que essa suposição estava **errada** em 3
pontos:
- `LSC.TTP` é **`UDPF3`** na distribuição, não o `IEC3S` da aquisição (mesmo
  `TCV=CNVH`) — generalizado via `params["ttp_distribuicao"]`.
- `CNF.CONFIG` **também** leva `TZBR`/`DnpLvl` na distribuição (`"PlPr= 2
  LiPr= 5 PlRe= 2 LiRe= 6 TZBR= 0 DnpLvl= 3"`) — diferente do 104/101,
  confirmados sem esses extras do lado distribuição — generalizado via
  `params["cnf_extra_tambem_distribuicao"]`.
- Sai **1 `TDD`** só (mesmo `ID` do `LSC`), sem o split `_DIG`/`_ANA` do
  104/101/MODBUS — generalizado via `params["tdd_unico"]`.
- O grupo de comando da distribuição roteia **`CDUP` e `CSIM` juntos**
  (confirmado no `nv2.dat` real) — a aquisição (`DJ9E539`, revalidada nesta
  mesma base) só tem `CDUP`, sem `CSIM` — generalizado via
  `params["grupos_comando_distribuicao"]`.

Todos os 4 pontos de extensão são opcionais (default preserva o comportamento
dos outros protocolos) — só o DNP3 os usa por ora. Validado: smoke test
atualizado (108 checks) + teste ponta-a-ponta via UNO real específico pra essa
correção.

**61850 corrigido para também gerar `MUL`/`ENM`** (fechando a lacuna encontrada
na entrega do ICCP). Confirmado contra os mesmos 90 `MUL`/180 `ENM` reais de
61850 já presentes na planilha do usuário: 1 `MUL` por canal com
`ID`=`CNF`=`ID` do canal (diferente do ICCP, que usa sufixo `_AQ` — 61850 não
tem "domain name" separado por direção), e **sempre 2 `ENM`** (confirmado
90/90, não condicional a `Redundante`, ao contrário do ICCP — mais parecido
com o `ENU` sempre-em-par dos protocolos clássicos). Validado via
`teste_uno_protocolos.py`, incluindo confirmação de que o upsert só somou (+2
`MUL`, +3 `ENM` = 1 do ICCP + 2 do 61850) sem alterar nenhuma das linhas reais
pré-existentes.

**Também pendente — extração reversa de IEDs**: `gerar_ied()` só funciona num
sentido (config `IEDs` → `LSC`/`CNF`/`NV1`/`NV2`/etc.). Diferente de
`extrair_pontos()` (que já reconstrói `PontoDigital`/`PontoAnalogico`/
`ComandoAvulso`/`CanaisDistribuicao` a partir de entidades já importadas), não
existe hoje uma função simétrica que reconstrua a aba `IEDs` a partir de
`LSC`/`CNF`/etc. já existentes numa base importada — na planilha real do
usuário (que já tem 106 `LSC`/`CNF` reais) a aba `IEDs` continua vazia depois
de `extrair_pontos()`, porque essa função nunca olha pra essas entidades.
Simétrico à lacuna do 61850 `MUL`/`ENM` fechada acima, mas num escopo maior
(precisaria inferir `Protocolo` a partir de `LSC.TTP`, `Direcao` a partir de
`LSC.TIPO`, e desmontar `CNF.CONFIG` de volta em colunas — um "parser reverso"
por protocolo, o oposto do que `_montar_config_cnf`/`_CAMPOS_CNF_*` fazem hoje).

**Ainda pendente**: OPC UA e C37.118 — sem base real nem manual completo o
bastante disponível; retomar se aparecer uma base real ou um manual
equivalente ao que resolveu o ICCP.

### Ganhos rápidos (baixo esforço, alto retorno) — ✅ entregues
- **Troca de ID global** (`trocar_id_global`) — renomeia um ponto e propaga a todas as
  referências (usa o mesmo grafo `REGRAS_REFS_PADRAO` do verificador); origem:
  `fTrocaIdPDS` do Eletronorte‑2. Suporta lote/encadeamento; ID ambíguo ou inexistente
  não altera nada.
- **Estatística** (`estatistica_base`) — contagem de linhas totais/ativas por
  entidade.
- **Gestão de includes** (`gerir_includes`) — lista todos os includes e aplica
  substituições em lote no caminho (aba `SubstituirIncludes`).

Detalhes de cada um em [completa/README.md](completa/README.md).

## Princípios de design

1. **Genérico via configuração, não hard‑code.** As macros VBA originais são
   acopladas a um layout específico (named ranges fixos, índices de coluna, padrões
   de ID de uma concessionária). No SageBonis, **regras de verificação e padrões de
   naming moram em abas de config** (`MaisUsadas`, `EntidadeAtributoValor` ou novas),
   não no código. É isso que mantém a ferramenta genérica.
2. **Python idiomático.** Índices `dict`/`set`, funções puras testáveis, em vez de
   varreduras O(n²) e estado global do VBA.
3. **Não quebrar a exportação.** O formato `.dat` e o encoding ISO‑8859‑1 já
   funcionam; recursos novos não podem regredir isso.
4. **Sincronização macro ↔ .ods** continua via `sync_macro.py`, para as duas trilhas.

## Critérios de maturidade (para convergir)
A unificação numa planilha "modo duplo" só deve ser considerada quando a Trilha
Completa cumprir, no mínimo:
- [ ] As 3 famílias de funcionalidade acima estáveis e em uso real (estáveis: sim;
      em uso real no dia a dia: ainda não confirmado);
- [ ] Recursos avançados **desligáveis** (um "modo simples" que não atrapalhe quem
      só quer importar/exportar rápido);
- [x] Regras dirigidas por config (sem padrões de cliente hard‑coded) — FK/domínios
      via `EntidadeAtributoValor`/`VerificacaoRefs`; as formas por protocolo
      (`PARAMS_PROTOCOLO`) são estruturais (não variam por cliente), então ficam em
      código mesmo, por design;
- [x] Cobertura de teste mínima das funções puras (parser, geradores, verificador) —
      movida do scratchpad temporário pro repo em `completa/tests/` (antes só existia
      em `/tmp`, perdida entre sessões): 4 smoke tests em memória (~200+ checks) +
      2 testes UNO reais (harness reutilizável em `uno_harness.py`), todos
      rodáveis via `python completa/tests/run_all.py`;
- [x] Paridade de import/export com a Trilha Simples (mesmo resultado de `.dat`) —
      verificado de 2 formas: (1) diff de código confirma que o núcleo compartilhado
      (`importar_dats`/`exportar_dats`/parser, ~1000 linhas) é **byte-a-byte idêntico**
      entre `ImportadorSAGE.py` da raiz e da Completa, só divergindo depois do ponto
      onde a Completa anexa seus recursos próprios; (2) `teste_paridade_import_export.py`
      confirma empiricamente: importa uma base `.dat` sintética nas duas trilhas,
      exporta de volta, e os `.dat` resultantes são idênticos byte-a-byte (incluindo um
      caractere acentuado, pra testar o round-trip Latin-1). Esse teste funciona como
      guarda-corpo pra manter esse critério cumprido no futuro (se alguém editar um
      arquivo sem sincronizar o outro, o teste denuncia).

## Referências
Macros VBA de origem (fora deste repo), analisadas para extrair as ideias:
- `G:\Meu Drive\SAGE\Planilhas\Manipulação\Codigos Planilhas SAGE\eletronorte-2.txt`
  — verificação de base (`fAnálise`, `fAnálise<Entidade>`), troca de ID, includes.
- `G:\Meu Drive\SAGE\Planilhas\Manipulação\Codigos Planilhas SAGE\ge.txt`
  — unificação/distribuição (`distribuicao*`, `GeraDigital*`) e criação de
  protocolos (`CriaAquisicao_*`, `CriaAquisicao61850_*`, `CriaDistribuicao_*`).
