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
documentação tem), **OPC UA** e **C37.118** — sem base real disponível pra
validar nenhum dos três; retomar se aparecer uma.

**Entregue**: **104**, **101**, **DNP3** e **MODBUS**, confirmados contra bases
reais (`conv_iccp104`/GRD para 104, aquisição e distribuição; base do próprio
usuário — SE Miracema/`neoenergia` — para 101, aquisição e distribuição;
`ctl_dnp_mdb`/DJ9E539 para DNP3 e `mdb_alat_calc`/MDB1 para MODBUS, ambos só
aquisição — distribuição de ambos extrapolada por consistência, sem base real
disponível). `PARAMS_PROTOCOLO` generalizado para cobrir as diferenças reais
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

**Ainda pendente**: SNMP, ICCP/SICCP — cada um precisa de pesquisa própria
(endereçamento, `CNF.CONFIG`, grupos NV1/NV2 específicos) antes de implementar,
seguindo o mesmo padrão.

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
- [ ] As 3 famílias de funcionalidade acima estáveis e em uso real;
- [ ] Recursos avançados **desligáveis** (um "modo simples" que não atrapalhe quem
      só quer importar/exportar rápido);
- [ ] Regras dirigidas por config (sem padrões de cliente hard‑coded);
- [ ] Cobertura de teste mínima das funções puras (parser, geradores, verificador);
- [ ] Paridade de import/export com a Trilha Simples (mesmo resultado de `.dat`).

## Referências
Macros VBA de origem (fora deste repo), analisadas para extrair as ideias:
- `G:\Meu Drive\SAGE\Planilhas\Manipulação\Codigos Planilhas SAGE\eletronorte-2.txt`
  — verificação de base (`fAnálise`, `fAnálise<Entidade>`), troca de ID, includes.
- `G:\Meu Drive\SAGE\Planilhas\Manipulação\Codigos Planilhas SAGE\ge.txt`
  — unificação/distribuição (`distribuicao*`, `GeraDigital*`) e criação de
  protocolos (`CriaAquisicao_*`, `CriaAquisicao61850_*`, `CriaDistribuicao_*`).
