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
A partir de **IED + protocolo (DNP3/101/104/61850) + tipo**, gera o esqueleto padrão
de aquisição/distribuição (ex.: para DNP3, as linhas ASIM/APFL/ADUP/CSIM/CDUP).
Escopo bem definido; depende de modelar NV1/NV2.

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
