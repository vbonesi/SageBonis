# SageBonis — Trilha Completa (em desenvolvimento)

Esta é a variante **Completa** do SageBonis: um fork da planilha/macro
[Simples](../README.md) que mantém todo o import/export e vai acumulando recursos
avançados. A estratégia das duas trilhas está em [PLANEJAMENTO.md](../PLANEJAMENTO.md).

> A planilha **Simples** (na raiz do repo) continua sendo a recomendada para quem
> quer só importar/exportar base rápido. Use a Completa se quiser os recursos abaixo.

## Recursos além da Simples

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
  [Prefixo/Sufixo/Substituir], Valor1, Valor2, Ativo`). Substitui os "4 slots fixos"
  que macros de referência (GE) hard-codificam por quantos canais fizerem sentido.
- **`DistribuicaoPontos`** — liga um `ID_Logico` a 1+ canais (`ID_Logico, Canal,
  Ativo`). Um ponto sem nenhuma linha aqui simplesmente não gera distribuição.

## Instalação e uso
Igual à Simples: abra `SageBonis.ods` e habilite as macros do documento (a macro vem
embutida). Atribua as funções `verificar_base` e `unificar_pontos` a botões ou
atalhos, como as demais.

## Sincronizar a macro com o .ods
A partir da raiz do repositório:

```bash
python sync_macro.py inject  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
python sync_macro.py status  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
```

## Status
🚧 Em desenvolvimento. Próximo item do roadmap (ver [PLANEJAMENTO.md](../PLANEJAMENTO.md)):
assistente de protocolo/IED.
