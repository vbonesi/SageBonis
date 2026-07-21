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

## Instalação e uso
Igual à Simples: abra `SageBonis.ods` e habilite as macros do documento (a macro vem
embutida). Atribua a função `verificar_base` a um botão ou atalho, como as demais.

## Sincronizar a macro com o .ods
A partir da raiz do repositório:

```bash
python sync_macro.py inject  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
python sync_macro.py status  --ods completa/SageBonis.ods --py completa/ImportadorSAGE.py
```

## Status
🚧 Em desenvolvimento. Próximos itens do roadmap (ver [PLANEJAMENTO.md](../PLANEJAMENTO.md)):
unificação de pontos (Digital/Analógico/Comando → fan-out) e assistente de protocolo/IED.
