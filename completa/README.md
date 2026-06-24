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
- **Integridade referencial cruzada** — dirigida pela aba de config `VerificacaoRefs`.

#### Configurando a integridade referencial (`VerificacaoRefs`)
Na primeira execução do verificador, a aba `VerificacaoRefs` é criada com regras de
**exemplo inativas**. Cada linha é uma regra "o atributo X da entidade de origem deve
existir como atributo Y na entidade de destino":

| EntidadeOrigem | AtributoOrigem | EntidadeDestino | AtributoDestino | Ativa |
|----------------|----------------|-----------------|-----------------|-------|
| PDS | TAC | TAC | ID | N |
| PDD | PDS | PDS | ID | N |

Ajuste as regras conforme o padrão da **sua** base e mude `Ativa` para `S` nas que
quiser ligar. Só regras ativas são checadas. (Por isso a 1ª execução não gera erros
de referência — você ativa o que faz sentido.)

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
