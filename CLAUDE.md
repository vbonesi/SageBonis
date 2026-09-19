# SageBonis — Contexto do Projeto

## O que é
Ferramenta para editar em massa a base de dados do **SAGE** (Sistema Aberto de
Gerenciamento de Energia) usando o **LibreOffice Calc**. A planilha `SageBonis.ods`
importa vários arquivos `.dat`, organiza as entidades em abas, permite edição em massa
e exporta de volta para arquivos `.dat`.

Autor: Victor Bonesi · Repo: https://github.com/vbonesi/SageBonis · Licença: GPLv3.

## Escopo (importante)
O projeto faz **uma coisa**: import/export de base `.dat` com edição em massa no meio.
Recursos avançados — verificador de base, unificação de pontos, assistente de
protocolo/IED — foram desenvolvidos como variante "Completa" e, **desde 19/09/2026**,
seguem como ferramenta interna da Automa (`SAGEAutoma`, repositório privado). Aqui não
se empilha funcionalidade avançada: o valor desta planilha é ser rápida e previsível.

Os dois núcleos **divergem de propósito** — não há sincronização automática entre os
repositórios. Melhoria de import/export feita lá pode ser trazida pra cá à mão, e
vice-versa.

## Arquitetura
- **`ImportadorSAGE.py`** — a macro do LibreOffice. É o cérebro: parser de `.dat`,
  importação/exportação, formatação (zebra, cores), diagnóstico. Toda a lógica vive aqui.
- **`SageBonis.ods`** — a planilha (ZIP no formato ODF). Contém as abas de trabalho, uma
  cópia embutida da macro em `Scripts/python/ImportadorSAGE.py` (location=document) e o
  menu/barra `SageBonis`.
- **`sync_macro.py`** — sincroniza `ImportadorSAGE.py` ↔ macro embutida no `.ods` sem
  abrir o LibreOffice.
- **`tests/`** — smoke test do parser (rápido, sem LibreOffice) e round-trip
  import/export via UNO real. `python tests/run_all.py`.
- **`.github/workflows/testes.yml`** — roda a suíte a cada push/PR e reprova macro fora
  de sincronia com o `.ods`.

## Conceitos do domínio
- Abas de dados: `PDS`, `PDF`, `PDD`, etc. (uma por entidade SAGE).
- Abas de configuração (ignoradas na exportação): `Geral`, `MaisUsadas`,
  `EntidadeAtributoValor`, `opmsk`, `Cores`. **São protegidas com senha** — o LibreOffice
  ignora em silêncio escrita de cor/conteúdo nelas via macro.
- **Coluna "Gera"** controla a exportação por linha: `x` ativo, `c` comentado, `n`
  comentário simples, `i` include, `u` include comentado, `q` ignora.
- Encoding dos `.dat`: exporta em `latin-1` (ISO-8859-1, padrão do SAGE), importa
  tentando `utf-8` antes de `latin-1`.
- Lista de entidades da importação/exportação parcial: coluna C da aba `Geral`, da linha
  14 até o fim da área usada.

## Fluxo de atualização da macro
O `.ods` tem sua própria cópia da macro, que pode divergir do `ImportadorSAGE.py`:

```bash
python sync_macro.py status    # mostra o diff (sai com código 1 se divergirem)
python sync_macro.py extract   # macro do .ods  -> ImportadorSAGE.py  (puxar)
python sync_macro.py inject    # ImportadorSAGE.py -> macro do .ods    (empurrar)
```

`inject` cria um `.ods.bak` e preserva a regra do ODF (mimetype como primeira entrada e
sem compressão). Decida qual lado é a fonte da verdade antes de gravar.

## Convenções de manutenção
- Fonte da verdade da lógica = `ImportadorSAGE.py` da raiz.
- Após editar a macro, rodar `python sync_macro.py inject` e commitar `.py` + `.ods`
  juntos — a CI reprova se ficarem diferentes.
- `auditoria/` fica fora do Git (relatórios e execução ficam só na cópia em nuvem).
