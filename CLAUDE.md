# SageBonis — Contexto do Projeto

## O que é
Ferramenta para editar em massa a base de dados do **SAGE** (Sistema Aberto de
Gerenciamento de Energia) usando o **LibreOffice Calc**. A planilha
`SageBonis.ods` importa vários arquivos `.dat`, organiza as entidades em abas,
permite edição em massa e exporta de volta para arquivos `.dat`.

Autor: Victor Bonesi · Repo: https://github.com/vbonesi/SageBonis · Licença: GPL.

## Arquitetura (peças do projeto)
- **`ImportadorSAGE.py`** — macro Python do LibreOffice. É o cérebro: parser de
  `.dat`, importação/exportação, formatação (zebra, cores), validação. Toda a
  lógica vive aqui.
- **`SageBonis.ods`** — a planilha (um ZIP no formato ODF). Contém as abas de
  trabalho **e** uma cópia embutida da macro em `Scripts/python/ImportadorSAGE.py`
  (location=document), além de menu/barra `SageBonis` próprios.
- **`README.md`** — guia de instalação e uso para o usuário final.
- **`sync_macro.py`** — utilitário para sincronizar `ImportadorSAGE.py` ↔ macro
  embutida no `.ods` sem abrir o LibreOffice (ver abaixo).
- **`temp_ods/`, `temp_ods_content.xml`** — artefatos de extração do `.ods`. São
  descartáveis (1,6 MB) e idealmente não deveriam estar versionados.

## Conceitos do domínio
- Abas de dados: `PDS`, `PDF`, `PDD`, etc. (uma por entidade SAGE).
- Abas de configuração (ignoradas na exportação): `Geral`, `MaisUsadas`,
  `EntidadeAtributoValor`, `opmsk`, `Cores`.
- **Coluna "Gera"** controla a exportação por linha: `x` ativo, `c` comentado,
  `n` comentário simples, `i` include, `u` include comentado, `q` ignora.
- Encoding dos `.dat`: exporta em `latin-1` (ISO-8859-1, padrão do SAGE), importa
  aceitando `latin-1` e `utf-8`.

## A macro embutida (descoberta importante)
A v0.9.2 (por Felipe Santos) provou que dá para **embutir a macro dentro do
`.ods`** e chamá-la com `location=document`. Vantagens: distribuição de um único
arquivo autossuficiente (sem instalação manual em `%APPDATA%`), e compatibilidade
com LibreOffice Flatpak. A v0.9.2 também tornou a busca de abas tolerante a
maiúsculas/minúsculas (`_get_sheet`) e corrigiu `Geral`.

## Fluxo de atualização da macro (sync_macro.py)
O `.ods` tem sua própria cópia da macro, que pode divergir do `ImportadorSAGE.py`
do disco. Para mantê-los em sincronia:

```bash
python sync_macro.py status    # mostra o diff entre o .py e a macro embutida
python sync_macro.py extract   # macro do .ods  -> ImportadorSAGE.py  (puxar)
python sync_macro.py inject    # ImportadorSAGE.py -> macro do .ods    (empurrar)
```

`inject` cria um `.ods.bak` e preserva a regra do ODF (mimetype como primeira
entrada e sem compressão). Decida qual lado é a fonte da verdade antes de gravar.

## Versão atual: 0.9.3
Resultado da reconciliação das três versões que existiam: a macro embutida 0.9.2
(de Felipe Santos) virou a base, e a sanitização `_sanitizar_para_latin1` foi
reincorporada na exportação. O `ImportadorSAGE.py` da raiz e a macro embutida no
`.ods` estão **idênticos** (mantidos em sync via `sync_macro.py`).

## Convenções de manutenção
- Fonte da verdade da lógica = `ImportadorSAGE.py` da raiz.
- Após editar a macro, rodar `python sync_macro.py inject` e commitar `.py` + `.ods` juntos.
- `temp_ods/` e `temp_ods_content.xml` são ignorados pelo git (artefatos de extração).
