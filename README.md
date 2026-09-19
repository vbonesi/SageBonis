# SageBonis - Editor Rápido para Base de Dados SAGE

**Versão atual:** 0.9.3

O SageBonis é uma ferramenta em LibreOffice Calc projetada para otimizar a edição da base de dados do **SAGE (Sistema Aberto de Gerenciamento de Energia)**. Ela importa múltiplos arquivos `.dat`, organiza as entidades em abas, permite a edição em massa e exporta a configuração de volta para múltiplos arquivos `.dat` de forma controlada.

A partir da versão 0.9.x, o script lê dinamicamente as configurações de ordenação, cores e validação das abas "MaisUsadas" e "EntidadesValoresAtributos".

**Novidades recentes:**

- **0.9.2** — A macro agora pode vir **embutida no próprio arquivo `SageBonis.ods`** (executada com `location=document`), dispensando a instalação manual. Compatível com LibreOffice Flatpak. Busca de abas tolerante a maiúsculas/minúsculas e tratamento de erros mais seguro. Menu e barra de ferramentas `SageBonis` próprios dentro do ODS.
- **0.9.3** — Reincorporada a sanitização para ISO-8859-1 na exportação, evitando erros quando há caracteres Unicode (aspas curvas, travessões, etc.) que não existem no encoding esperado pelo SAGE.

## Guia Rápido

### Passo 1: Instalar a Macro

Para que os botões da planilha funcionem, a macro `ImportadorSAGE.py` precisa estar disponível. Há três caminhos:

#### Método 0: Macro embutida no `.ods` (Mais simples)

A partir da versão 0.9.2, o arquivo `SageBonis.ods` **já traz a macro embutida** (em `Scripts/python/ImportadorSAGE.py`, executada com `location=document`). Basta abrir a planilha e, quando o LibreOffice perguntar, **habilitar as macros do documento**. Não é preciso copiar nada para a pasta do usuário.

> Para confiar nas macros do documento sem o aviso a cada abertura, ajuste em `Ferramentas > Opções > LibreOffice > Segurança > Segurança de macros`.

#### Método 1: Copiando o Arquivo

Copie o arquivo `ImportadorSAGE.py` para a pasta de scripts do seu usuário no LibreOffice.

- **Windows:**

    1.  Abra o Executar (`Win + R`) e cole o caminho:

        ```
        %APPDATA%\LibreOffice\4\user\Scripts\python
        ```

    2.  Pressione Enter e cole o arquivo `.py` nesta pasta.

- **Linux:**

    1.  Abra o terminal e execute o comando para copiar o arquivo (ajuste o caminho de origem se necessário):

        ```bash
        cp ImportadorSAGE.py ~/.config/libreoffice/4/user/Scripts/python/
        ```

*(Nota: O número `4` no caminho pode variar dependendo da sua versão do LibreOffice.)*

#### Método 2: Pela Interface do LibreOffice

1.  Abra o LibreOffice Calc.
2.  Vá em `Ferramentas > Macros > Organizar macros > Python...`.
3.  Clique em `Minhas macros` e depois em `Novo`. Dê o nome `ImportadorSAGE` e clique em OK.
4.  O editor de scripts abrirá. Apague todo o conteúdo e cole o código do arquivo `ImportadorSAGE.py` no lugar.
5.  Salve (`Ctrl + S`) e feche o editor.

### Passo 2: Uso da Planilha

1.  **Importar:**
    - Coloque todos os seus arquivos `.dat` em uma pasta.
    - Abra `SageBonis.ods`. Na aba **geral**, cole o caminho completo da pasta no campo correspondente.
    - Clique no botão **`Importar Arquivos .dat`**. A planilha irá processar os arquivos e criar/preencher as abas, aplicando cores e ordenação de acordo com as configurações da aba `MaisUsadas`.
    - Para importação parcial, preencha o campo na aba `geral` com as entidades desejadas ou selecione a aba da entidade e use o botão **`Importar Parcial`**.

2.  **Editar:**
    - Navegue pelas abas (`PDS`, `PDF`, `PDD`, etc.) para editar os dados.
    - A formatação em "estilo zebra" ajuda a visualizar as linhas de forma mais clara.
    - Utilize a **coluna "Gera"** para definir como cada linha será tratada na exportação (veja detalhes abaixo).

3.  **Exportar:**
    - Após a edição, clique no botão **`Exportar para .dat`** para exportar todas as entidades.
    - Para exportar apenas a aba ativa ou a lista de entidades na aba `geral`, use o botão **`Exportar Parcial`**.
    - Os arquivos finais (ex: `pds.dat`) serão salvos na pasta de destino na aba `geral`.

## Funcionalidades Dinâmicas

- **Ordenação Personalizada:** A macro lê a aba `MaisUsadas` para determinar a ordem de importação das abas e também a ordem de exibição das colunas de atributos, o que torna a visualização mais organizada.
- **Cores de Abas:** As cores de cada aba podem ser definidas na aba `MaisUsadas`, permitindo uma identificação visual rápida.
- **Efeito Zebra:** As linhas importadas são formatadas com cores alternadas para melhorar a legibilidade.

## A Coluna "Gera"

Esta coluna é o principal controle da exportação. Ela indica à macro o que fazer com cada linha de dados.

- `x`: **Exportar Ativo** - A linha será convertida em um bloco de configuração padrão e ativo no arquivo `.dat` final.
- `c`: **Exportar Comentado** - A linha será convertida em um bloco de configuração, mas todas as suas linhas serão comentadas com um ponto e vírgula (`;`). Útil para desativar pontos sem perdê-los.
- `n`: **Exportar como Simples Comentário** - A linha será exportada como uma única linha de comentário. O texto do comentário deve estar na coluna ao lado da coluna "Gera".
- `i`: **Include Ativo** - A linha será convertida em um include, com o caminho inserido na coluna C.
- `u`: **Include Comentado** - A linha será convertida em um include comentado, com o caminho inserido na coluna C. Útil para desativar includes sem perdê-los.
- `q`: **Ignora Linha** - A linha será ignorada na exportação.

Se a célula na coluna "Gera" estiver vazia, a linha será ignorada durante a exportação.

## Aba `opmsk`

A planilha também contém uma aba auxiliar chamada `opmsk`, que pode ser usada para facilitar o cálculo e a configuração das máscaras de bits do protocolo 61850.

## Precauções e Boas Práticas

- ⚠️ **Backup é Essencial:** A função de exportação **sobrescreve** o arquivo de saída, mas cria um backup (`.bak`) da versão anterior na mesma pasta de destino.
- **Caminho Absoluto:** Use o caminho completo (absoluto) para a pasta dos arquivos `.dat` para evitar erros.
- **Revisão:** Antes de exportar, revise a coluna "Gera" para garantir que apenas os pontos desejados estão marcados com `x` ou `c`.

## Como Configurar Atalhos de Teclado no LibreOffice

Para agilizar seu fluxo de trabalho, você pode associar as macros a atalhos de teclado.

1.  No LibreOffice Calc, vá em `Ferramentas > Personalizar...`.
2.  Na janela que se abrir, selecione a aba `Teclado`.
3.  No campo `Categoria`, procure por `Macros do LibreOffice`.
4.  No campo `Macro`, expanda `Minhas Macros > ImportadorSAGE` e selecione uma das funções (por exemplo, `importar_dats`).
5.  No campo `Teclas de atalho`, selecione o atalho desejado (`Ctrl + Shift + S` para `importar_dats`).
6.  Clique no botão `Modificar`.
7.  Repita o processo para as outras macros:
    - `importar_parcial` -> `Ctrl + Shift + D`
    - `exportar_dats` -> `Ctrl + Shift + W`
    - `exportar_parcial` -> `Ctrl + Shift + E`
8.  Clique em `OK` para salvar as configurações.

## Para Desenvolvedores: sincronizar a macro embutida

Como o `.ods` é um arquivo ZIP, ele guarda sua própria cópia da macro em
`Scripts/python/ImportadorSAGE.py`. Essa cópia pode divergir do
`ImportadorSAGE.py` da raiz do projeto. O utilitário `sync_macro.py` mantém os
dois em sincronia **sem precisar abrir o LibreOffice**:

```bash
python sync_macro.py status    # mostra o diff entre o .py e a macro embutida
python sync_macro.py extract   # macro do .ods  -> ImportadorSAGE.py  (puxar p/ disco)
python sync_macro.py inject    # ImportadorSAGE.py -> macro do .ods    (empurrar p/ planilha)
```

Fluxo recomendado: edite `ImportadorSAGE.py`, rode `python sync_macro.py inject`
(ele cria um `.ods.bak` automático e preserva a estrutura do ODF) e versione os
dois arquivos juntos no mesmo commit. O `ImportadorSAGE.py` da raiz é a fonte da
verdade.

## Roadmap

O SageBonis é focado em fazer bem uma coisa: **importar a base `.dat` para a planilha,
editar em massa e exportar de volta**, sem atrito e sem surpresa no formato. As
melhorias daqui vêm nessa direção — parser, encoding, desempenho, diagnóstico claro
quando algo dá errado.

> **Sobre os recursos avançados:** verificador de base, unificação de pontos e
> assistente de protocolo/IED foram desenvolvidos como uma variante deste projeto e,
> desde 09/2026, seguem como ferramenta interna da [Automa](https://automa.com.br)
> (`SAGEAutoma`), onde nasceram e são usados. O SageBonis continua aberto sob GPLv3 com
> a parte de import/export, e recebe de volta as melhorias que aparecerem nessa parte.

## Testes

```bash
python tests/run_all.py            # tudo
python tests/run_all.py --sem-uno  # só o smoke test (sem LibreOffice)
```

- `tests/smoke_test_parser.py` — o parser de `.dat` contra fixtures: bloco ativo e
  comentado, entidade com dígito no nome, `utf-8` e `latin-1`, CRLF, includes e
  comentários soltos. Roda em segundos, sem LibreOffice.
- `tests/teste_uno_roundtrip.py` — sobe um `soffice --headless`, importa uma base
  sintética numa cópia descartável do `.ods` e exporta de volta, conferindo o que a
  ferramenta promete: acento em Latin-1, includes, blocos comentados, linhas ignoradas,
  backup `.bak` e exportação parcial. **Nunca escreve no `.ods` rastreado.**

Tudo isso roda sozinho no GitHub Actions a cada push/PR — inclusive uma checagem que
reprova quem editar o `ImportadorSAGE.py` e esquecer o `sync_macro.py inject`.

## Contato

- **Victor Bonesi**
- Dúvidas, bugs ou sugestões: (11) 95456-4510
