# SageBonis - Editor Rápido para Base de Dados SAGE

**Versão atual:** 0.9.1

[cite_start]O SageBonis é uma ferramenta em LibreOffice Calc projetada para otimizar a edição da base de dados do **SAGE (Sistema Aberto de Gerenciamento de Energia)**. [cite: 7495] [cite_start]Ela importa múltiplos arquivos `.dat`, organiza as entidades em abas, permite a edição em massa e exporta a configuração de volta para múltiplos arquivos `.dat` de forma controlada. [cite: 7495]

[cite_start]A partir desta versão, o script lê dinamicamente as configurações de ordenação, cores e validação das abas "MaisUsadas" e "EntidadesValoresAtributos". [cite: 7495]

[cite_start]Adicionado a aba cores para facilitar a aplicação de temas de cores. Correção de importação de comentarios. [cite: 7495]

## Guia Rápido

### Passo 1: Instalar a Macro

Para que os botões da planilha funcionem, você precisa instalar o script `ImportadorSAGE.py`.

#### Método 1: Copiando o Arquivo (Recomendado)

Copie o arquivo `ImportadorSAGE.py` para a pasta de scripts do seu usuário no LibreOffice.

- **Windows:**

    1.  Abra o Executar (`Win + R`) e cole o caminho:

        ```
        %APPDATA%\LibreOffice\4\user\Scripts\python
        ```

    2.  Pressione Enter e cole o arquivo `.py` nesta pasta.

- **Linux:**

    1.  Abra o terminal e execute o comando para copiar o arquivo (ajuste o caminho de origem se necessário):

        Bash

        ```
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
    - [cite_start]Abra `SageBonis.ods`. [cite: 7495] [cite_start]Na aba **geral**, cole o caminho completo da pasta no campo correspondente. [cite: 7495]
    - Clique no botão **`Importar Arquivos .dat`**. A planilha irá processar os arquivos e criar/preencher as abas, aplicando cores e ordenação de acordo com as configurações da aba `MaisUsadas`.
    - Para importação parcial, preencha o campo na aba `geral` com as entidades desejadas ou selecione a aba da entidade e use o botão **`Importar Parcial`**.

2.  **Editar:**
    - Navegue pelas abas (`PDS`, `PDF`, `PDD`, etc.) para editar os dados.
    - A formatação em "estilo zebra" ajuda a visualizar as linhas de forma mais clara.
    - Utilize a **coluna "Gera"** para definir como cada linha será tratada na exportação (veja detalhes abaixo).

3.  **Exportar:**
    - [cite_start]Após a edição, clique no botão **`Exportar para .dat`** para exportar todas as entidades. [cite: 7495]
    - [cite_start]Para exportar apenas a aba ativa ou a lista de entidades na aba `geral`, use o botão **`Exportar Parcial`**. [cite: 7495]
    - [cite_start]Os arquivos finais (ex: `pds.dat`) serão salvos na pasta de destino na aba `geral`. [cite: 7495]

## Funcionalidades Dinâmicas

- **Ordenação Personalizada:** A macro lê a aba `MaisUsadas` para determinar a ordem de importação das abas e também a ordem de exibição das colunas de atributos, o que torna a visualização mais organizada.
- **Cores de Abas:** As cores de cada aba podem ser definidas na aba `MaisUsadas`, permitindo uma identificação visual rápida.
- **Efeito Zebra:** As linhas importadas são formatadas com cores alternadas para melhorar a legibilidade.

## A Coluna "Gera"

[cite_start]Esta coluna é o principal controle da exportação. [cite: 7495] [cite_start]Ela indica à macro o que fazer com cada linha de dados. [cite: 7495]

- [cite_start]`x`: **Exportar Ativo** - A linha será convertida em um bloco de configuração padrão e ativo no arquivo `.dat` final. [cite: 7495]
- [cite_start]`c`: **Exportar Comentado** - A linha será convertida em um bloco de configuração, mas todas as suas linhas serão comentadas com um ponto e vírgula (`;`). [cite: 7495] [cite_start]Útil para desativar pontos sem perdê-los. [cite: 7495]
- [cite_start]`n`: **Exportar como Simples Comentário** - A linha será exportada como uma única linha de comentário. [cite: 7495] [cite_start]O texto do comentário deve estar na coluna ao lado da coluna "Gera". [cite: 7495]
- [cite_start]`i`: **Include Ativo** - A linha será convertida um include, com o caminho inserido na coluna C. [cite: 7495]
- [cite_start]`u`: **Include Comentado** - A linha será convertida um include comentado, com o caminho inserido na coluna C. [cite: 7495] [cite_start]Útil para desativar includes sem perdê-los. [cite: 7495]
- [cite_start]`q`: **Ignora Linha** - A linha será ignorada na exportação. [cite: 7495]

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

## Objetivos futuros (roadmap)

- **Unificar abas de entidades** em grupos mais compactos (ex.: digital, analógico, comando, comunicações, sistemas, infos, cores, ocorrências, etc.) para reduzir o número de abas e acelerar a configuração de uma SE completa.
- **Importar uma base existente** para esse modelo unificado (converter DAT → planilhas do SageBonis) para reaproveitar bases já prontas.
- **Criar uma aba “simul”** para geração de scripts de simulação, mesmo que o formato ainda esteja em definição.

## Contato

- **Victor Bonesi**
- Dúvidas, bugs ou sugestões: (11) 95456-4510
