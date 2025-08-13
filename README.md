# SageBonis - Editor Rápido para Base de Dados SAGE

O SageBonis é uma ferramenta em LibreOffice Calc projetada para otimizar a edição da base de dados do **SAGE (Sistema Aberto de Gerenciamento de Energia)**. Ela importa múltiplos arquivos `.dat`, organiza as entidades em abas, permite a edição em massa e exporta a configuração de volta para um único arquivo `.dat` de forma controlada.

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
        
    - Abra `SageBonis.ods`. Na aba **geral**, cole o caminho completo da pasta no campo correspondente.
        
    - Clique no botão **`Importar Arquivos .dat`**. A planilha irá processar os arquivos e criar/preencher as abas.
        
2.  **Editar:**
    
    - Navegue pelas abas (`PDS`, `PDF`, `PDD`, etc.) para editar os dados.
        
    - Utilize a **coluna "Gera"** para definir como cada linha será tratada na exportação (veja detalhes abaixo).
        
3.  **Exportar:**
    
    - Após a edição, clique no botão **`Exportar para .dat`**.
        
    - O arquivo final (ex: `pds.dat`) será salvo na pasta de destino na aba geral.
        

## A Coluna "Gera"

Esta coluna é o principal controle da exportação. Ela indica à macro o que fazer com cada linha de dados.

- `x`: **Exportar Ativo** - A linha será convertida em um bloco de configuração padrão e ativo no arquivo `.dat` final.
    
- `c`: **Exportar Comentado** - A linha será convertida em um bloco de configuração, mas todas as suas linhas serão comentadas com um ponto e vírgula (`;`). Útil para desativar pontos sem perdê-los.
    
- `s`: **Exportar como Simples Comentário** - A linha será exportada como uma única linha de comentário. O texto do comentário deve estar na coluna ao lado da coluna "Gera".
    

Se a célula na coluna "Gera" estiver vazia, a linha será ignorada durante a exportação.

## Precauções e Boas Práticas

- ⚠️ **Backup é Essencial:** Sempre faça uma cópia de segurança dos seus arquivos `.dat` originais antes de começar a usar a ferramenta.
    
- **Sobrescrita de Arquivos:** A função de exportação **sobrescreve** o arquivo de saída (definido na aba `geral`) sem aviso. Tenha certeza do que está fazendo.
    
- **Caminho Absoluto:** Use o caminho completo (absoluto) para a pasta dos arquivos `.dat` para evitar erros.
    
- **Revisão:** Antes de exportar, revise a coluna "Gera" para garantir que apenas os pontos desejados estão marcados com `x` ou `c`.
    

## Contato

- **Victor Bonesi**
    
- Dúvidas, bugs ou sugestões: (11) 95456-4510
