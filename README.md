Sobre o SAGE
O SAGE (Sistema Aberto de Gerenciamento de Energia) é um software de supervisão e controle (SCADA/EMS) amplamente utilizado no setor elétrico brasileiro. Sua configuração é modular e baseada em arquivos de texto (.dat), que definem a estrutura da base de dados, incluindo equipamentos, medições, alarmes e outros pontos do sistema. A edição manual desses arquivos pode ser complexa e suscetível a erros, problema que o SageBonis busca solucionar.

Funcionalidades
Importação Automatizada: Lê uma pasta com arquivos .dat e importa todas as configurações.

Parsing Inteligente: Interpreta a estrutura dos arquivos, incluindo comandos !include para carregar dependências de forma recursiva.

Organização por Entidades: Separa cada tipo de entidade da base de dados (como EST, CDT, VAR, MED) em sua própria aba na planilha, facilitando a navegação e edição.

Interface Amigável: Utiliza a interface familiar de uma planilha para visualizar e manipular dados complexos.

Exportação Controlada: Permite ao usuário controlar quais blocos de configuração serão exportados para o arquivo .dat final, usando flags simples na coluna "Gera":

x: Exporta o bloco de configuração ativo.

c: Exporta o bloco de configuração como um comentário (desativado).

s: Exporta a linha como um comentário simples.

Pré-requisitos
LibreOffice Calc: A ferramenta foi desenvolvida e testada com o LibreOffice. É necessário ter o suporte a macros em Python ativado.

Python: O interpretador Python já vem integrado na maioria das instalações do LibreOffice.

Como Usar
1. Configuração do Ambiente
Baixe os arquivos: Clone ou faça o download deste repositório.

Instale a Macro:

Abra o LibreOffice Calc.

Vá em Ferramentas > Macros > Organizar macros > Python....

Na janela "Macros Python", clique em Minhas macros > Novo. Dê um nome ao módulo (ex: SageBonis) e clique em OK.

Isso abrirá o editor de scripts. Substitua o conteúdo do novo módulo pelo código do arquivo ImportadorSAGE.py.

Salve e feche o editor.

Associe as Macros aos Botões:

Abra o arquivo SageBonis.ods.

Clique com o botão direito no botão "Importar Arquivos .dat" e selecione Propriedades de controle > Eventos.

Associe o evento "Ação a ser executada" à macro importar_dats recém-adicionada.

Repita o processo para o botão "Exportar para .dat", associando-o à macro exportar_para_dat.

Salve a planilha.

2. Importando a Base de Dados
Coloque todos os arquivos .dat da sua base fonte do SAGE em uma única pasta.

Abra a planilha SageBonis.ods.

Na aba geral, insira o caminho completo da pasta que contém os arquivos .dat no campo "Caminho da Pasta".

Clique no botão Importar Arquivos .dat.

Aguarde o processo. A macro irá criar e preencher as abas com os dados importados. O status da operação será exibido na aba "geral".

3. Editando e Exportando
Navegue pelas abas (ex: EST, CDT, etc.) para visualizar e editar os dados.

Na coluna Gera de cada aba, utilize as flags x, c ou s para definir como cada item será tratado na exportação.

Após realizar as modificações, volte para a aba geral.

Clique no botão Exportar para .dat.

Um novo arquivo .dat será gerado na pasta especificada, contendo todas as suas alterações. O nome do arquivo de saída é definido na aba "geral".

Estrutura dos Arquivos
SageBonis.ods: A planilha do LibreOffice Calc, que serve como interface principal para o usuário.

ImportadorSAGE.py: O coração da ferramenta. Contém a lógica em Python para as macros de importação e exportação.

Manuais e Apostilas (/docs): Documentação oficial do SAGE para referência sobre a estrutura da base de dados e funcionamento do sistema.

Contato e Sugestões
Dúvidas, bugs ou sugestões? Entre em contato:

Victor Bonesi - (11) 95456-4510







gere ele mais enxuto, os profissionais da area ja tem familiaridade com essa planilha, então explique brevemente o que ela faz, descreva os atalhos, precauções, forma de uso, a coluna gera (serve para indicar o que é a linha na importação e como irá exportar tambem) (crie um guia rapido de como inserir a macro no caminho do libreoffice para windows e linux)




Mostrar raciocínio
Claro, aqui está uma versão mais enxuta e direta do README, focada nos profissionais que já conhecem o SAGE e precisam de um guia rápido para usar a ferramenta.

SageBonis - Editor Rápido para Base de Dados SAGE
O SageBonis é uma ferramenta em LibreOffice Calc projetada para otimizar a edição da base de dados do SAGE (Sistema Aberto de Gerenciamento de Energia). Ela importa múltiplos arquivos .dat, organiza as entidades em abas, permite a edição em massa e exporta a configuração de volta para um único arquivo .dat de forma controlada.

Guia Rápido
Passo 1: Instalar a Macro
Para que os botões da planilha funcionem, você precisa instalar o script ImportadorSAGE.py.

Método 1: Copiando o Arquivo (Recomendado)
Copie o arquivo ImportadorSAGE.py para a pasta de scripts do seu usuário no LibreOffice.

Windows:

Abra o Executar (Win + R) e cole o caminho:

%APPDATA%\LibreOffice\4\user\Scripts\python
Pressione Enter e cole o arquivo .py nesta pasta.

Linux:

Abra o terminal e execute o comando para copiar o arquivo (ajuste o caminho de origem se necessário):

Bash

cp ImportadorSAGE.py ~/.config/libreoffice/4/user/Scripts/python/
(Nota: O número 4 no caminho pode variar dependendo da sua versão do LibreOffice.)

Método 2: Pela Interface do LibreOffice
Abra o LibreOffice Calc.

Vá em Ferramentas > Macros > Organizar macros > Python....

Clique em Minhas macros e depois em Novo. Dê o nome ImportadorSAGE e clique em OK.

O editor de scripts abrirá. Apague todo o conteúdo e cole o código do arquivo ImportadorSAGE.py no lugar.

Salve (Ctrl + S) e feche o editor.

Passo 2: Uso da Planilha
Importar:

Coloque todos os seus arquivos .dat em uma pasta.

Abra SageBonis.ods. Na aba geral, cole o caminho completo da pasta no campo correspondente.

Clique no botão Importar Arquivos .dat. A planilha irá processar os arquivos e criar/preencher as abas.

Editar:

Navegue pelas abas (EST, CDT, VAR, etc.) para editar os dados.

Utilize a coluna "Gera" para definir como cada linha será tratada na exportação (veja detalhes abaixo).

Exportar:

Após a edição, clique no botão Exportar para .dat.

O arquivo final (ex: SAIDA.dat) será salvo na mesma pasta dos arquivos de origem.

A Coluna "Gera"
Esta coluna é o principal controle da exportação. Ela indica à macro o que fazer com cada linha de dados.

x: Exportar Ativo - A linha será convertida em um bloco de configuração padrão e ativo no arquivo .dat final.

c: Exportar Comentado - A linha será convertida em um bloco de configuração, mas todas as suas linhas serão comentadas com um ponto e vírgula (;). Útil para desativar pontos sem perdê-los.

s: Exportar como Simples Comentário - A linha será exportada como uma única linha de comentário. O texto do comentário deve estar na coluna ao lado da coluna "Gera".

Se a célula na coluna "Gera" estiver vazia, a linha será ignorada durante a exportação.

Precauções e Boas Práticas
⚠️ Backup é Essencial: Sempre faça uma cópia de segurança dos seus arquivos .dat originais antes de começar a usar a ferramenta.

Sobrescrita de Arquivos: A função de exportação sobrescreve o arquivo de saída (definido na aba geral) sem aviso. Tenha certeza do que está fazendo.

Caminho Absoluto: Use o caminho completo (absoluto) para a pasta dos arquivos .dat para evitar erros.

Revisão: Antes de exportar, revise a coluna "Gera" para garantir que apenas os pontos desejados estão marcados com x ou c.

Contato
Victor Bonesi

Dúvidas, bugs ou sugestões: (11) 95456-4510
