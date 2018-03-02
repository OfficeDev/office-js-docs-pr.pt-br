---
title: Parceiros de interação para Suplementos do Office
description: ''
ms.date: 12/04/2017
---



# <a name="interaction-patterns-for-office-add-ins"></a>Parceiros de interação para Suplementos do Office


Os Suplementos do Office podem aprimorar as experiências de criação e a produtividade, bem como conectar conteúdo em aplicativos de host do Office para maiores fluxos de trabalho baseados na Web. Uma série de cenários comuns se aplicam ao conteúdo, painel de tarefas e suplementos do Outlook que você possa desenvolver. Este artigo descreve alguns dos cenários mais comuns e fornece padrões de interação recomendados para a experiência do usuário do suplemento. Você pode detalhar, combinar ou misturar e relacionar esses padrões de interação de acordo com seus cenários exclusivos.

 **Cenários comuns de suplemento**

| Tipo de suplemento | Cenários comuns |
| ------ | ------ |
|  Conteúdo  |  Visualização de dados <br> Widgets e ferramentas  |
|  Painel de tarefas  |  Transformar e processar dados <br> Criação de forma eficiente e eficaz <br> Localizar conteúdo e inserir dados <br> Publicar ou carregar conteúdo para um serviço Web  |
|  Outlook  |  Ponte entre o conteúdo de email e um aplicativo externo <br> Fornecer mais informações sobre o conteúdo em um compromisso ou uma mensagem de email <br> Fornecer informações que o ajudem a ser mais produtivo  |

## <a name="visualize-data-with-a-content-add-in"></a>Visualizar dados com um suplemento de conteúdo


Este exemplo mostra um suplemento de conteúdo para Excel que gera um gráfico a partir dos dados em uma planilha.

Nesse padrão de interação, o suplemento não se torna ativo até que você selecione e vincule dados para gerar o gráfico. É importante comunicar a finalidade do suplemento e as etapas para ativá-lo no modo de exibição inicial do suplemento. 

**Suplemento de conteúdo para Excel que gera um gráfico a partir dos dados em uma planilha**
<br>
![Aplicativo de conteúdo para Excel que gera um gráfico com base nos dados em uma planilha](../../images/office15-app-ux-fig01.png)
<br>
<ul><li><p>Para reforçar que você deve executar uma ação antes de escolher o botão, exiba instruções juntamente com um botão desabilitado (A).</p></li><li><p>Depois que você selecionar um intervalo de células, o botão <span class="ui">Criar Gráfico</span> ficará ativo (B - C).</p></li><li><p>A visualização preenche o contêiner e substitui o modo de exibição anterior (D).</p></li><li><p>Exiba qualquer interface do usuário adicional na borda inferior do suplemento junto com um botão de configurações (engrenagem) para levá-lo a um modo de exibição em que você pode redefinir ou gerenciar o suplemento.</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos que exigem que você selecione dados antes da ativação.</p></li></ul>

## <a name="transform-content-with-a-task-pane-add-in"></a>Transformar o conteúdo com um suplemento do painel de tarefas


Este exemplo mostra um suplemento do painel de tarefas que traduz o texto no seu documento para outro idioma.

Nesse padrão de interação, você deve primeiro selecionar o texto que deseja traduzir no documento.

**Suplemento do painel de tarefas que traduz o texto no seu documento para outro idioma**
<br>
![Aplicativo do painel de tarefas que traduz o texto no seu documento para outro idioma](../../images/office15-app-ux-fig02.png)
<br>
<ul><li><p>Comunique a finalidade do suplemento em um título e mencione o fato de que primeiro você deve fazer uma seleção (A).</p></li><li><p>O menu de idioma e o botão <span class="ui">Traduzir</span> estão desabilitados, reforçando que você deve executar uma ação antes para poder continuar. Depois que você selecionar o conteúdo do documento, esses dois elementos ficarão ativos (D).</p></li><li><p>Depois de escolher <span class="ui">Traduzir</span>, a IU será aberta, mostrando o conteúdo traduzido, juntamente com um botão para inseri-lo novamente no documento (R).</p></li><li><p>Você pode fornecer um botão <span class="ui">Limpar</span> ou <span class="ui">Redefinir</span> que retorna à exibição inicial.</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos que exigem que você selecione dados antes da ativação.</p></li><li><p>Interface do usuário que é aberta ou revelada à medida que você avança em um cenário.</p></li></ul>

## <a name="process-data-with-a-task-pane-add-in"></a>Processar dados com um suplemento do painel de tarefas


Este exemplo mostra um suplemento do painel de tarefas que verifica os dados no Excel.

Nesse padrão de interação, você deve selecionar um intervalo de células na planilha para começar.

**Suplemento do painel de tarefas que verifica dados no Excel**
<br>
![Aplicativo do painel de tarefas que verifica dados no Excel](../../images/office15-app-ux-fig03.png)
<br>
<ul><li><p>A finalidade do suplemento é descrita no cabeçalho. Instruções ajudam você a começar.</p></li><li><p>O botão <span class="ui">Enviar dados selecionados</span> está desabilitado, reforçando que você deve executar uma ação para avançar (A).</p></li><li><p>Depois de selecionar um intervalo de células na planilha (B), o botão <span class="ui">Enviar dados selecionados</span> é ativado.</p></li><li><p>Depois de escolher esse botão, a IU é substituída pelo intervalo de células selecionado, uma barra de progresso e um botão <span class="ui">Cancelar</span>.</p></li><li><p>A barra de progresso comunica o status do processo e o botão <span class="ui">Cancelar</span> permite interrompê-lo (D).</p></li><li><p>Quando o processo é concluído, os resultados são exibidos automaticamente (R). Selecionar um elemento na lista ativará a célula correspondente na planilha.</p></li></ul>Mais adequado para:
<ul><li><p>Processos que levam um tempo indeterminado.</p></li></ul>

## <a name="analyze-content-with-a-task-pane-add-in"></a>Analisar o conteúdo com um suplemento do painel de tarefas


Este exemplo mostra um suplemento do painel de tarefas que exibe definições de palavras conforme você digita.

Nesse padrão de interação, você deve primeiro selecionar o texto no documento para ver os resultados.

**Suplemento do painel de tarefas que exibe definições de palavras conforme você digita**
<br>
![Aplicativo de painel de tarefas que exibe definições de palavras conforme você digita](../../images/office15-app-ux-fig04.png)
<br>
<ul><li><p>Um título explica a finalidade do suplemento e como começar (A).</p></li><li><p>A pesquisa automática é habilitada por padrão, com a opção para desabilitá-la (B).</p></li><li><p>Depois que você faz uma seleção, o suplemento exibe o conteúdo correspondente (D).</p></li><li><p>Fornecer um link para exibir mais informações (E).</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos que retornam automaticamente o conteúdo à medida que você cria.</p></li><li><p>Suplementos que exigem que você selecione conteúdo antes da ativação.</p></li></ul>

## <a name="locate-content-with-a-task-pane-add-in"></a>Localizar o conteúdo com um suplemento do painel de tarefas


Este exemplo mostra um suplemento do painel de tarefas para a pesquisa de conteúdo.

Nesse padrão de interação, você insere uma cadeia de caracteres na caixa de pesquisa ou seleciona de uma lista de conteúdos em destaque para começar.

**Suplemento do painel de tarefas para pesquisar conteúdo**
<br>
![Aplicativo do painel de tarefas que pesquisa por conteúdo](../../images/office15-app-ux-fig05.png)
<br>
<ul><li><p>A exibição inicial contém a caixa <span class="ui">Pesquisar</span> (A) e uma lista de conteúdo em destaque (B).</p></li><li><p>Quando você insere uma cadeia de caracteres na caixa de pesquisa, o ícone de pesquisa é substituído por um ícone fechar (C).</p></li><li><p>A escolha do ícone fechar o leva de volta ao modo de exibição inicial.</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos que retornam automaticamente o conteúdo à medida que você cria.</p></li><li><p>Suplementos que exigem que você selecione conteúdo antes da ativação.</p></li></ul>

## <a name="insert-media-with-a-task-pane-add-in"></a>Inserir mídia com um suplemento do painel de tarefas


Nesse padrão de interação, é possível selecionar uma imagem dos resultados da pesquisa para inserir em seu documento.

**Suplemento do painel de tarefas para inserir uma imagem**
<br>
![Aplicativo do painel de tarefas para inserir uma imagem](../../images/office15-app-ux-fig06.png)
<br>
<ul><li><p>Você filtrou uma lista de retornos de pesquisa (A) e selecionou o conteúdo a ser inserido (B).</p></li><li><p>Um modo de exibição Detalhes do conteúdo selecionado é exibido (C) com um botão que o leva de volta à lista.</p></li><li><p>Um botão <span class="ui">Inserir Foto</span> está localizado no rodapé (D). Depois que você escolhe esse botão, a imagem é inserida no documento.</p></li><li><p>Uma breve descrição de onde veio a imagem é incluída no conteúdo inserido (R). </p></li><li><p>A interface do usuário no suplemento confirma visualmente o êxito da ação.</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos para inserir conteúdo.</p></li></ul>

## <a name="insert-selected-text-with-a-task-pane-add-in"></a>Inserir o texto selecionado com um suplemento do painel de tarefas


Neste padrão de interação, você seleciona um texto dos resultados da pesquisa para inserir no documento.

**Suplemento do painel de tarefas para inserir um texto**
<br>
![Aplicativo do painel de tarefas para inserir texto](../../images/office15-app-ux-fig07.png)
<br>
<ul><li><p>Você já localizou uma parte do conteúdo (A).</p></li><li><p>Um botão <span class="ui">Inserir Seleção</span> desabilitado é exibido no rodapé (B).</p></li><li><p>Quando você seleciona uma cadeia de caracteres de texto (C), o botão <span class="ui">Inserir Seleção</span> fica ativo.</p></li><li><p>Depois que você escolhe esse botão, o suplemento insere o texto selecionado no documento, com uma referência para a fonte do conteúdo (R).</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos para realizar pesquisas e inserir conteúdo.</p></li></ul>

## <a name="publish-to-a-web-service-with-a-task-pane-add-in"></a>Publicar em um serviço Web com um suplemento do painel de tarefas


Este exemplo mostra um suplemento do painel de tarefas para publicar um documento como postagem de blog.

Nesse padrão de interação, você concluiu a gravação do conteúdo em um documento e deseja postá-lo no seu blog.

**Suplemento do painel de tarefas para publicar um documento como uma postagem de blog**
<br>
![Aplicativo do painel de tarefas para publicação de documento como postagem de blog](../../images/office15-app-ux-fig08.png)
<br>
<ul><li><p>Primeiro, um formulário de entrada é exibido para inserir suas credenciais (A).</p></li><li><p>Links para criar uma conta e lidar com problemas de entrada típicos são fornecidos (B). A escolha desses links abre estas páginas em um navegador.</p></li><li><p>Depois que você está conectado, o suplemento exibe um formulário para criar uma nova postagem de blog (C).</p></li><li><p>O nome da conta em que você entrou (e na qual postará) é mostrada na parte superior do suplemento. O suplemento fornece um link para visualizar a postagem (D). A escolha desse link exibe a visualização em um navegador.</p></li><li><p>Depois que você escolhe <span class="ui">Criar postagem</span>, o suplemento exibe uma exibição confirmando que o conteúdo do documento foi postado (R).</p></li><li><p>O suplemento fornece um link para exibir a postagem em um navegador (F), bem como um botão para criar outra postagem (G).</p></li></ul>Mais adequado para:
<ul><li><p>Suplementos que geram, publicam ou compartilham conteúdo em redes sociais, sites de blog e serviços Web.</p></li><li><p>Suplementos que exigem que você entre em um serviço.</p></li></ul>

## <a name="get-more-information-about-people-with-an-outlook-add-in"></a>Obter mais informações sobre pessoas com um suplemento do Outlook


 **Exemplo 1**

**Página de resultados e detalhes**
<br>
![Página de resultados e detalhes](../../images/office15-app-ux-fig09.jpg)
<br>
Mais adequado para:
<ul><li><p>Exponha a abrangência de seu conteúdo, se você tiver grandes conjuntos de dados que sejam úteis para apresentação.</p></li><li><p>Páginas de detalhes que usam o tamanho completo do contêiner de suplemento</p></li><li><p>Modelos de navegação que se beneficiam de um fluxo de "trocas".</p></li></ul>
 
 **Exemplo 2**

**Página de detalhes com navegação persistente**
<br>
![Página de detalhes com navegação persistente](../../images/office15-app-ux-fig10.jpg)
<br>
Mais adequado para:
<ul><li><p>Exibir, por padrão, o primeiro resultado de um conjunto de dados.</p></li><li><p>Estruturas de navegação que funcionam como guias (único nível de navegação linear).</p></li><li><p>Reduzir ações de seleção, mantendo a navegação sempre disponível.</p></li><li><p>Fornecer espaço para exibir sempre a navegação.</p></li></ul>

## <a name="get-more-information-about-content-with-an-outlook-add-in"></a>Obter mais informações sobre o conteúdo com um suplemento do Outlook


 **Exemplo 1**

**Página de resultados e detalhes**
<br>
![Página de resultados e detalhes](../../images/office15-app-ux-fig11.jpg)
<br>
Mais adequado para:
<ul><li><p>Exponha a abrangência de seu conteúdo, se você tiver grandes conjuntos de dados que sejam úteis para exibição.</p></li><li><p>Exigir que você faça uma escolha ou seleção antes de mostrar mais detalhes.</p></li><li><p>Páginas de detalhes que usam o tamanho máximo do contêiner de suplemento.</p></li><li><p>Modelos de navegação que se beneficiam de um fluxo de "trocas".</p></li></ul>
 
 **Exemplo 2**

**Página de detalhes com conteúdo secundário**
<br>
![Página de detalhes com conteúdo secundário](../../images/office15-app-ux-fig12.jpg)
<br>
Mais adequado para:
<ul><li><p>Casos em que você deseja destacar uma parte do conteúdo.</p></li><li><p>Expor o conteúdo sem interação do usuário.</p></li><li><p>Navegação persistente (pode ser adicionada a esse modelo para fornecer uma mistura de simplicidade e facilidade de navegação).</p></li></ul>

## <a name="connect-to-an-online-service-and-present-data"></a>Conectar-se a um serviço online e apresentar dados


Esses exemplos mostram padrões de interação para a obtenção de dados e conteúdos de um serviço online. Eles podem ser usados em todos os três tipos de suplementos: suplementos de conteúdo, suplementos do painel de tarefas e suplementos do Outlook.

 **Exemplo 1**

**Carrossel**
<br>
![Carrossel](../../images/office15-app-ux-fig13.jpg)
<br>
Mais adequado para:
<ul><li><p>Pequenas quantidades de dados que podem ser expostos um de cada vez ou em grupos.</p></li><li><p>Expor o conteúdo em um formato de galeria, como apresentações de slides ou galerias de imagens.</p></li><li><p>Quando um modelo de navegação anterior/seguinte funciona bem.</p></li></ul>
 
 **Exemplo 2**

**Assistente**
<br>
![Assistente](../../images/office15-app-ux-fig14.jpg)
<br>
Mais adequado para:
<ul><li><p>Conteúdo que precisa ser apresentado em uma ordem específica.</p></li><li><p>Grandes quantidades de conteúdo que é consumido da melhor forma em uma série de pequenos itens.</p></li><li><p>Experiências de consumo semelhantes a livros.</p></li><li><p>Quando uma série de etapas ou ações são necessárias para concluir uma tarefa.</p></li></ul>
 
 **Exemplo 3**

**Formulário, resultados e detalhes**
<br>
![Formulário, resultados e detalhes](../../images/office15-app-ux-fig15.jpg)
<br>
Mais adequado para:
<ul><li><p>Suplementos que exigem a entrada de dados.</p></li><li><p>Um ponto de entrada para um padrão de resultados e detalhes.</p></li></ul>

## <a name="see-also"></a>Veja também

- [Diretrizes de design para suplementos do Office](../add-in-design.md)
    
