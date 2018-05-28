---
title: Criar e depurar Suplementos do Office no Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Criar e depurar Suplementos do Office no Visual Studio

Esse artigo descreve como usar o Visual Studio para criar o seu primeiro suplemento do Office. As etapas desse artigo t?m como base o Visual Studio 2015. Se voc? estiver usando outra vers?o do Visual Studio, os procedimentos poder?o variar um pouco.

> [!NOTE]
> Para come?ar a usar um suplemento do OneNote, confira o artigo [Crie seu primeiro suplemento do OneNote](../onenote/onenote-add-ins-getting-started.md).

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Criar um projeto de suplemento do Office no Visual Studio


Para come?ar, verifique se voc? tem as [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) instaladas e uma vers?o do Microsoft Office. ? poss?vel ingressar no [Programa do Desenvolvedor do Office 365](https://developer.microsoft.com/en-us/office/dev-program) ou seguir estas instru??es para receber a [?ltima vers?o](../develop/install-latest-office-version.md).


1. Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.
    
2. Na lista de tipos de projeto em **Visual C#** ou **Visual Basic**, expanda **Office/SharePoint**, escolha **Suplementos Web** e selecione um dos projetos de suplemento.  
    
3. Nomeie o projeto e escolha **OK** para cri?-lo.
    
4. O Visual Studio cria uma solu??o, e os dois projetos dele s?o exibidos no **Gerenciador de Solu??es**. A p?gina padr?o Home.html ? exibida no Visual Studio.
    
No Visual Studio 2015, alguns dos modelos de projetos de suplementos foram atualizados para refletir a funcionalidade adicional:


- Os suplementos de conte?do podem aparecer no corpo de documentos do Access e do PowerPoint, e em planilhas do Excel. Voc? tamb?m pode escolher a op??o Projeto B?sico para criar um projeto de suplemento de conte?do b?sico com c?digo inicial m?nimo, ou a op??o Projeto de Visualiza??o de Documento (apenas para Access e Excel) para criar um suplemento de conte?do mais completo que inclui c?digo inicial para visualizar e associar dados.
    
- Os suplementos do Outlook incluem op??es para incluir o suplemento em mensagens de email ou compromissos e para especificar se o suplemento est? dispon?vel quando uma mensagem de email ou um compromisso est? sendo redigido ou lido.
    

> [!NOTE]
> No Visual Studio, a maioria das op??es pode ser compreendida por meio das pr?prias descri??es, exceto a caixa de sele??o **Mensagem de Email**. Use essa caixa de sele??o se quiser criar um suplemento do Outlook exibido em itens de email e em solicita??es, respostas e cancelamentos de reuni?o.

Ao concluir o assistente, o Visual Studio cria uma solu??o que cont?m dois projetos.



|**Projeto**|**Descri??o**|
|:-----|:-----|
|Projeto de suplemento|Cont?m somente um arquivo de manifesto XML, que cont?m todas as configura??es que descrevem o suplemento. As configura??es ajudam o host do Office a determinar quando o suplemento dever? ser ativado e onde ele dever? aparecer. O Visual Studio gera o conte?do desse arquivo para que voc? possa executar o projeto e usar o suplemento imediatamente . Voc? pode alterar as configura??es a qualquer momento usando o editor de Manifesto.|
|Projeto de aplicativo Web|Cont?m as p?ginas de conte?do do suplemento, incluindo todos os arquivos e refer?ncias de arquivo de que voc? precisa para desenvolver p?ginas HTML e JavaScript com reconhecimento do Office. Enquanto voc? desenvolve o suplemento, o Visual Studio hospeda o aplicativo Web no servidor IIS local. Quando estiver pronto para publicar, voc? ter? de localizar um servidor para hospedar o projeto. Para saber mais sobre projetos de aplicativos Web ASP.NET, confira [Projetos Web ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## <a name="modify-your-add-in-settings"></a>Modificar as configura??es de suplemento


Para alterar as configura??es do seu suplemento, edite o arquivo de manifesto XML do projeto. No **Gerenciador de Solu??es**, expanda o n? de projeto do suplemento, expanda a pasta que cont?m o manifesto XML e escolha o manifesto XML. Voc? pode apontar para qualquer elemento do arquivo para exibir uma dica de ferramenta que descreve a finalidade do elemento. Para saber mais sobre o arquivo de manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


## <a name="develop-the-contents-of-your-add-in"></a>Desenvolver o conte?do do suplemento


Enquanto o projeto de suplemento permite modificar as configura??es que descrevem o suplemento, o aplicativo Web fornece o conte?do que aparece no suplemento. 

O projeto de aplicativo Web cont?m uma p?gina HTML padr?o e o arquivo JavaScript que voc? pode usar para come?ar. O projeto tamb?m cont?m um arquivo JavaScript que ? comum a todas as p?ginas que voc? adiciona ao projeto. Esses arquivos s?o convenientes porque cont?m refer?ncias a outras bibliotecas JavaScript, incluindo a API JavaScript para Office. 

? medida que o suplemento se tornar mais sofisticado, voc? poder? adicionar mais arquivos HTML e JavaScript. Voc? pode usar o conte?do dos arquivos HTML e JavaScript padr?o como exemplos dos tipos de refer?ncias que talvez queira adicionar a outras p?ginas no projeto para faz?-las funcionar com o suplemento. A tabela a seguir descreve os arquivos HTML e JavaScript padr?o.



|**Arquivo**|**Descri??o**|
|:-----|:-----|
|**Home.html**|Localizado na pasta **Home** do projeto, essa ? a p?gina HTML padr?o do suplemento. Essa p?gina ? exibida como a primeira no suplemento quando ele ? ativado em um documento, mensagem de email ou item de compromisso. Esse arquivo ? conveniente porque cont?m todas as refer?ncias de arquivo de que voc? precisa para come?ar. Quando estiver pronto para criar seu primeiro suplemento, basta adicionar o c?digo HTML a esse arquivo.|
|**Home.js**|Localizado na pasta **Home** do projeto, esse ? o arquivo JavaScript associado ? p?gina Home.js. Voc? pode colocar qualquer c?digo que seja espec?fico ao comportamento da p?gina Home.html no arquivo Home.js. O arquivo Home.js cont?m c?digo de exemplo para voc? come?ar.|
|**App.js**|Localizado na pasta **Add-in** do projeto, esse ? o arquivo JavaScript padr?o do suplemento inteiro. Voc? pode colocar c?digo comum ao comportamento de v?rias p?ginas do suplemento no arquivo App.js. O arquivo App.js cont?m c?digo de exemplo para voc? come?ar.|

> [!NOTE]
> N?o ? necess?rio usar esses arquivos. Fique ? vontade para adicionar outros arquivos ao projeto e us?-los. Se desejar que outro arquivo HTML apare?a como a p?gina inicial do suplemento, abra o editor de manifesto e aponte a propriedade **SourceLocation** para o nome do arquivo.


## <a name="debug-your-add-in"></a>Depurar o suplemento


Quando estiver pronto para iniciar o suplemento, examine as propriedades relacionadas ? compila??o e ? depura??o e inicie a solu??o.


### <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compila??o e depura??o

Antes de iniciar a solu??o, verifique se o Visual Studio abrir? o aplicativo host desejado. Essa informa??o ? exibida nas p?ginas de propriedades do projeto, com v?rias outras propriedades relacionadas ? compila??o e ? depura??o do suplemento.


### <a name="to-open-the-property-pages-of-a-project"></a>Para abrir as p?ginas de propriedades de um projeto


1. No **Gerenciador de Solu??es**, escolha o nome do projeto.
    
2. Na barra de menus, escolha **Exibir**, **Janela Propriedades**.
    
A tabela a seguir descreve as propriedades do projeto.



|**Propriedade**|**Descri??o**|
|:-----|:-----|
|**Iniciar A??o**|Especifica se o suplemento deve ser depurado em um cliente da ?rea de trabalho do Office ou em um cliente do Office Online no navegador especificado.|
|**Iniciar Documento** (apenas suplementos de conte?do e de painel de tarefas)|Especifica o documento a ser aberto quando voc? iniciar o projeto.|
|**Projeto da Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Endere?o de Email** (apenas suplementos do Outlook)|Especifica o endere?o de email da conta de usu?rio no Exchange Server ou no Exchange Online com a qual voc? deseja testar o suplemento do Outlook.|
|**Url EWS** (apenas suplementos do Outlook)|URL do servi?o Web do Exchange (por exemplo: https://www.contoso.com/ews/exchange.aspx). |
|**Url OWA** (apenas suplementos do Outlook)|URL do Outlook Web App (Por exemplo: https://www.contoso.com/owa).|
|**Nome de usu?rio** (apenas suplementos do Outlook)|Especifica o nome de sua conta de usu?rio no Exchange Server ou no Exchange Online.|
|**Arquivo do projeto**|Especifica o nome do arquivo que cont?m informa??es de compila??o, configura??o e outras informa??es sobre o projeto.|
|**Pasta do projeto**|O local do arquivo do projeto.|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>Use um documento existente para depurar o suplemento (apenas suplementos de conte?do e de painel de tarefas)


Voc? pode adicionar documentos ao projeto de suplemento. Se voc? tiver um documento que contenha os dados de teste que deseja usar com o suplemento, o Visual Studio abrir? esse documento quando voc? iniciar o projeto.


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>Para usar um documento existente para depurar o suplemento


1. No **Gerenciador de Solu??es**, escolha a pasta do projeto de suplemento.
    
    > [!NOTE]
    > Escolha o projeto do suplemento, n?o o projeto do aplicativo Web.

2. No menu **Projeto**, escolha **Adicionar Item Existente**.
    
3. Na caixa de di?logo **Adicionar Item Existente**, localize e selecione o documento que voc? deseja adicionar.
    
4. Escolha o bot?o **Adicionar** para adicionar o documento ao projeto.
    
5. No **Gerenciador de Solu??es**, abra o menu de atalho do projeto e escolha  **Propriedades**.
    
    As p?ginas de propriedades do projeto s?o exibidas.
    
6. Na lista **Iniciar Documento**, escolha o documento que voc? adicionou ao projeto e escolha o bot?o **OK** para fechar as p?ginas de propriedades.
    

### <a name="start-the-solution"></a>Iniciar a solu??o


O Visual Studio compilar? automaticamente a solu??o ao iniciar. Voc? pode iniciar a solu??o por meio da barra de **Menus** escolhendo **Depurar**, **Iniciar**. 


> [!NOTE]
> Se a depura??o de script n?o estiver habilitada no Internet Explorer, voc? n?o poder? iniciar o depurador no Visual Studio. ? poss?vel habilitar a depura??o de scripts abrindo a caixa de di?logo **Op??es da Internet**, escolhendo a guia **Avan?ado** e desmarcando as caixas de sele??o **Desabilitar depura??o de script (Internet Explorer)** e **Desabilitar a depura??o de script (outros)**.

O Visual Studio compila o projeto e faz o seguinte:


1. Cria uma c?pia do arquivo de manifesto XML e a adiciona ao diret?rio _NomedoProjeto_\Output. O aplicativo host consome esta c?pia quando voc? inicia o Visual Studio e depura o suplemento.
    
2. Cria um conjunto de entradas de registro no computador que permitem que o suplemento seja exibido no aplicativo host.
    
3. Compila o projeto de aplicativo da Web e o implanta no servidor Web IIS local (http://localhost). 
    
Depois, o Visual Studio faz o seguinte:


1. Modifica o elemento [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) do arquivo de manifesto XML, substituindo o token ~ remoteAppUrl pelo endere?o totalmente qualificado da p?gina inicial (por exemplo, http://localhost/MyAgave.html).
    
2. Inicia o projeto de aplicativo da Web no IIS Express.
    
3. Abre o aplicativo host. 
    
O Visual Studio n?o mostra erros de valida??o na janela **OUTPUT** ao compilar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** ? medida que eles ocorrem. O Visual Studio tamb?m relata erros de valida??o mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de c?digo e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no c?digo. Para saber mais, confira [Editor de c?digo e texto](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Para saber mais sobre como habilitar ou desabilitar a valida??o, confira: 

- [Op??es, editor de texto, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)
    
- [Tutorial: Definir op??es de valida??o para edi??o de HTML no Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, confira Valida??o, CSS, editor de texto, caixa de di?logo Op??es](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)
    
Para examinar as regras de valida??o do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a>Mostrar um suplemento no Excel, no Word ou no Project e percorrer o c?digo


Se voc? definir a propriedade **Start Document** do projeto de suplemento para o Excel ou o Word, o Visual Studio criar? um novo documento e o suplemento ser? exibido. Se voc? definir a propriedade **Start Document** do projeto de suplemento para usar um documento existente, o Visual Studio abrir? o documento, mas voc? precisar? inserir manualmente o suplemento. Se definir **Start Document** como **Microsoft Project**, voc? precisar? inserir manualmente o suplemento.


### <a name="to-show-an-office-add-in-in-excel-or-word"></a>Para mostrar um suplemento do Office no Excel ou no Word


1. No Excel ou no Word, na guia **Inserir**, escolha **Suplementos do Office**.
    
2. Na lista exibida, escolha o suplemento.
    

### <a name="to-show-an-office-add-in-in-project"></a>Para mostrar um suplemento do Office no Project


1. No Project, na guia **Projeto**, escolha **Suplementos do Office**.
    
2. Na lista exibida, escolha o suplemento.
    
No Visual Studio, voc? pode ent?o definir pontos de interrup??o. Depois, voc? interage com o suplemento e percorre o c?digo nos arquivos de c?digo HTML, JavaScript e C# ou VB.


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Mostrar o suplemento do Outlook no Outlook e percorrer o c?digo


Para exibir o suplemento no Outlook, abra uma mensagem de email ou um item de compromisso.

O Outlook ativa o suplemento para o item, contanto que os crit?rios de ativa??o sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um bot?o na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecer? um bot?o na faixa de op??es, na guia padr?o ou em uma guia personalizada especificada, e o suplemento n?o aparecer? na barra de suplementos.

Para exibir o suplemento do Outlook, escolha o bot?o do suplemento do Outlook.

No Visual Studio, voc? pode definir pontos de interrup??o. Depois, voc? interage com o suplemento do Outlook e percorre o c?digo nos arquivos de c?digo HTML, JavaScript e C# ou VB. 

Voc? tamb?m pode alterar o c?digo e examinar os efeitos das altera??es no suplemento do Outlook sem ter que fechar o Suplemento do Office e reiniciar o projeto. No Outlook, basta abrir o menu de atalho do suplemento do Outlook e escolher **Recarregar**.


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>Modificar o c?digo e continuar a depurar o suplemento sem ter que reiniciar o projeto


Voc? pode alterar o c?digo e examinar os efeitos das altera??es no suplemento sem ter que fechar o aplicativo host e reiniciar o projeto. Ap?s alterar o c?digo, abra o menu de atalho do suplemento e escolha **Recarregar**. Quando voc? recarregar o suplemento, ele ? desconectado do depurador do Visual Studio. Portanto, voc? pode exibir os efeitos da altera??o, mas n?o pode percorrer o c?digo novamente at? anexar o depurador do Visual Studio a todos os processos Iexplore.exe dispon?veis.


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a>Para anexar o depurador do Visual Studio a todos os processos Iexplore.exe dispon?veis


1. No Visual Studio, escolha **DEPURAR**, **Anexar ao Processo**.
    
2. Na caixa de di?logo **Anexar ao Processo**, escolha todos os processos **Iexplore.exe** dispon?veis e, em seguida, selecione o bot?o **Anexar**.
    

## <a name="next-steps"></a>Pr?ximas etapas

- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
    
