---
title: Criar e depurar Suplementos do Office no Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437602"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Criar e depurar Suplementos do Office no Visual Studio

Esse artigo descreve como usar o Visual Studio para criar o seu primeiro suplemento do Office. As etapas desse artigo têm como base o Visual Studio 2015. Se você estiver usando outra versão do Visual Studio, os procedimentos poderão variar um pouco.

> [!NOTE]
> Para começar a usar um suplemento do OneNote, confira o artigo [Crie seu primeiro suplemento do OneNote](../onenote/onenote-add-ins-getting-started.md).

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Criar um projeto de suplemento do Office no Visual Studio


Para começar, verifique se você tem as [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) instaladas e uma versão do Microsoft Office. É possível ingressar no [Programa do Desenvolvedor do Office 365](https://developer.microsoft.com/en-us/office/dev-program) ou seguir estas instruções para receber a [última versão](../develop/install-latest-office-version.md).


1. Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.
    
2. Na lista de tipos de projeto em **Visual C#** ou **Visual Basic**, expanda **Office/SharePoint**, escolha **Suplementos Web** e selecione um dos projetos de suplemento.  
    
3. Nomeie o projeto e escolha **OK** para criá-lo.
    
4. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. A página padrão Home.html é exibida no Visual Studio.
    
No Visual Studio 2015, alguns dos modelos de projetos de suplementos foram atualizados para refletir a funcionalidade adicional:


- Os suplementos de conteúdo podem aparecer no corpo de documentos do Access e do PowerPoint, e em planilhas do Excel. Você também pode escolher a opção Projeto Básico para criar um projeto de suplemento de conteúdo básico com código inicial mínimo, ou a opção Projeto de Visualização de Documento (apenas para Access e Excel) para criar um suplemento de conteúdo mais completo que inclui código inicial para visualizar e associar dados.
    
- Os suplementos do Outlook incluem opções para incluir o suplemento em mensagens de email ou compromissos e para especificar se o suplemento está disponível quando uma mensagem de email ou um compromisso está sendo redigido ou lido.
    

> [!NOTE]
> No Visual Studio, a maioria das opções pode ser compreendida por meio das próprias descrições, exceto a caixa de seleção **Mensagem de Email**. Use essa caixa de seleção se quiser criar um suplemento do Outlook exibido em itens de email e em solicitações, respostas e cancelamentos de reunião.

Ao concluir o assistente, o Visual Studio cria uma solução que contém dois projetos.



|**Projeto**|**Descrição**|
|:-----|:-----|
|Projeto de suplemento|Contém somente um arquivo de manifesto XML, que contém todas as configurações que descrevem o suplemento. As configurações ajudam o host do Office a determinar quando o suplemento deverá ser ativado e onde ele deverá aparecer. O Visual Studio gera o conteúdo desse arquivo para que você possa executar o projeto e usar o suplemento imediatamente . Você pode alterar as configurações a qualquer momento usando o editor de Manifesto.|
|Projeto de aplicativo Web|Contém as páginas de conteúdo do suplemento, incluindo todos os arquivos e referências de arquivo de que você precisa para desenvolver páginas HTML e JavaScript com reconhecimento do Office. Enquanto você desenvolve o suplemento, o Visual Studio hospeda o aplicativo Web no servidor IIS local. Quando estiver pronto para publicar, você terá de localizar um servidor para hospedar o projeto. Para saber mais sobre projetos de aplicativos Web ASP.NET, confira [Projetos Web ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## <a name="modify-your-add-in-settings"></a>Modificar as configurações de suplemento


Para alterar as configurações do seu suplemento, edite o arquivo de manifesto XML do projeto. No **Gerenciador de Soluções**, expanda o nó de projeto do suplemento, expanda a pasta que contém o manifesto XML e escolha o manifesto XML. Você pode apontar para qualquer elemento do arquivo para exibir uma dica de ferramenta que descreve a finalidade do elemento. Para saber mais sobre o arquivo de manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


## <a name="develop-the-contents-of-your-add-in"></a>Desenvolver o conteúdo do suplemento


Enquanto o projeto de suplemento permite modificar as configurações que descrevem o suplemento, o aplicativo Web fornece o conteúdo que aparece no suplemento. 

O projeto de aplicativo Web contém uma página HTML padrão e o arquivo JavaScript que você pode usar para começar. O projeto também contém um arquivo JavaScript que é comum a todas as páginas que você adiciona ao projeto. Esses arquivos são convenientes porque contêm referências a outras bibliotecas JavaScript, incluindo a API JavaScript para Office. 

À medida que o suplemento se tornar mais sofisticado, você poderá adicionar mais arquivos HTML e JavaScript. Você pode usar o conteúdo dos arquivos HTML e JavaScript padrão como exemplos dos tipos de referências que talvez queira adicionar a outras páginas no projeto para fazê-las funcionar com o suplemento. A tabela a seguir descreve os arquivos HTML e JavaScript padrão.



|**Arquivo**|**Descrição**|
|:-----|:-----|
|**Home.html**|Localizado na pasta **Home** do projeto, essa é a página HTML padrão do suplemento. Essa página é exibida como a primeira no suplemento quando ele é ativado em um documento, mensagem de email ou item de compromisso. Esse arquivo é conveniente porque contém todas as referências de arquivo de que você precisa para começar. Quando estiver pronto para criar seu primeiro suplemento, basta adicionar o código HTML a esse arquivo.|
|**Home.js**|Localizado na pasta **Home** do projeto, esse é o arquivo JavaScript associado à página Home.js. Você pode colocar qualquer código que seja específico ao comportamento da página Home.html no arquivo Home.js. O arquivo Home.js contém código de exemplo para você começar.|
|**App.js**|Localizado na pasta **Add-in** do projeto, esse é o arquivo JavaScript padrão do suplemento inteiro. Você pode colocar código comum ao comportamento de várias páginas do suplemento no arquivo App.js. O arquivo App.js contém código de exemplo para você começar.|

> [!NOTE]
> Não é necessário usar esses arquivos. Fique à vontade para adicionar outros arquivos ao projeto e usá-los. Se desejar que outro arquivo HTML apareça como a página inicial do suplemento, abra o editor de manifesto e aponte a propriedade **SourceLocation** para o nome do arquivo.


## <a name="debug-your-add-in"></a>Depurar o suplemento


Quando estiver pronto para iniciar o suplemento, examine as propriedades relacionadas à compilação e à depuração e inicie a solução.


### <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compilação e depuração

Antes de iniciar a solução, verifique se o Visual Studio abrirá o aplicativo host desejado. Essa informação é exibida nas páginas de propriedades do projeto, com várias outras propriedades relacionadas à compilação e à depuração do suplemento.


### <a name="to-open-the-property-pages-of-a-project"></a>Para abrir as páginas de propriedades de um projeto


1. No **Gerenciador de Soluções**, escolha o nome do projeto.
    
2. Na barra de menus, escolha **Exibir**, **Janela Propriedades**.
    
A tabela a seguir descreve as propriedades do projeto.



|**Propriedade**|**Descrição**|
|:-----|:-----|
|**Iniciar Ação**|Especifica se o suplemento deve ser depurado em um cliente da área de trabalho do Office ou em um cliente do Office Online no navegador especificado.|
|**Iniciar Documento** (apenas suplementos de conteúdo e de painel de tarefas)|Especifica o documento a ser aberto quando você iniciar o projeto.|
|**Projeto da Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Endereço de Email** (apenas suplementos do Outlook)|Especifica o endereço de email da conta de usuário no Exchange Server ou no Exchange Online com a qual você deseja testar o suplemento do Outlook.|
|**Url EWS** (apenas suplementos do Outlook)|URL do serviço Web do Exchange (por exemplo: https://www.contoso.com/ews/exchange.aspx). |
|**Url OWA** (apenas suplementos do Outlook)|URL do Outlook Web App (Por exemplo: https://www.contoso.com/owa).|
|**Nome de usuário** (apenas suplementos do Outlook)|Especifica o nome de sua conta de usuário no Exchange Server ou no Exchange Online.|
|**Arquivo do projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|O local do arquivo do projeto.|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>Use um documento existente para depurar o suplemento (apenas suplementos de conteúdo e de painel de tarefas)


Você pode adicionar documentos ao projeto de suplemento. Se você tiver um documento que contenha os dados de teste que deseja usar com o suplemento, o Visual Studio abrirá esse documento quando você iniciar o projeto.


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>Para usar um documento existente para depurar o suplemento


1. No **Gerenciador de Soluções**, escolha a pasta do projeto de suplemento.
    
    > [!NOTE]
    > Escolha o projeto do suplemento, não o projeto do aplicativo Web.

2. No menu **Projeto**, escolha **Adicionar Item Existente**.
    
3. Na caixa de diálogo **Adicionar Item Existente**, localize e selecione o documento que você deseja adicionar.
    
4. Escolha o botão **Adicionar** para adicionar o documento ao projeto.
    
5. No **Gerenciador de Soluções**, abra o menu de atalho do projeto e escolha  **Propriedades**.
    
    As páginas de propriedades do projeto são exibidas.
    
6. Na lista **Iniciar Documento**, escolha o documento que você adicionou ao projeto e escolha o botão **OK** para fechar as páginas de propriedades.
    

### <a name="start-the-solution"></a>Iniciar a solução


O Visual Studio compilará automaticamente a solução ao iniciar. Você pode iniciar a solução por meio da barra de **Menus** escolhendo **Depurar**, **Iniciar**. 


> [!NOTE]
> Se a depuração de script não estiver habilitada no Internet Explorer, você não poderá iniciar o depurador no Visual Studio. É possível habilitar a depuração de scripts abrindo a caixa de diálogo **Opções da Internet**, escolhendo a guia **Avançado** e desmarcando as caixas de seleção **Desabilitar depuração de script (Internet Explorer)** e **Desabilitar a depuração de script (outros)**.

O Visual Studio compila o projeto e faz o seguinte:


1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao diretório _NomedoProjeto_\Output. O aplicativo host consome esta cópia quando você inicia o Visual Studio e depura o suplemento.
    
2. Cria um conjunto de entradas de registro no computador que permitem que o suplemento seja exibido no aplicativo host.
    
3. Compila o projeto de aplicativo da Web e o implanta no servidor Web IIS local (http://localhost). 
    
Depois, o Visual Studio faz o seguinte:


1. Modifica o elemento [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) do arquivo de manifesto XML, substituindo o token ~ remoteAppUrl pelo endereço totalmente qualificado da página inicial (por exemplo, http://localhost/MyAgave.html).
    
2. Inicia o projeto de aplicativo da Web no IIS Express.
    
3. Abre o aplicativo host. 
    
O Visual Studio não mostra erros de validação na janela **OUTPUT** ao compilar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** à medida que eles ocorrem. O Visual Studio também relata erros de validação mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de código e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no código. Para saber mais, confira [Editor de código e texto](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Para saber mais sobre como habilitar ou desabilitar a validação, confira: 

- [Opções, editor de texto, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)
    
- [Tutorial: Definir opções de validação para edição de HTML no Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, confira Validação, CSS, editor de texto, caixa de diálogo Opções](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)
    
Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a>Mostrar um suplemento no Excel, no Word ou no Project e percorrer o código


Se você definir a propriedade **Start Document** do projeto de suplemento para o Excel ou o Word, o Visual Studio criará um novo documento e o suplemento será exibido. Se você definir a propriedade **Start Document** do projeto de suplemento para usar um documento existente, o Visual Studio abrirá o documento, mas você precisará inserir manualmente o suplemento. Se definir **Start Document** como **Microsoft Project**, você precisará inserir manualmente o suplemento.


### <a name="to-show-an-office-add-in-in-excel-or-word"></a>Para mostrar um suplemento do Office no Excel ou no Word


1. No Excel ou no Word, na guia **Inserir**, escolha **Suplementos do Office**.
    
2. Na lista exibida, escolha o suplemento.
    

### <a name="to-show-an-office-add-in-in-project"></a>Para mostrar um suplemento do Office no Project


1. No Project, na guia **Projeto**, escolha **Suplementos do Office**.
    
2. Na lista exibida, escolha o suplemento.
    
No Visual Studio, você pode então definir pontos de interrupção. Depois, você interage com o suplemento e percorre o código nos arquivos de código HTML, JavaScript e C# ou VB.


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Mostrar o suplemento do Outlook no Outlook e percorrer o código


Para exibir o suplemento no Outlook, abra uma mensagem de email ou um item de compromisso.

O Outlook ativa o suplemento para o item, contanto que os critérios de ativação sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um botão na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecerá um botão na faixa de opções, na guia padrão ou em uma guia personalizada especificada, e o suplemento não aparecerá na barra de suplementos.

Para exibir o suplemento do Outlook, escolha o botão do suplemento do Outlook.

No Visual Studio, você pode definir pontos de interrupção. Depois, você interage com o suplemento do Outlook e percorre o código nos arquivos de código HTML, JavaScript e C# ou VB. 

Você também pode alterar o código e examinar os efeitos das alterações no suplemento do Outlook sem ter que fechar o Suplemento do Office e reiniciar o projeto. No Outlook, basta abrir o menu de atalho do suplemento do Outlook e escolher **Recarregar**.


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>Modificar o código e continuar a depurar o suplemento sem ter que reiniciar o projeto


Você pode alterar o código e examinar os efeitos das alterações no suplemento sem ter que fechar o aplicativo host e reiniciar o projeto. Após alterar o código, abra o menu de atalho do suplemento e escolha **Recarregar**. Quando você recarregar o suplemento, ele é desconectado do depurador do Visual Studio. Portanto, você pode exibir os efeitos da alteração, mas não pode percorrer o código novamente até anexar o depurador do Visual Studio a todos os processos Iexplore.exe disponíveis.


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a>Para anexar o depurador do Visual Studio a todos os processos Iexplore.exe disponíveis


1. No Visual Studio, escolha **DEPURAR**, **Anexar ao Processo**.
    
2. Na caixa de diálogo **Anexar ao Processo**, escolha todos os processos **Iexplore.exe** disponíveis e, em seguida, selecione o botão **Anexar**.
    

## <a name="next-steps"></a>Próximas etapas

- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
    
