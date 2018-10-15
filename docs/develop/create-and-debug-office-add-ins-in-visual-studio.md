---
title: Criar e depurar Suplementos do Office no Visual Studio
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 0bbc1b167924ce4b7a1310f04b41751173312cd6
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506123"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Criar e depurar Suplementos do Office no Visual Studio

Este artigo descreve como usar o Visual Studio para criar o seu primeiro Suplemento do Office. As etapas deste artigo têm como base o Visual Studio 2015. Se você estiver usando outra versão do Visual Studio, os procedimentos poderão variar um pouco.

> [!NOTE]
> Para começar a usar um suplemento do OneNote, confira o artigo [Criar seu primeiro suplemento do OneNote](../onenote/onenote-add-ins-getting-started.md).

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Criar um projeto de Suplemento do Office no Visual Studio


Para começar, verifique se você tem as [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) instaladas e uma versão do Microsoft Office. É possível ingressar no [Programa do Desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program) ou seguir estas instruções para receber a [última versão](../develop/install-latest-office-version.md).

1. Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo** > **Projeto**.
2. Na lista de tipos de projeto, em **Visual C#** ou **Visual Basic**, expanda **Office/SharePoint**, escolha **Suplementos** e escolha um dos projetos de suplemento.
3. Nomeie o projeto e escolha **OK** para criá-lo.

No Visual Studio 2017, os seguintes modelos de projeto de suplementos têm opções adicionais depois de escolher **OK**:

**PowerPoint**
- Você pode escolher **Adicionar novas funcionalidades no PowerPoint** para criar um suplemento do painel de tarefas.
- Ou você pode escolher **Inserir conteúdo nos slides do PowerPoint** para criar um suplemento de conteúdo.

**Excel** 
- Você pode escolher **Adicionar novas funcionalidades no Excel** para criar um suplemento do painel de tarefas.
- Ou você pode escolher **Inserir conteúdo na planilha do Excel** para criar um suplemento de conteúdo.
    - Se você criar um suplemento de conteúdo, terá uma escolha adicional de **Suplemento básico** que cria um projeto de suplemento de conteúdo com código inicial mínimo.
    - Ou você pode escolher um **Suplemento de visualização de documento** que inclui o código inicial para visualizar e vincular dados.

Após concluir o assistente, o Visual Studio cria uma solução para você que contém dois projetos. Você verá a página Home.html padrão abrir.

|**Projeto**|**Descrição**|
|:-----|:-----|
|Projeto de suplemento|Contém somente um arquivo de manifesto XML, que contém todas as configurações que descrevem o suplemento. As configurações ajudam o host do Office a determinar quando o suplemento deverá ser ativado e onde ele deverá aparecer. O Visual Studio gera o conteúdo desse arquivo para que você possa executar o projeto e usar o suplemento imediatamente. Você pode alterar as configurações a qualquer momento usando o editor de Manifesto.|
|Projeto de aplicativo da Web|Contém as páginas de conteúdo do suplemento, incluindo todos os arquivos e referências de arquivo de que você precisa para desenvolver páginas HTML e JavaScript com reconhecimento do Office. Enquanto você desenvolve o suplemento, o Visual Studio hospeda o aplicativo da Web no servidor IIS local. Quando você estiver pronto para publicar, deverá localizar um servidor para hospedar esse projeto. Para saber mais sobre projetos de aplicativo Web ASP .NET, consulte [Projetos Web ASP.NET](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## <a name="modify-your-add-in-settings"></a>Modificar as configurações de suplemento


Para alterar as configurações do seu suplemento, edite o arquivo de manifesto XML do projeto. No **Gerenciador de soluções**, expanda o nó de projeto do suplemento, expanda a pasta que contém o manifesto XML e escolha o manifesto XML. Você pode apontar para qualquer elemento do arquivo para exibir uma dica de ferramenta que descreve a finalidade do elemento. Para saber mais sobre o arquivo de manifesto, consulte [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).


## <a name="develop-the-contents-of-your-add-in"></a>Desenvolver o conteúdo do seu suplemento

Enquanto o projeto de suplemento permite modificar as configurações que descrevem o suplemento, o aplicativo da Web fornece o conteúdo que aparece no suplemento. 

O projeto de aplicativo da Web contém uma página HTML padrão e o arquivo de JavaScript que você pode usar para começar. Esses arquivos contêm referências a outras bibliotecas de JavaScript, incluindo a API JavaScript para Office. Para desenvolver seu suplemento, atualize esses arquivos e adicione mais arquivos HTML e JavaScript. A tabela a seguir descreve os arquivos JavaScript e HTML padrão.

> [!NOTE]
> Os arquivos na tabela a seguir podem estar na pasta raiz do projeto Web, ou na pasta **Inicial**, dependendo do tipo de modelo de projeto que você usou.

|**Arquivo**|**Descrição**|
|:-----|:-----|
|**Home.html**|A página HTML padrão do suplemento. Essa página aparece como a primeira página dentro do suplemento quando é ativada em um documento, uma mensagem de email ou um item de compromisso. Esse arquivo contém todas as referências de arquivo de que você precisa para começar. Para começar a desenvolver seu suplemento, adicione seu código HTML nesse arquivo.|
|**Home.js**|O arquivo JavaScript associado à página Home.html. Você pode colocar qualquer código que é específico para o comportamento da página Home.html no arquivo Home.js. O arquivo Home.js contém alguns códigos de exemplo para você começar.|
|**Home.css**|Define os estilos padrão a serem aplicados ao seu suplemento. É recomendável usar o Office UI Fabric para design e estilos. Para obter mais informações, consulte [Office UI Fabric nos Suplementos do Office](../design/office-ui-fabric.md).|

> [!NOTE]
> Você não precisa usar esses arquivos. Você pode adicionar e usar outros arquivos no projeto. Se quiser que outro arquivo HTML apareça como a página inicial do suplemento, abra o editor de manifesto e, em seguida, defina a propriedade **SourceLocation** no nome do arquivo.

## <a name="debug-your-add-in"></a>Depurar o suplemento

O Visual Studio oferece propriedades de compilação e depuração para ajudar na depuração do seu suplemento.

### <a name="review-the-build-and-debug-properties"></a>Examinar as propriedades de compilação e depuração

Antes de iniciar a solução, verifique se o Visual Studio abrirá o aplicativo host desejado. Essa informação é exibida nas páginas de propriedades do projeto, com várias outras propriedades relacionadas à compilação e à depuração do suplemento.

### <a name="to-open-the-property-pages-of-a-project"></a>Para abrir as páginas de propriedade de um projeto

1. No **Gerenciador de soluções**, escolha o projeto de suplemento básico (não o projeto Web).    
2. Na barra de menus, escolha **Visualizar** >  **Janela Propriedades**.
    
A tabela a seguir descreve as propriedades do projeto.



|**Propriedade**|**Descrição**|
|:-----|:-----|
|**Iniciar ação**|Especifica se o suplemento deve ser depurado em um cliente da área de trabalho do Office ou em um cliente do Office Online no navegador especificado.|
|**Iniciar documento** (apenas suplementos de conteúdo e de painel de tarefas)|Especifica o documento a ser aberto quando você inicia o projeto.|
|**Projeto Web**|Especifica o nome do projeto Web associado ao suplemento.|
|**Endereço de email** (apenas suplementos do Outlook)|Especifica o endereço de email da conta de usuário no Exchange Server ou no Exchange Online com a qual você deseja testar o suplemento do Outlook.|
|**Url EWS** (apenas suplementos do Outlook)|URL do serviço Web do Exchange (por exemplo: https://www.contoso.com/ews/exchange.aspx). |
|**Url OWA** (apenas suplementos do Outlook)|URL do aplicativo Web do Outlook (por exemplo: https://www.contoso.com/owa).|
|**Nome de usuário** (apenas suplementos do Outlook)|Especifica o nome de sua conta de usuário no Exchange Server ou no Exchange Online.|
|**Arquivo de projeto**|Especifica o nome do arquivo que contém informações de compilação, configuração e outras informações sobre o projeto.|
|**Pasta do projeto**|A localização do arquivo de projeto.|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>Use um documento existente para depurar o suplemento (apenas suplementos de conteúdo e de painel de tarefas)

Você pode adicionar documentos ao projeto de suplemento. Se você tiver um documento que contenha os dados de teste que deseja usar com o suplemento, o Visual Studio abrirá esse documento quando você iniciar o projeto.

### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>Para usar um documento existente para depurar o suplemento

1. No **Gerenciador de soluções**, escolha a pasta do projeto de suplemento.
    
    > [!NOTE]
    > Escolha o projeto do suplemento, não o projeto do aplicativo da Web.

2. No menu **Projeto**, escolha **Adicionar item existente**.
    
3. Na caixa de diálogo **Adicionar item existente**, localize e selecione o documento que você deseja adicionar.
    
4. Escolha o botão **Adicionar** para adicionar o documento ao projeto.
    
5. No **Gerenciador de soluções**, escolha a pasta do projeto de suplemento.
6. Na barra de menus, escolha **Visualizar** > **Janela Propriedades**.
7. Na janela Propriedades, escolha a lista **Iniciar documento** e, em seguida, escolha o documento que você adicionou ao projeto. Agora o projeto está configurado para iniciar seu suplemento no seu documento existente.

### <a name="start-the-solution"></a>Iniciar solução

Inicia a solução na barra de menus, escolhendo **Depurar** > **Iniciar depuração**. O Visual Studio vai automaticamente compilar a solução e iniciar o Office para hospedar o suplemento.

Quando o projeto do Visual Studio compilar o projeto, ele vai executar as seguintes tarefas:

1. Cria uma cópia do arquivo de manifesto XML e a adiciona ao diretório _NomedoProjeto_\Output. O aplicativo host consome esta cópia quando você inicia o Visual Studio e depura o suplemento.
    
2. Cria um conjunto de entradas de registro no seu computador que permitem que o suplemento apareça no aplicativo host.
    
3. Compila o projeto de aplicativo da Web e o implanta no servidor Web IIS local (http://localhost). 
    
Depois, o Visual Studio faz o seguinte:

1. Modifica o elemento [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js) do arquivo de manifesto XML, substituindo o token ~remoteAppUrl pelo endereço totalmente qualificado da página inicial (por exemplo, http://localhost/MyAgave.html).
    
2. Inicia o projeto de aplicativo Web no IIS Express.
    
3. Abre o aplicativo host. 
    
O Visual Studio não mostra erros de validação na janela **OUTPUT** ao compilar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** à medida que eles ocorrem. O Visual Studio também relata erros de validação mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de código e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no código. Para saber mais, confira [Editor de código e texto](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). Para saber mais sobre como habilitar ou desabilitar a validação, confira: 

- [Opções, Editor de texto, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [Tutorial: Definir opções de validação para edição de HTML no Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, confira Validação, CSS, editor de texto, caixa de diálogo Opções](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
Para examinar as regras de validação do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a>Veja um suplemento do Excel ou Word e percorra o código

Se você definir a propriedade **Iniciar documento** do projeto de suplemento para o Excel ou Word, o Visual Studio cria um novo documento e o suplemento aparece. Se você definir a propriedade **Iniciar documento** do projeto de suplemento para usar um documento existente, o Visual Studio abre o documento, mas você precisará inserir o suplemento manualmente.

1. No Excel ou Word, na guia **Inserir**, escolha a lista suspensa **Meus suplementos**. Escolha a lista na seta suspensa e não no botão que abre a caixa de diálogo **Suplementos do Office**.
2. Em **Suplementos do desenvolvedor**, escolha seu suplemento.

No Visual Studio, você poderá definir pontos de interrupção, interagir com seu suplemento e percorrer o código nos seus arquivos HTML ou JavaScript.

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Mostrar o suplemento do Outlook no Outlook e percorrer o código

Para exibir o suplemento no Outlook, abra uma mensagem de email ou um item de compromisso.

O Outlook ativa o suplemento para o item, contanto que os critérios de ativação sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um botão na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecerá um botão na faixa de opções, na guia padrão ou em uma guia personalizada especificada, e o suplemento não aparecerá na barra de suplementos.

Para exibir o suplemento do Outlook, escolha o botão do suplemento do Outlook.

No Visual Studio, você poderá definir pontos de interrupção, interagir com seu suplemento e percorrer o código nos seus arquivos HTML ou JavaScript.

Você também pode alterar o código e examinar os efeitos das alterações no suplemento do Outlook sem ter que fechar o Suplemento do Office e reiniciar o projeto. No Outlook, basta abrir o menu de atalho do suplemento do Outlook e escolher **Recarregar**.


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>Modificar o código e continuar a depurar o suplemento sem ter que reiniciar o projeto

Você pode alterar o seu código e revisar os efeitos dessas alterações em seu suplemento sem ter que fechar o aplicativo host e iniciar o projeto novamente. Depois de alterar e salvar seu código, abra o menu de atalho para o suplemento e, em seguida, escolha **Recarregar**.
    

## <a name="next-steps"></a>Próximas etapas

- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
    
