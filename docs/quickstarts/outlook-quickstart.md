---
title: Criar seu primeiro suplemento do Outlook
description: Saiba como criar um Suplemento do Outlook simples usando a API JS do Office.
ms.date: 08/24/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 6ce09b3b2f60cd4c77e966f6b920aa63caab299c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294329"
---
# <a name="build-your-first-outlook-add-in"></a>Criar seu primeiro suplemento do Outlook

Neste artigo, você acompanhará o processo de criação de um suplemento do painel de tarefas do Outlook que exibe pelo menos uma propriedade da mensagem selecionada.

## <a name="create-the-add-in"></a>Criar o suplemento

Você pode criar um suplemento do Office usando o [Gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) ou Visual Studio. O gerador Yeoman cria um projeto Node.js que pode ser gerenciado com o Visual Studio Code ou com qualquer outro editor, enquanto o Visual Studio cria uma solução do Visual Studio.  Selecione a guia do que você deseja usar e, em seguida, siga as instruções para criar o suplemento e testá-lo localmente.

# <a name="yeoman-generator"></a>[Gerador do Yeoman](#tab/yeomangenerator)

### <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

- A versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do [Yeoman gerador de suplementos do Office](https://github.com/OfficeDev/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Mesmo se você já instalou o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do npm.

### <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Escolha o tipo de projeto** - `Office Add-in Task Pane project`

    - **Escolha o tipo de script** - `Javascript`

    - **Qual será o nome do suplemento?** - `My Office Add-in`

    - **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** - `Outlook`

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-outlook.png)
    
    Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Navegue até a pasta raiz do projeto do aplicativo Web.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a>Explore o projeto

O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico. 

- O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.
- O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.
- O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.
- O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o Outlook.

### <a name="update-the-code"></a>Atualizar o código

1. No seu editor de código, abra o arquivo **./src/taskpane/taskpane.html** e substitua o elemento `<main>` inteiro (dentro do elemento `<body>`) com a seguinte marcação. A próxima marcação adiciona uma etiqueta onde o script no **./src/taskpane/taskpane.js** gravará os dados.

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código na função `run`. Este código usa a API JavaScript do Office para obter uma referência da mensagem atual e gravar o seu valor de propriedade `subject` no painel de tarefas.

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a>Experimente

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).

    ```command&nbsp;line
    npm start
    ```

    > [!IMPORTANT]
    > Se você vir um erro "Sideload não tem suporte", ignore-o e continue.

1. Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)para realizar o sideload do suplemento do Outlook.

1. No Outlook, escolha ou abra uma mensagem.

1. Escolha a guia **Página Inicial** (ou a guia **Mensagem**, se você abriu a mensagem em uma nova janela), e em seguida o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela de uma janela de mensagem do Outlook com o botão do suplemento realçado](../images/quick-start-button-1.png)

    > [!NOTE]
    > Se você receber a mensagem de erro "Não é possível abrir este suplemento do localhost" no painel de tarefas, siga as etapas descritas no [artigo de solução de problemas](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

1. Role para parte inferior do painel de tarefas e escolha o link **Executar** para escrever o assunto da mensagem no painel de tarefas.

    ![Uma captura de tela do painel de tarefas do suplemento com o link Executar realçado](../images/quick-start-task-pane-2.png)

    ![Uma captura de tela do painel de tarefas do suplemento exibindo o assunto da mensagem](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a>Próximas etapas

Parabéns, você criou o seu primeiro suplemento do painel de tarefas do Outlook! Em seguida, saiba mais sobre os recursos de um suplemento do Outlook e crie um suplemento mais complexo seguindo as etapas deste [tutorial de suplemento do Word](../tutorials/outlook-tutorial.md).

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio 2019](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada

    > [!NOTE]
    > Se você já instalou o Visual Studio 2019, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada.

- Office 365

    > [!NOTE]
    > Se você não tiver uma assinatura do Microsoft 365, poderá obter uma assinatura gratuita inscrevendo-se no [programa para desenvolvedores do Microsoft 365](https://developer.microsoft.com/office/dev-program).

### <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo** > **Projeto**.

1. Na lista de tipos de projeto sob **Visual C#** ou **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.

1. Dê um nome ao projeto e escolha **OK**.

1. O Visual Studio cria uma solução, e os dois projetos dele aparecem no **Gerenciador de Soluções**. O arquivo **MessageRead.html** é aberto no Visual Studio.

### <a name="explore-the-visual-studio-solution"></a>Explorar a solução do Visual Studio

Ao concluir o assistente, o Visual Studio cria uma solução que contém dois projetos.

|**Projeto**|**Descrição**|
|:-----|:-----|
|Projeto de suplemento|Contém somente um arquivo de manifesto XML, que contém todas as configurações que descrevem o suplemento. As configurações ajudam o aplicativo do Office a determinar quando o suplemento deverá ser ativado e onde ele deverá aparecer. O Visual Studio gera o conteúdo desse arquivo para que você possa executar o projeto e usar o suplemento imediatamente. Você pode alterar essas configurações a qualquer momento modificando o arquivo XML.|
|Projeto de aplicativo Web|Contém as páginas de conteúdo do suplemento, incluindo todos os arquivos e referências de arquivo de que você precisa para desenvolver páginas HTML e JavaScript com reconhecimento do Office. Enquanto você desenvolve o suplemento, o Visual Studio hospeda o aplicativo Web no servidor IIS local. Quando estiver pronto para publicar, você precisará implantar este projeto de aplicativo Web em um servidor Web.|

### <a name="update-the-code"></a>Atualizar o código

1. **MessageRead.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **MessageRead.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.
 
    ```HTML
    <body class="ms-font-m ms-welcome">
        <div class="ms-Fabric content-main">
            <h1 class="ms-font-xxl">Message properties</h1>
            <table class="ms-Table ms-Table--selectable">
                <thead>
                    <tr>
                        <th>Property</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>Id</strong></td>
                        <td class="prop-val"><code><label id="item-id"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Subject</strong></td>
                        <td class="prop-val"><code><label id="item-subject"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Message Id</strong></td>
                        <td class="prop-val"><code><label id="item-internetMessageId"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>From</strong></td>
                        <td class="prop-val"><code><label id="item-from"></label></code></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    ```

1. Abra o arquivo **MessageRead.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

    ```js
    'use strict';

    (function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                loadItemProps(Office.context.mailbox.item);
            });
        });

        function loadItemProps(item) {
            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        }
    })();
    ```

1. Abra o arquivo **MessageRead.css** na raiz do projeto do aplicativo Web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

    ```CSS
    html,
    body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    td.prop-val {
        word-break: break-all;
    }

    .content-main {
        margin: 10px;
    }
    ```

### <a name="update-the-manifest"></a>Atualizar o manifesto

1. Abra o arquivo de manifesto XML do projeto do Suplemento. Este arquivo define as configurações e os recursos do suplemento.

1. O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.

1. O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o por `My Office Add-in`.

1. O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o por `My First Outlook add-in`.

1. Salve o arquivo.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a>Experimente

1. Usando o Visual Studio, teste o suplemento do Outlook recém-criado ao pressionar F5 ou o botão **Iniciar**. O suplemento será hospedado localmente no IIS.

1. Na caixa de diálogo**Conectar-se à conta de email do Exchange**, digite o endereço de email e senha da sua [conta da Microsoft](https://account.microsoft.com/account) e, em seguida, escolha **Conectar**. Quando a página de login do Outlook.com for aberta em um navegador, entre em sua conta de email com as mesmas credenciais que você inseriu anteriormente.

    > [!NOTE]
    > Se a caixa de diálogo **Conectar a uma conta de email do Exchange** repetidamente solicitar que você entre, a Autenticação Básica poderá ser desabilitada para contas no locatário do Microsoft 365. Para testar esse suplemento, entre usando uma [Conta da Microsoft](https://account.microsoft.com/account).

1. No Outlook na Web, escolha ou abra uma mensagem.

1. Dentro da mensagem, localize as reticências do menu de estouro que contém o botão do suplemento.

    ![Uma captura de tela de uma janela de mensagem do Outlook na Web com as reticências realçadas](../images/quick-start-button-owa-1.png)

1. No menu estouro, localize o botão do suplemento.

    ![Uma captura de tela de uma janela de mensagem do Outlook na Web com o botão do suplemento realçado](../images/quick-start-button-owa-2.png)

1. Clique no botão para abrir o painel de tarefas do suplemento.

    ![Uma captura de tela do painel de tarefas do suplemento no Outlook na Web exibindo as propriedades da mensagem](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > Se o painel de tarefas não carregar, tente verificar abrindo-o em um navegador no mesmo computador.

### <a name="next-steps"></a>Próximas etapas

Parabéns, você criou o seu primeiro suplemento do painel de tarefas do Outlook! Em seguida, saiba mais sobre como [desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).

---
