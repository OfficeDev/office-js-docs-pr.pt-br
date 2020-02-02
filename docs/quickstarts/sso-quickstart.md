---
title: Use o gerador Yeoman para criar um Suplemento do Office que use SSO (prévia)
description: Use o gerador Yeoman para criar um Suplemento do Office com Node.js que use logon único (prévia).
ms.date: 01/30/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 263a84a9084f7f75beb13b4336b61027de0bf907
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650023"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="cc1dc-103">Use o gerador Yeoman para criar um Suplemento do Office que use logon único (prévia)</span><span class="sxs-lookup"><span data-stu-id="cc1dc-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="cc1dc-104">Neste artigo, você seguirá pelo processo de uso do gerador Yeoman para criar um Suplemento do Office para Excel, Outlook, Word ou PowerPoint que usa o logon único (SSO) sempre que possível, e usa um método alternativo de autenticação do usuário quando não há suporte ao SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="cc1dc-105">Antes de tentar concluir o início rápido, revise [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md) para aprender conceitos básicos sobre o SSO em Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="cc1dc-106">O gerador Yeoman simplifica o processo de criação de um suplemento de SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="cc1dc-107">Para um passo a passo detalhado descrevendo como concluir manualmente as etapas que o gerador Yeoman automatiza, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cc1dc-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="cc1dc-108">Prerequisites</span></span>

* <span data-ttu-id="cc1dc-109">[Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="cc1dc-109">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="cc1dc-110">A versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do [Yeoman gerador de suplementos do Office](https://github.com/OfficeDev/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="cc1dc-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="cc1dc-111">Se você estiver usando um Mac e não tiver a CLI do Azure instalada no computador, instale o [Homebrew](https://brew.sh/).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-111">If you're using a Mac and don't have the Azure CLI installed on your machine, you must install [Homebrew](https://brew.sh/).</span></span> <span data-ttu-id="cc1dc-112">O script de configuração do SSO executado durante o início rápido usará o Homebrew para instalar a CLI do Azure e, em seguida, usará a CLI do Azure para configurar o SSO no Azure.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-112">The SSO configuration script that you'll run during this quick start will use Homebrew to install the Azure CLI, and will then use the Azure CLI to configure SSO within Azure.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="cc1dc-113">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="cc1dc-113">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="cc1dc-114">O gerador Yeoman pode criar um Suplemento do Office habilitado para SSO do Excel, Outlook, Word ou PowerPoint e pode ser criado com o tipo de script JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-114">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="cc1dc-115">As instruções a seguir especificam o `JavaScript` e o `Excel`, mas você deverá escolher o tipo de script e o aplicativo cliente do Office que atendem melhor ao seu cenário.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-115">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="cc1dc-116">**Escolha o tipo de projeto:** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="cc1dc-116">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="cc1dc-117">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="cc1dc-117">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="cc1dc-118">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="cc1dc-118">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="cc1dc-119">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="cc1dc-119">**Which Office client application would you like to support?**</span></span> `Excel`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="cc1dc-121">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-121">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="cc1dc-122">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="cc1dc-122">Explore the project</span></span>

<span data-ttu-id="cc1dc-123">O projeto de suplemento que você criou com o gerador do Yeoman contém um código para um suplemento de painel de tarefas habilitado para SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-123">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

- <span data-ttu-id="cc1dc-124">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="cc1dc-125">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="cc1dc-126">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="cc1dc-127">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="cc1dc-128">O arquivo **./src/helpers/documentHelper.js** usa a biblioteca JavaScript do Office para adicionar os dados do Microsoft Graph ao documento do Office.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-128">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="cc1dc-129">O arquivo **./src/helpers/fallbackauthdialog.html** é a página sem interface do usuário que carrega o JavaScript do método de autenticação de fallback.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-129">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="cc1dc-130">O arquivo **./src/Helpers/fallbackauthdialog.js** contém o JavaScript do método de autenticação fallback que entra no usuário com o msal.js.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-130">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="cc1dc-131">O arquivo **./src/helpers/fallbackauthhelper.js** contém o painel de tarefas JavaScript que chama o método de autenticação de fallback em cenários em que não há suporte à autenticação SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-131">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="cc1dc-132">O arquivo **./src/helpers/ssoauthhelper.js** contém a chamada JavaScript à API de SSO, `getAccessToken`, recebe o token de inicialização, inicia a troca do token de inicialização por um token de acesso ao Microsoft Graph e chama o Microsoft Graph para obter os dados.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-132">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="cc1dc-133">O arquivo **./ENV** no diretório raiz do projeto define as constantes que são usadas pelo projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-133">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="cc1dc-134">Algumas das constantes definidas neste arquivo são usadas para facilitar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-134">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="cc1dc-135">Talvez você queira atualizar os valores nesse arquivo para que eles correspondam ao seu cenário específico.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-135">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="cc1dc-136">Por exemplo, você pode atualizar o arquivo para especificar um escopo diferente, se o seu suplemento exigir algo diferente de `User.Read`.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-136">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="cc1dc-137">Configure o SSO</span><span class="sxs-lookup"><span data-stu-id="cc1dc-137">Configure SSO</span></span>

<span data-ttu-id="cc1dc-138">Nesse ponto, seu projeto de suplemento foi criado e contém o código necessário para facilitar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-138">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="cc1dc-139">Depois, execute as etapas a seguir para configurar o SSO do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-139">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="cc1dc-140">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-140">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="cc1dc-141">Execute o comando a seguir para configurar o SSO do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-141">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="cc1dc-142">Esse comando falhará se o locatário estiver configurado para exigir autenticação de dois fatores.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-142">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="cc1dc-143">Nesse cenário, será necessário concluir manualmente as etapas de configuração do SSO e registro do aplicativo Azure, conforme descrito no tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-143">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="cc1dc-144">Uma janela de navegador da Web será exibida e solicitará que você entre no Azure. </span><span class="sxs-lookup"><span data-stu-id="cc1dc-144">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="cc1dc-145">Entre no Azure com as suas credenciais de administrador do Office 365.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-145">Sign in to Azure using your Office 365 administrator credentials.</span></span> <span data-ttu-id="cc1dc-146">Essas credenciais serão usadas para registrar um novo aplicativo no Azure e definir as configurações necessárias para o SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-146">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cc1dc-147">Se você entrar no Azure usando credenciais de não administrador durante essa etapa, o script `configure-sso` não conseguirá fornecer consentimento de administrador para o suplemento aos usuários da organização.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-147">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="cc1dc-148">Portanto, o SSO não estará disponível aos usuários do suplemento e eles serão solicitados a entrar.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-148">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="cc1dc-149">Depois de inserir suas credenciais, feche a janela do navegador e retorne ao prompt de comando.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-149">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="cc1dc-150">Durante o processo de configuração do SSO, você verá mensagens de status sendo gravadas no console.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-150">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="cc1dc-151">Conforme descrito nas mensagens do console, os arquivos no projeto do suplemento que o gerador Yeoman criou são atualizados automaticamente com os dados necessários ao processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-151">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="cc1dc-152">Experimente</span><span class="sxs-lookup"><span data-stu-id="cc1dc-152">Try it out</span></span>

<span data-ttu-id="cc1dc-153">Se você tiver criado um suplemento do Excel, do Word ou do PowerPoint, conclua as etapas na seção a seguir para testá-lo. Se você criou um suplemento do Outlook, conclua as etapas na seção [Outlook](#outlook).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-153">If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="cc1dc-154">Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cc1dc-154">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="cc1dc-155">Execute as etapas a seguir para experimentar um suplemento do Excel, do Word ou do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-155">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="cc1dc-156">Quando o processo de configuração do SSO for concluído, execute o seguinte comando para criar o projeto: inicie o servidor Web local e sideload o suplemento no aplicativo cliente do Office selecionado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-156">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cc1dc-157">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-157">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="cc1dc-158">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-158">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="cc1dc-159">No aplicativo cliente do Office que é aberto ao executar o comando anterior (por exemplo, Excel, Word ou PowerPoint), certifique-se de estar conectado com um usuário que seja membro da mesma organização do Office 365, como uma conta de administrador do Office 365 que você usou para se conectar ao Azure, enquanto configura o SSO na etapa 3 da [seção anterior](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-159">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="cc1dc-160">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-160">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="cc1dc-161">No aplicativo cliente do Office, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-161">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="cc1dc-162">A imagem a seguir mostra esse botão no Excel.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-162">The following image shows this button in Excel.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="cc1dc-164">Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-164">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="cc1dc-165">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-165">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="cc1dc-166">Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante").</span><span class="sxs-lookup"><span data-stu-id="cc1dc-166">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="cc1dc-167">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-167">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="cc1dc-169">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-169">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="cc1dc-170">O suplemento recupera as informações de perfil do usuário conectado e as grava no documento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-170">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="cc1dc-171">A imagem a seguir mostra um exemplo de informações de perfil gravadas em uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-171">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Informações de perfil de usuário na planilha do Excel](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="cc1dc-173">Outlook</span><span class="sxs-lookup"><span data-stu-id="cc1dc-173">Outlook</span></span>

<span data-ttu-id="cc1dc-174">Execute as etapas a seguir para experimentar um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-174">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="cc1dc-175">Quando concluir o processo de configuração de SSO, execute o seguinte comando para criar o projeto e iniciar o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-175">When the SSO configuration process completes, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cc1dc-176">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-176">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="cc1dc-177">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-177">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="cc1dc-178">Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing)para realizar o sideload do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-178">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="cc1dc-179">Certifique-se de que você está conectado ao Outlook com um usuário que seja membro da mesma organização do Office 365, como a conta de administrador do Office 365 que você usou para se conectar ao Azure, enquanto configura o SSO na etapa 3 da [seção anterior](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-179">Make sure that you're signed in to Outlook with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="cc1dc-180">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-180">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="cc1dc-181">Escreva uma nova mensagem no Outlook.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-181">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="cc1dc-182">Na janela redigir mensagem, escolha o botão **Exibir painel de tarefas** na faixa de opções para abrir o painel de tarefas de suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-182">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="cc1dc-184">Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-184">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="cc1dc-185">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-185">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="cc1dc-186">Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante").</span><span class="sxs-lookup"><span data-stu-id="cc1dc-186">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="cc1dc-187">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-187">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="cc1dc-189">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-189">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="cc1dc-190">O suplemento recupera as informações de perfil do usuário conectado e as grava no corpo da mensagem do e-mail.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-190">The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.</span></span> 

    ![Informações de perfil de usuário na mensagem do Outlook](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="cc1dc-192">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="cc1dc-192">Next steps</span></span>

<span data-ttu-id="cc1dc-193">Parabéns, você criou com êxito um suplemento do painel de tarefas que usa SSO sempre que possível; e usa um método alternativo de autenticação de usuário quando não há suporte ao SSO.</span><span class="sxs-lookup"><span data-stu-id="cc1dc-193">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="cc1dc-194">Para saber mais sobre as etapas de configuração do SSO que o gerador Yeoman concluiu automaticamente e o código que facilita o processo de SSO, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="cc1dc-194">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="cc1dc-195">Confira também</span><span class="sxs-lookup"><span data-stu-id="cc1dc-195">See also</span></span>

- [<span data-ttu-id="cc1dc-196">Habilitar o logon único para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cc1dc-196">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="cc1dc-197">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="cc1dc-197">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="cc1dc-198">Solucionar problemas de mensagens de erro no logon único (SSO)</span><span class="sxs-lookup"><span data-stu-id="cc1dc-198">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)