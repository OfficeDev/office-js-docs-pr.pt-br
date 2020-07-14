---
title: Use o gerador Yeoman para criar um Suplemento do Office que use SSO (prévia)
description: Use o gerador Yeoman para criar um Suplemento do Office com Node.js que use logon único (prévia).
ms.date: 02/20/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 4ebe48054b06ae5022d57d3846b0f97b7c205164
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094460"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="ae07d-103">Use o gerador Yeoman para criar um Suplemento do Office que use logon único (prévia)</span><span class="sxs-lookup"><span data-stu-id="ae07d-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="ae07d-104">Neste artigo, você seguirá pelo processo de uso do gerador Yeoman para criar um Suplemento do Office para Excel, Outlook, Word ou PowerPoint que usa o logon único (SSO) sempre que possível, e usa um método alternativo de autenticação do usuário quando não há suporte ao SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="ae07d-105">Antes de tentar concluir o início rápido, revise [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md) para aprender conceitos básicos sobre o SSO em Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="ae07d-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="ae07d-106">O gerador Yeoman simplifica o processo de criação de um suplemento de SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="ae07d-107">Para um passo a passo detalhado descrevendo como concluir manualmente as etapas que o gerador Yeoman automatiza, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="ae07d-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ae07d-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="ae07d-108">Prerequisites</span></span>

* <span data-ttu-id="ae07d-109">[Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases)).</span><span class="sxs-lookup"><span data-stu-id="ae07d-109">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

* <span data-ttu-id="ae07d-110">A versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do [Yeoman gerador de suplementos do Office](https://github.com/OfficeDev/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="ae07d-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="ae07d-111">Se você estiver usando um Mac e não tiver a CLI do Azure instalada no computador, instale o [Homebrew](https://brew.sh/).</span><span class="sxs-lookup"><span data-stu-id="ae07d-111">If you're using a Mac and don't have the Azure CLI installed on your machine, you must install [Homebrew](https://brew.sh/).</span></span> <span data-ttu-id="ae07d-112">O script de configuração do SSO executado durante o início rápido usará o Homebrew para instalar a CLI do Azure e, em seguida, usará a CLI do Azure para configurar o SSO no Azure.</span><span class="sxs-lookup"><span data-stu-id="ae07d-112">The SSO configuration script that you'll run during this quick start will use Homebrew to install the Azure CLI, and will then use the Azure CLI to configure SSO within Azure.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="ae07d-113">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="ae07d-113">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="ae07d-114">O gerador Yeoman pode criar um Suplemento do Office habilitado para SSO do Excel, Outlook, Word ou PowerPoint e pode ser criado com o tipo de script JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="ae07d-114">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="ae07d-115">As instruções a seguir especificam o `JavaScript` e o `Excel`, mas você deverá escolher o tipo de script e o aplicativo cliente do Office que atendem melhor ao seu cenário.</span><span class="sxs-lookup"><span data-stu-id="ae07d-115">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="ae07d-116">**Escolha o tipo de projeto:** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="ae07d-116">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="ae07d-117">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="ae07d-117">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="ae07d-118">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="ae07d-118">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="ae07d-119">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="ae07d-119">**Which Office client application would you like to support?**</span></span> `Excel`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="ae07d-121">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="ae07d-121">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="ae07d-122">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="ae07d-122">Explore the project</span></span>

<span data-ttu-id="ae07d-123">O projeto de suplemento que você criou com o gerador do Yeoman contém um código para um suplemento de painel de tarefas habilitado para SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-123">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a><span data-ttu-id="ae07d-124">Configure o SSO</span><span class="sxs-lookup"><span data-stu-id="ae07d-124">Configure SSO</span></span>

<span data-ttu-id="ae07d-125">Nesse ponto, seu projeto de suplemento foi criado e contém o código necessário para facilitar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-125">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="ae07d-126">Depois, execute as etapas a seguir para configurar o SSO do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-126">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="ae07d-127">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="ae07d-127">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="ae07d-128">Execute o comando a seguir para configurar o SSO do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-128">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="ae07d-129">Esse comando falhará se o locatário estiver configurado para exigir autenticação de dois fatores.</span><span class="sxs-lookup"><span data-stu-id="ae07d-129">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="ae07d-130">Nesse cenário, será necessário concluir manualmente as etapas de configuração do SSO e registro do aplicativo Azure, conforme descrito no tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="ae07d-130">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="ae07d-131">Uma janela de navegador da Web será exibida e solicitará que você entre no Azure. </span><span class="sxs-lookup"><span data-stu-id="ae07d-131">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="ae07d-132">Entre no Azure com as suas credenciais de administrador do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="ae07d-132">Sign in to Azure using your Microsoft 365 administrator credentials.</span></span> <span data-ttu-id="ae07d-133">Essas credenciais serão usadas para registrar um novo aplicativo no Azure e definir as configurações necessárias para o SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-133">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ae07d-134">Se você entrar no Azure usando credenciais de não administrador durante essa etapa, o script `configure-sso` não conseguirá fornecer consentimento de administrador para o suplemento aos usuários da organização.</span><span class="sxs-lookup"><span data-stu-id="ae07d-134">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="ae07d-135">Portanto, o SSO não estará disponível aos usuários do suplemento e eles serão solicitados a entrar.</span><span class="sxs-lookup"><span data-stu-id="ae07d-135">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="ae07d-136">Depois de inserir suas credenciais, feche a janela do navegador e retorne ao prompt de comando.</span><span class="sxs-lookup"><span data-stu-id="ae07d-136">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="ae07d-137">Durante o processo de configuração do SSO, você verá mensagens de status sendo gravadas no console.</span><span class="sxs-lookup"><span data-stu-id="ae07d-137">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="ae07d-138">Conforme descrito nas mensagens do console, os arquivos no projeto do suplemento que o gerador Yeoman criou são atualizados automaticamente com os dados necessários ao processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-138">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="ae07d-139">Experimente</span><span class="sxs-lookup"><span data-stu-id="ae07d-139">Try it out</span></span>

<span data-ttu-id="ae07d-140">Se você tiver criado um suplemento do Excel, do Word ou do PowerPoint, conclua as etapas na seção a seguir para testá-lo. Se você criou um suplemento do Outlook, conclua as etapas na seção [Outlook](#outlook).</span><span class="sxs-lookup"><span data-stu-id="ae07d-140">If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="ae07d-141">Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ae07d-141">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="ae07d-142">Execute as etapas a seguir para experimentar um suplemento do Excel, do Word ou do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ae07d-142">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="ae07d-143">Quando o processo de configuração do SSO for concluído, execute o seguinte comando para criar o projeto: inicie o servidor Web local e sideload o suplemento no aplicativo cliente do Office selecionado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="ae07d-143">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ae07d-144">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-144">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="ae07d-145">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="ae07d-145">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="ae07d-146">No aplicativo cliente do Office que é aberto ao executar o comando anterior (por exemplo, Excel, Word ou PowerPoint), certifique-se de estar conectado com um usuário que seja membro da mesma organização do Microsoft 365, como uma conta de administrador do Microsoft 365 que você usou para se conectar ao Azure ao configurar o SSO na etapa 3 da [seção anterior](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="ae07d-146">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="ae07d-147">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="ae07d-147">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="ae07d-148">No aplicativo cliente do Office, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-148">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="ae07d-149">A imagem a seguir mostra esse botão no Excel.</span><span class="sxs-lookup"><span data-stu-id="ae07d-149">The following image shows this button in Excel.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="ae07d-151">Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-151">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="ae07d-152">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="ae07d-152">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="ae07d-153">Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft, ou com uma conta corporativa ou de estudante do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="ae07d-153">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="ae07d-154">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="ae07d-154">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="ae07d-156">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="ae07d-156">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="ae07d-157">O suplemento recupera as informações de perfil do usuário conectado e as grava no documento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-157">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="ae07d-158">A imagem a seguir mostra um exemplo de informações de perfil gravadas em uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="ae07d-158">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Informações de perfil de usuário na planilha do Excel](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="ae07d-160">Outlook</span><span class="sxs-lookup"><span data-stu-id="ae07d-160">Outlook</span></span>

<span data-ttu-id="ae07d-161">Execute as etapas a seguir para experimentar um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ae07d-161">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="ae07d-162">Quando concluir o processo de configuração de SSO, execute o seguinte comando para criar o projeto e iniciar o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="ae07d-162">When the SSO configuration process completes, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ae07d-163">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-163">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="ae07d-164">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="ae07d-164">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="ae07d-165">Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)para realizar o sideload do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ae07d-165">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="ae07d-166">Certifique-se de que você está conectado ao Outlook com um usuário que seja membro da mesma organização do Microsoft 365, como a conta de administrador do Microsoft 365 que você usou para se conectar ao Azure, ao configurar o SSO na etapa 3 da [seção anterior](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="ae07d-166">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="ae07d-167">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="ae07d-167">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="ae07d-168">Escreva uma nova mensagem no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ae07d-168">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="ae07d-169">Na janela redigir mensagem, escolha o botão **Exibir painel de tarefas** na faixa de opções para abrir o painel de tarefas de suplemento.</span><span class="sxs-lookup"><span data-stu-id="ae07d-169">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="ae07d-171">Na parte inferior do painel de tarefas, escolha o botão **Obter Informações do Meu Perfil de Usuário** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-171">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="ae07d-172">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="ae07d-172">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="ae07d-173">Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft, ou com uma conta corporativa ou de estudante do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="ae07d-173">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="ae07d-174">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="ae07d-174">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="ae07d-176">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="ae07d-176">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="ae07d-177">O suplemento recupera as informações de perfil do usuário conectado e as grava no corpo da mensagem do e-mail.</span><span class="sxs-lookup"><span data-stu-id="ae07d-177">The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.</span></span> 

    ![Informações de perfil de usuário na mensagem do Outlook](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="ae07d-179">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ae07d-179">Next steps</span></span>

<span data-ttu-id="ae07d-180">Parabéns, você criou com êxito um suplemento do painel de tarefas que usa SSO sempre que possível; e usa um método alternativo de autenticação de usuário quando não há suporte ao SSO.</span><span class="sxs-lookup"><span data-stu-id="ae07d-180">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="ae07d-181">Para saber como personalizar seu suplemento para adicionar novas funcionalidades que requerem permissões diferentes, consulte [Personalizar o suplemento habilitado para SSO do Node.js](sso-quickstart-customize.md).</span><span class="sxs-lookup"><span data-stu-id="ae07d-181">To learn about customizing your add-in to add new functionality that requires different permissions, see [Customize your Node.js SSO-enabled add-in](sso-quickstart-customize.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ae07d-182">Confira também</span><span class="sxs-lookup"><span data-stu-id="ae07d-182">See also</span></span>

- [<span data-ttu-id="ae07d-183">Habilitar o logon único para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ae07d-183">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- <span data-ttu-id="ae07d-184">[Personalizar o suplemento habilitado para SSO do Node.js](sso-quickstart-customize.md).</span><span class="sxs-lookup"><span data-stu-id="ae07d-184">[Customize your Node.js SSO-enabled add-in](sso-quickstart-customize.md)</span></span>
- [<span data-ttu-id="ae07d-185">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="ae07d-185">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="ae07d-186">Solucionar problemas de mensagens de erro no logon único (SSO)</span><span class="sxs-lookup"><span data-stu-id="ae07d-186">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)