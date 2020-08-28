---
title: Crie um Suplemento do Office com Node.js que use logon único
description: Aprenda a criar um suplemento baseado em node.js que usa o logon único do Office
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 2ca1cf37bade124498c99b0b25171871522c2bc7
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292873"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a><span data-ttu-id="96a94-103">Crie um Suplemento do Office com Node.js que use logon único</span><span class="sxs-lookup"><span data-stu-id="96a94-103">Create a Node.js Office Add-in that uses single sign-on</span></span>

<span data-ttu-id="96a94-p101">Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="96a94-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="96a94-106">Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e Express.</span><span class="sxs-lookup"><span data-stu-id="96a94-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> <span data-ttu-id="96a94-107">Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="96a94-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

> [!NOTE]
> <span data-ttu-id="96a94-108">Como alternativa para concluir as etapas descritas neste artigo, você pode usar o gerador Yeoman para criar um Suplemento do Office com Node.js habilitado para SSO.</span><span class="sxs-lookup"><span data-stu-id="96a94-108">As an alternative to completing the steps described in this article, you can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="96a94-109">O gerador Yeoman simplifica o processo de criação de um suplemento habilitado para SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO.</span><span class="sxs-lookup"><span data-stu-id="96a94-109">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="96a94-110">Para obter mais informações, confira [Início rápido de logon único (SSO)](../quickstarts/sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="96a94-110">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="96a94-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="96a94-111">Prerequisites</span></span>

* <span data-ttu-id="96a94-112">[Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="96a94-112">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="96a94-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="96a94-113">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="96a94-114">TypeScript, versão 3.6.2 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-114">TypeScript, version 3.6.2 or later</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="96a94-115">Um editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="96a94-115">A code editor.</span></span> <span data-ttu-id="96a94-116">Recomendamos o código do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="96a94-116">We recommend Visual Studio Code.</span></span>

* <span data-ttu-id="96a94-117">Pelo menos alguns arquivos e pastas armazenados no OneDrive for Business em sua assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="96a94-117">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="96a94-118">Uma assinatura do Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="96a94-118">A Microsoft Azure subscription.</span></span> <span data-ttu-id="96a94-119">Este suplemento requer o Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="96a94-119">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="96a94-120">O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização.</span><span class="sxs-lookup"><span data-stu-id="96a94-120">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="96a94-121">Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="96a94-121">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="96a94-122">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="96a94-122">Set up the starter project</span></span>

1. <span data-ttu-id="96a94-123">Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="96a94-123">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-124">Há três versões do exemplo:</span><span class="sxs-lookup"><span data-stu-id="96a94-124">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="96a94-p106">A pasta **inicial** é um projeto inicial. A interface de usuário e outros aspectos do suplemento que não estejam diretamente conectados ao SSO ou à autorização já foram feitos. Seções posteriores deste artigo orientam você durante o processo de conclusão.</span><span class="sxs-lookup"><span data-stu-id="96a94-p106">The **Begin** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="96a94-128">A versão **Complete** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo.</span><span class="sxs-lookup"><span data-stu-id="96a94-128">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="96a94-129">Para usar a versão completa, basta seguir as instruções deste artigo, mas substituir "Begin" por "concluído" e ignorar as seções **codificadas pelo cliente** e **codificar o** lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="96a94-129">To use the completed version, just follow the instructions in this article, but replace "Begin" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="96a94-130">A versão **SSOAutoSetup** é um exemplo concluído que automatiza a maioria das etapas para registrar o suplemento com o Azure AD e configurá-lo.</span><span class="sxs-lookup"><span data-stu-id="96a94-130">The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it.</span></span> <span data-ttu-id="96a94-131">Use esta versão se desejar ver um suplemento de trabalho com SSO rapidamente.</span><span class="sxs-lookup"><span data-stu-id="96a94-131">Use this version if you want to see a working add-in with SSO quickly.</span></span> <span data-ttu-id="96a94-132">Basta seguir as etapas no README da pasta.</span><span class="sxs-lookup"><span data-stu-id="96a94-132">Just follow the steps in the Readme of the folder.</span></span> <span data-ttu-id="96a94-133">É recomendável que, em algum momento, você siga as etapas de configuração e registro manuais deste artigo para entender melhor a relação entre o Azure AD e um suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-133">We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in.</span></span> 

1. <span data-ttu-id="96a94-134">Abra um prompt de comando na pasta **Iniciar** .</span><span class="sxs-lookup"><span data-stu-id="96a94-134">Open a command prompt in the **Begin** folder.</span></span>

1. <span data-ttu-id="96a94-135">Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="96a94-135">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="96a94-136">Execute o comando `npm run install-dev-certs`.</span><span class="sxs-lookup"><span data-stu-id="96a94-136">Run the command `npm run install-dev-certs`.</span></span> <span data-ttu-id="96a94-137">Selecione **Sim** à solicitação para instalar o certificado.</span><span class="sxs-lookup"><span data-stu-id="96a94-137">Select **Yes** to the prompt to install the certificate.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="96a94-138">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="96a94-138">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="96a94-139">Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="96a94-139">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="96a94-140">Entre com as credenciais de ***administrador*** em seu Microsoft 365 locação.</span><span class="sxs-lookup"><span data-stu-id="96a94-140">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="96a94-141">Por exemplo, MeuNome@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="96a94-141">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="96a94-142">Selecione **Novo registro**.</span><span class="sxs-lookup"><span data-stu-id="96a94-142">Select **New registration**.</span></span> <span data-ttu-id="96a94-143">Na página **Registrar um aplicativo**, defina os valores da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="96a94-143">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="96a94-144">Defina **Nome** para `Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="96a94-144">Set **Name** to `Office-Add-in-NodeJS-SSO`.</span></span>
    * <span data-ttu-id="96a94-145">Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="96a94-145">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span>
    * <span data-ttu-id="96a94-146">Defina o tipo de aplicativo como **Web** e, em seguida, defina **URI de redirecionamento** como ` https://localhost:44355/dialog.html` .</span><span class="sxs-lookup"><span data-stu-id="96a94-146">Set the application type to **Web** and then set **Redirect URI** to ` https://localhost:44355/dialog.html`.</span></span>
    * <span data-ttu-id="96a94-147">Escolha **Registrar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-147">Choose **Register**.</span></span>

1. <span data-ttu-id="96a94-148">Na página **Office-Add-in-NodeJS-SSO**, copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**.</span><span class="sxs-lookup"><span data-stu-id="96a94-148">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="96a94-149">Use ambos os valores nos procedimentos posteriores.</span><span class="sxs-lookup"><span data-stu-id="96a94-149">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-150">Essa ID é o valor "Audience" quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo.</span><span class="sxs-lookup"><span data-stu-id="96a94-150">This ID is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="96a94-151">Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-151">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="96a94-152">Selecione **Autenticação** em **Gerenciar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-152">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="96a94-153">Na seção **concessão implícita** , habilite as caixas de seleção para token de **acesso** e **token de ID**.</span><span class="sxs-lookup"><span data-stu-id="96a94-153">In the **Implicit grant** section, enable the checkboxes for both **Access token** and **ID token**.</span></span> <span data-ttu-id="96a94-154">O exemplo tem um sistema de autorização de fallback que é chamado quando o SSO não está disponível.</span><span class="sxs-lookup"><span data-stu-id="96a94-154">The sample has a fallback authorization system that is invoked when SSO is not available.</span></span> <span data-ttu-id="96a94-155">Esse sistema usa o fluxo implícito.</span><span class="sxs-lookup"><span data-stu-id="96a94-155">This system uses the Implicit Flow.</span></span>

1. <span data-ttu-id="96a94-156">Na parte superior da página, selecione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-156">Select **Save** at the top of the form.</span></span>

1. <span data-ttu-id="96a94-157">Selecione **Certificados e segredos** sob **Gerenciar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-157">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="96a94-158">Selecione o botão **Novo segredo do cliente**.</span><span class="sxs-lookup"><span data-stu-id="96a94-158">Select the **New client secret** button.</span></span> <span data-ttu-id="96a94-159">Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-159">Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="96a94-160">*Copiar o valor de segredo do cliente imediatamente e salvá-lo com a ID de aplicativo* antes de prosseguir, pois ele será necessário em um procedimento posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-160">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="96a94-161">Selecionar **Expor uma API** em **Gerenciar**.</span><span class="sxs-lookup"><span data-stu-id="96a94-161">Select **Expose an API** under **Manage**.</span></span> <span data-ttu-id="96a94-162">Selecione o link **definir** .</span><span class="sxs-lookup"><span data-stu-id="96a94-162">Select the **Set** link.</span></span> <span data-ttu-id="96a94-163">Isso gerará o URI da ID do aplicativo no formato "api://$App ID GUID $", onde $App GUID de ID $ é a **ID do aplicativo (cliente)**.</span><span class="sxs-lookup"><span data-stu-id="96a94-163">This will generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span>

1. <span data-ttu-id="96a94-164">Na ID gerada, insira `localhost:44355/` (Observe a barra "/" anexada ao final) entre as barras duplas e o GUID.</span><span class="sxs-lookup"><span data-stu-id="96a94-164">In the generated ID, insert `localhost:44355/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="96a94-165">Quando você terminar, a ID inteira deverá ter a forma `api://localhost:44355/$App ID GUID$` ; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7` .</span><span class="sxs-lookup"><span data-stu-id="96a94-165">When you are finished, the entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="96a94-166">Selecione o botão **Adicionar um escopo**.</span><span class="sxs-lookup"><span data-stu-id="96a94-166">Select the **Add a scope** button.</span></span> <span data-ttu-id="96a94-167">No painel que se abre, insira `access_as_user` como o **Nome de escopo**.</span><span class="sxs-lookup"><span data-stu-id="96a94-167">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="96a94-168">Definir **Quem pode consentir?** aos **Administradores e usuários**.</span><span class="sxs-lookup"><span data-stu-id="96a94-168">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="96a94-169">Preencha os campos para configurar os prompts de consentimento de usuário e administrador com os valores que são apropriados para o `access_as_user` escopo, o que permite que o aplicativo cliente do Office Use as APIs Web do seu suplemento com os mesmos direitos que o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="96a94-169">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="96a94-170">Sugestões:</span><span class="sxs-lookup"><span data-stu-id="96a94-170">Suggestions:</span></span>

    - <span data-ttu-id="96a94-171">**Nome para exibição do consentimento do administrador**: o Office pode atuar como o usuário.</span><span class="sxs-lookup"><span data-stu-id="96a94-171">**Admin consent display name**: Office can act as the user.</span></span>
    - <span data-ttu-id="96a94-172">**Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="96a94-172">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="96a94-173">**Nome para exibição do consentimento do usuário**: o Office pode agir como você.</span><span class="sxs-lookup"><span data-stu-id="96a94-173">**User consent display name**: Office can act as you.</span></span>
    - <span data-ttu-id="96a94-174">**Descrição do consentimento do usuário**: habilitar o Office para chamar as APIs Web do suplemento com os mesmos direitos que você tem.</span><span class="sxs-lookup"><span data-stu-id="96a94-174">**User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="96a94-175">Verifique se o **Estado** está definido como **Habilitado**.</span><span class="sxs-lookup"><span data-stu-id="96a94-175">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="96a94-176">Selecione **Adicionar escopo**.</span><span class="sxs-lookup"><span data-stu-id="96a94-176">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-177">A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="96a94-177">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="96a94-178">Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-178">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="96a94-179">Cada uma das seguintes IDs precisa ser pré-autorizada.</span><span class="sxs-lookup"><span data-stu-id="96a94-179">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="96a94-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="96a94-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="96a94-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="96a94-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="96a94-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="96a94-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="96a94-183">`08e18876-6177-487e-b8b5-cf950c1e598c`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="96a94-183">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
    - <span data-ttu-id="96a94-184">`bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)</span><span class="sxs-lookup"><span data-stu-id="96a94-184">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="96a94-185">Para cada ID, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="96a94-185">For each ID, take these steps:</span></span>

    <span data-ttu-id="96a94-186">a.</span><span class="sxs-lookup"><span data-stu-id="96a94-186">a.</span></span> <span data-ttu-id="96a94-187">Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o ID do cliente para o respectivo GUID e marque a caixa `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="96a94-187">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="96a94-188">b.</span><span class="sxs-lookup"><span data-stu-id="96a94-188">b.</span></span> <span data-ttu-id="96a94-189">Selecione **Adicionar aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="96a94-189">Select **Add application**.</span></span>

1. <span data-ttu-id="96a94-190">Selecione **Permissões para API** em **Gerenciar** e selecione **Adicionar uma permissão**.</span><span class="sxs-lookup"><span data-stu-id="96a94-190">Select **API permissions** under **Manage** and select **Add a permission**.</span></span> <span data-ttu-id="96a94-191">No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.</span><span class="sxs-lookup"><span data-stu-id="96a94-191">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="96a94-192">Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="96a94-192">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="96a94-193">Selecione estas opções.</span><span class="sxs-lookup"><span data-stu-id="96a94-193">Select the following.</span></span> <span data-ttu-id="96a94-194">Somente o primeiro é realmente necessário para o seu suplemento; Mas a `profile` permissão é necessária para que o aplicativo do Office obtenha um token para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-194">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office application to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="96a94-195">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="96a94-195">Files.Read.All</span></span>
    * <span data-ttu-id="96a94-196">perfil</span><span class="sxs-lookup"><span data-stu-id="96a94-196">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-197">A permissão `User.Read` pode já estar listada por padrão.</span><span class="sxs-lookup"><span data-stu-id="96a94-197">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="96a94-198">É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.</span><span class="sxs-lookup"><span data-stu-id="96a94-198">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="96a94-199">Marque a caixa de seleção para cada permissão conforme elas forem exibidas.</span><span class="sxs-lookup"><span data-stu-id="96a94-199">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="96a94-200">Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.</span><span class="sxs-lookup"><span data-stu-id="96a94-200">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="96a94-201">Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Sim** para a confirmação exibida.</span><span class="sxs-lookup"><span data-stu-id="96a94-201">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.</span></span>

## <a name="configure-the-add-in"></a><span data-ttu-id="96a94-202">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="96a94-202">Configure the add-in</span></span>

1. <span data-ttu-id="96a94-203">Abra a pasta `\Begin` no projeto clonado no editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="96a94-203">Open the `\Begin` folder in the cloned project in your code editor.</span></span>

1. <span data-ttu-id="96a94-204">Abra o arquivo `.ENV` e use os valores que você copiou anteriormente.</span><span class="sxs-lookup"><span data-stu-id="96a94-204">Open the `.ENV` file and use the values that you copied earlier.</span></span> <span data-ttu-id="96a94-205">Defina o **CLIENT_ID** para a identificação do seu **ID de aplicativo (cliente)** e defina **CLIENT_SECRET** para o seu segredo de cliente.</span><span class="sxs-lookup"><span data-stu-id="96a94-205">Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret.</span></span> <span data-ttu-id="96a94-206">Os valores **não** devem estar entre aspas.</span><span class="sxs-lookup"><span data-stu-id="96a94-206">The values should **not** be in quotation marks.</span></span> <span data-ttu-id="96a94-207">Quando terminar, o arquivo deverá ser semelhante ao seguinte:</span><span class="sxs-lookup"><span data-stu-id="96a94-207">When you are done, the file should be similar to the following:</span></span> 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. <span data-ttu-id="96a94-208">Abra o arquivo `\public\javascripts\fallbackAuthDialog.js`.</span><span class="sxs-lookup"><span data-stu-id="96a94-208">Open the `\public\javascripts\fallbackAuthDialog.js` file.</span></span> <span data-ttu-id="96a94-209">Na declaração `msalConfig` substitua o espaço reservado "{application_GUID here}", pela ID do Aplicativo que você copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-209">In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="96a94-210">O valor deve estar entre aspas.</span><span class="sxs-lookup"><span data-stu-id="96a94-210">The value should be in quotation marks.</span></span>

1. <span data-ttu-id="96a94-211">Abra o arquivo de manifesto de suplemento "manifest\ manifest_local.xml" e role até a parte inferior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-211">Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file.</span></span> <span data-ttu-id="96a94-212">Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="96a94-212">Just above the `</VersionOverrides>` end tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="96a94-213">Substitua o espaço reservado "{$application_GUID here$}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-213">Replace the placeholder "$application_GUID here$" *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="96a94-214">O símbolo "$" não faz parte da ID, portanto não o inclua.</span><span class="sxs-lookup"><span data-stu-id="96a94-214">The "$" symbols are not part of the ID, so do not include them.</span></span> <span data-ttu-id="96a94-215">Esta é a mesma ID usada para o CLIENT_ID e audiência no. ENV arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-215">This is the same ID you used in for the CLIENT_ID and Audience in the .ENV file.</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-216">O valor **Recurso** é o**URI da ID de aplicativo** que você definiu quando registrou o suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-216">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="96a94-217">A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="96a94-217">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="96a94-218">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="96a94-218">Code the client-side</span></span>

### <a name="create-the-sso-logic"></a><span data-ttu-id="96a94-219">Criar a lógica de SSO</span><span class="sxs-lookup"><span data-stu-id="96a94-219">Create the SSO logic</span></span>

1. <span data-ttu-id="96a94-220">No editor de códigos, abra o arquivo `public\javascripts\ssoAuthES6.js`.</span><span class="sxs-lookup"><span data-stu-id="96a94-220">In your code editor, open the file `public\javascripts\ssoAuthES6.js`.</span></span> <span data-ttu-id="96a94-221">Ele já tem um código que garante que o Promises seja suportado, mesmo no Internet Explorer 11, e uma chamada`Office.onReady` para atribuir um manipulador para o botão somente suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-221">It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.</span></span>

    > [!NOTE]
    > <span data-ttu-id="96a94-222">Como o nome sugere, o ssoAuthES6.js usa a sintaxe JavaScript ES6, pois usar `async` e `await` mostra melhor a simplicidade fundamental da API de SSO.</span><span class="sxs-lookup"><span data-stu-id="96a94-222">As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API.</span></span> <span data-ttu-id="96a94-223">Quando o servidor localhost for iniciado, esse arquivo será transformado em uma sintaxe ES5 para que o exemplo seja executado no Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="96a94-223">When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11.</span></span> 

1. <span data-ttu-id="96a94-224">Adicione o seguinte código abaixo do método Office. onReady:</span><span class="sxs-lookup"><span data-stu-id="96a94-224">Add the following code below the Office.onReady method:</span></span>

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exchange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. <span data-ttu-id="96a94-225">Substitua `TODO 1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-225">Replace `TODO 1` with the following code.</span></span> <span data-ttu-id="96a94-226">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-226">About this code, note:</span></span>

    - <span data-ttu-id="96a94-227">`OfficeRuntime.auth.getAccessToken` instrui o Office a obter um token de bootstrap do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="96a94-227">`OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD.</span></span> <span data-ttu-id="96a94-228">Um token de bootstrap é semelhante a um token de ID, mas tem uma propriedade `scp` (Scope) com o valor `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="96a94-228">A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`.</span></span> <span data-ttu-id="96a94-229">Esse tipo de token pode ser trocado por um aplicativo Web para um token de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-229">This kind of token can be exchanged by a web application for an access token to Microsoft Graph.</span></span>
    - <span data-ttu-id="96a94-230">Definir a `allowSignInPrompt` opção como true significa que, se nenhum usuário estiver conectado ao Office no momento, o Office abrirá um prompt de entrada pop-up.</span><span class="sxs-lookup"><span data-stu-id="96a94-230">Setting the `allowSignInPrompt` option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.</span></span>
    - <span data-ttu-id="96a94-231">A definição da `allowConsentPrompt` opção como true significa que, se o usuário não tiver sido remetido para permitir que o suplemento acesse o perfil AAD do usuário, o Office abrirá um prompt de consentimento.</span><span class="sxs-lookup"><span data-stu-id="96a94-231">Setting the `allowConsentPrompt` option to true means that if the user has not consented to let the add-in access the user's AAD profile, then Office will open a consent prompt.</span></span> <span data-ttu-id="96a94-232">(O prompt permite que o usuário consentisse no perfil AAD do usuário, e não nos escopos do Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="96a94-232">(The prompt only allows the user to consent to the user's AAD profile, not to Microsoft Graph scopes.)</span></span>
    - <span data-ttu-id="96a94-233">Definir a `forMSGraphAccess` opção como true indica ao Office que o suplemento pretende usar o token de inicialização para obter um token de acesso ao Microsoft Graph, em vez de apenas usá-lo como um token de ID.</span><span class="sxs-lookup"><span data-stu-id="96a94-233">Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Microsoft Graph, instead of just using it as an ID token.</span></span> <span data-ttu-id="96a94-234">Se o administrador locatário não tiver concedido consentimento para o acesso do suplemento ao Microsoft Graph, `OfficeRuntime.auth.getAccessToken` retornará o erro **13012**.</span><span class="sxs-lookup"><span data-stu-id="96a94-234">If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**.</span></span> <span data-ttu-id="96a94-235">O suplemento pode responder voltando para um sistema alternativo de autorização. Isso é necessário porque o Office pode solicitar apenas consentimento para o perfil do Azure AD do usuário, não para escopos do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-235">The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes.</span></span> <span data-ttu-id="96a94-236">O sistema de autorização de fallback exige que o usuário entre novamente e o usuário *pode* ser solicitado a se concordar com escopos do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-236">The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Microsoft Graph scopes.</span></span> <span data-ttu-id="96a94-237">Portanto, a opção `forMSGraphAccess` garante que o suplemento não fará uma troca de tokens que falhará devido à falta de consentimento.</span><span class="sxs-lookup"><span data-stu-id="96a94-237">So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent.</span></span> <span data-ttu-id="96a94-238">Uma vez que você concedeu consentimento de administrador em uma etapa anterior, esse cenário não acontecerá para esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-238">(Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in.</span></span> <span data-ttu-id="96a94-239">No entanto, a opção é incluída aqui para ilustrar uma prática recomendada.</span><span class="sxs-lookup"><span data-stu-id="96a94-239">But the option is included here anyway to illustrate a best practice.)</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true }); 
    ```

1. <span data-ttu-id="96a94-240">Substitua `TODO 2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-240">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="96a94-241">Você criará o método `getGraphToken` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-241">You'll create the `getGraphToken` method in a later step.</span></span>

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. <span data-ttu-id="96a94-242">Substitua `TODO 3` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-242">Replace `TODO 3` with the following.</span></span> <span data-ttu-id="96a94-243">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-243">About this code, note:</span></span> 

    - <span data-ttu-id="96a94-244">Se o Microsoft 365 locatário tiver sido configurado para exigir a autenticação multifator, o `exchangeResponse` incluirá uma `claims` propriedade com informações sobre os outros fatores necessários.</span><span class="sxs-lookup"><span data-stu-id="96a94-244">If the Microsoft 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors.</span></span> <span data-ttu-id="96a94-245">Nesse caso, `OfficeRuntime.auth.getAccessToken` deve ser chamado novamente com a opção `authChallenge` definida como o valor da propriedade de declarações.</span><span class="sxs-lookup"><span data-stu-id="96a94-245">In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property.</span></span> <span data-ttu-id="96a94-246">Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas.</span><span class="sxs-lookup"><span data-stu-id="96a94-246">This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. <span data-ttu-id="96a94-247">Substitua `TODO 4` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-247">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="96a94-248">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-248">About this code, note:</span></span> 

    - <span data-ttu-id="96a94-249">Você criará o método `handleAADErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-249">You'll create the `handleAADErrors` method in a later step.</span></span> <span data-ttu-id="96a94-250">Os erros do Azure AD são retornados para o cliente como respostas HTTP # 200.</span><span class="sxs-lookup"><span data-stu-id="96a94-250">Azure AD errors are returned to the client as HTTP code 200 Responses.</span></span> <span data-ttu-id="96a94-251">Eles não geram erros, portanto, não disparam o bloco `catch` do método `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="96a94-251">They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.</span></span>
    - <span data-ttu-id="96a94-252">Você criará o método `makeGraphApiCall` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-252">You'll create the `makeGraphApiCall` method in a later step.</span></span> <span data-ttu-id="96a94-253">Ele faz uma chamada AJAX para o ponto de extremidade do MS Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-253">It makes an AJAX call to the MS Graph endpoint.</span></span> <span data-ttu-id="96a94-254">Os erros são detectados na callback`.fail` da chamada, não no bloco `catch` do método `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="96a94-254">Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.</span></span>

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. <span data-ttu-id="96a94-255">Substitua `TODO 5` pelo seguinte</span><span class="sxs-lookup"><span data-stu-id="96a94-255">Replace `TODO 5` with the following</span></span>

    - <span data-ttu-id="96a94-256">Os erros da chamada `getAccessToken` terão uma propriedade `code` com um número de erro, normalmente no intervalo 13xxx.</span><span class="sxs-lookup"><span data-stu-id="96a94-256">Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range.</span></span> <span data-ttu-id="96a94-257">Você criará o método `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-257">You'll create the `handleClientSideErrors` method in a later step.</span></span>
    - <span data-ttu-id="96a94-258">O método `showMessage` exibe o texto no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="96a94-258">The `showMessage` method displays text on the task pane.</span></span>

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. <span data-ttu-id="96a94-259">Abaixo do método `getGraphData`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-259">Below the `getGraphData` method, add the following function.</span></span> <span data-ttu-id="96a94-260">Observe que `/auth` é uma rota expressa do servidor que troca o token de inicialização com o Azure ad para obter um token de acesso para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-260">Note that `/auth` is a server-side Express route that exchanges the bootstrap token with Azure AD for an access token to Microsoft Graph.</span></span>

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. <span data-ttu-id="96a94-261">Abaixo do método `getGraphToken`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-261">Below the `getGraphToken` method, add the following function.</span></span> <span data-ttu-id="96a94-262">Observe que `error.code` é um número, normalmente no intervalo 13xxx.</span><span class="sxs-lookup"><span data-stu-id="96a94-262">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="96a94-263">Substitua `TODO 6` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-263">Replace `TODO 6` with the following code.</span></span> <span data-ttu-id="96a94-264">Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="96a94-264">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span> 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the web.
        showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. <span data-ttu-id="96a94-265">Substitua `TODO 7` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-265">Replace `TODO 7` with the following code.</span></span> <span data-ttu-id="96a94-266">Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office](troubleshoot-sso-in-office-add-ins.md). A função `dialogFallback` invoca o sistema de autorização alternativo.</span><span class="sxs-lookup"><span data-stu-id="96a94-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization.</span></span> <span data-ttu-id="96a94-267">Neste suplemento, o sistema de fallback abre uma caixa de diálogo que exige que o usuário entre, mesmo que o usuário já esteja, e use o msal.js e Implicit Flow para obter um token de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-267">In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. <span data-ttu-id="96a94-268">Abaixo da função `handleClientSideErrors`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-268">Below the `handleClientSideErrors` function, add the following function.</span></span> 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. <span data-ttu-id="96a94-269">Em raras ocasiões, o token de bootstrap no cache do Office fica não vencido quando o Office o valida, mas vence no momento em que ele atinge o Azure AD para o Exchange.</span><span class="sxs-lookup"><span data-stu-id="96a94-269">On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange.</span></span> <span data-ttu-id="96a94-270">O Azure AD responderá com o erro **AADSTS500133**.</span><span class="sxs-lookup"><span data-stu-id="96a94-270">Azure AD will respond with error **AADSTS500133**.</span></span> <span data-ttu-id="96a94-271">Nesse caso, o suplemento deve simplesmente ligar recursivamente o `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="96a94-271">In this case, the add-in should simply recursively call `getGraphData`.</span></span> <span data-ttu-id="96a94-272">Como o token de inicialização em cache já expirou, o Office receberá um novo token do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="96a94-272">Since the cached bootstrap token is now expired, Office will get a new one from Azure AD.</span></span> <span data-ttu-id="96a94-273">Portanto, substitua `TODO 8` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-273">So replace `TODO 8` with the following.</span></span> 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
    {
        getGraphData();
    }
    ```

1. <span data-ttu-id="96a94-274">Para garantir que o suplemento não insira um loop infinito de chamadas para `getGraphData`, o suplemento deve controlar quantas vezes `getGraphData` foi chamado e ter a certeza de que o não é chamado recursivomente chamado mais de uma vez.</span><span class="sxs-lookup"><span data-stu-id="96a94-274">To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once.</span></span> <span data-ttu-id="96a94-275">Portanto, crie uma variável de contador em um escopo global para as funções `handleAADErrors` e `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="96a94-275">So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions.</span></span> <span data-ttu-id="96a94-276">Um bom lugar para as variáveis globais está logo abaixo da chamada de método `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="96a94-276">A good place for global variables is just below the `Office.onReady` method call.</span></span>

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. <span data-ttu-id="96a94-277">Altere a estrutura `if` no método `handleAADErrors` para que ela:</span><span class="sxs-lookup"><span data-stu-id="96a94-277">Change the `if` structure in the `handleAADErrors` method so that it:</span></span>

    - <span data-ttu-id="96a94-278">Incremente o contador antes de chamar `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="96a94-278">Increments the counter just before it calls `getGraphData`.</span></span>
    - <span data-ttu-id="96a94-279">Teste para garantir que `getGraphData` ainda não tenha sido chamado pela segunda vez.</span><span class="sxs-lookup"><span data-stu-id="96a94-279">Tests to ensure that `getGraphData` has not already been called a second time.</span></span> 

    <span data-ttu-id="96a94-280">Portanto, a versão final da estrutura `if` deve ter a seguinte aparência:</span><span class="sxs-lookup"><span data-stu-id="96a94-280">So the final version of the `if` structure should look like the following:</span></span>

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="96a94-281">Substitua `TODO 9` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-281">Replace `TODO 9` with the following.</span></span> 

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="96a94-282">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-282">Save and close the file.</span></span>

### <a name="get-the-data-and-add-it-to-the-office-document"></a><span data-ttu-id="96a94-283">Obtenha os dados e adicione-os ao documento do Office</span><span class="sxs-lookup"><span data-stu-id="96a94-283">Get the data and add it to the Office document</span></span>

1. <span data-ttu-id="96a94-284">Na pasta `public\javascripts`, crie um novo arquivo chamado `data.js`.</span><span class="sxs-lookup"><span data-stu-id="96a94-284">In the `public\javascripts` folder, create a new file named `data.js`.</span></span>

1. <span data-ttu-id="96a94-285">Adicione a seguinte função ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-285">Add the following function to the file.</span></span> <span data-ttu-id="96a94-286">Esta é a função que é chamada pela função `getGraphData` quando tiver adquirido um token de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-286">This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph.</span></span> 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. <span data-ttu-id="96a94-287">Substitua `TODO 10` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-287">Replace `TODO 10` with the following.</span></span> <span data-ttu-id="96a94-288">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-288">About this code, note:</span></span> 

    - <span data-ttu-id="96a94-289">Esse objeto é o parâmetro para o método `$.ajax`.</span><span class="sxs-lookup"><span data-stu-id="96a94-289">This object is the parameter to the `$.ajax` method.</span></span>
    - <span data-ttu-id="96a94-290">O `/getuserdata` é uma rota expressa no servidor do suplemento criado em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-290">The `/getuserdata` is an Express route on the add-in's server that you create in a later step.</span></span> <span data-ttu-id="96a94-291">Ele chamará um ponto de extremidade do Microsoft Graph e incluiremos o token de acesso em sua chamada.</span><span class="sxs-lookup"><span data-stu-id="96a94-291">It will call a Microsoft Graph endpoint and include the access token in its call.</span></span> 

    ```javascript
    {
        type: "GET",
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. <span data-ttu-id="96a94-292">Substitua `TODO11` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-292">Replace `TODO11` with the following.</span></span> <span data-ttu-id="96a94-293">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-293">About this code, note:</span></span>

    - <span data-ttu-id="96a94-294">`writeFileNamesToOfficeDocument` inserirá os dados do gráfico no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="96a94-294">The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document.</span></span> <span data-ttu-id="96a94-295">Ela é definida no arquivo `public\javascripts\document.js`.</span><span class="sxs-lookup"><span data-stu-id="96a94-295">It is defined in the `public\javascripts\document.js` file.</span></span> 
    - <span data-ttu-id="96a94-296">Se `writeFileNamesToOfficeDocument` retornar um erro, ele começará com "não é possível adicionar nomes de arquivo ao documento".</span><span class="sxs-lookup"><span data-stu-id="96a94-296">If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."</span></span>

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () {
        showMessage("Your data has been added to the document.");
    })
    .catch(function (error) {
        showMessage(error);
    });
    ```

1. <span data-ttu-id="96a94-297">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-297">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="96a94-298">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="96a94-298">Code the server-side</span></span>

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a><span data-ttu-id="96a94-299">Crie o roteador de autenticação e a lógica de troca de tokens</span><span class="sxs-lookup"><span data-stu-id="96a94-299">Create the auth router and the token exchange logic</span></span>

1. <span data-ttu-id="96a94-300">Abra o arquivo `routes\authRoute.js` e adicione a seguinte função de rota logo abaixo das instruções `require` e acima da instrução `module.exports`.</span><span class="sxs-lookup"><span data-stu-id="96a94-300">Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement.</span></span> <span data-ttu-id="96a94-301">Observe que o parâmetro de URL de `router.get` é '/'.</span><span class="sxs-lookup"><span data-stu-id="96a94-301">Note that the URL parameter of `router.get` is '/'.</span></span> <span data-ttu-id="96a94-302">Como esta rota está sendo definida em um roteador que tratará todas as solicitações HTTP para a URL "/auth", esta rota manipula todas as solicitações de "/auth".</span><span class="sxs-lookup"><span data-stu-id="96a94-302">Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'.</span></span> <span data-ttu-id="96a94-303">A função `getGraphToken` do lado do cliente que você criou anteriormente chama essa rota.</span><span class="sxs-lookup"><span data-stu-id="96a94-303">The client-side `getGraphToken` function that you created earlier calls this route.</span></span>  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exchange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. <span data-ttu-id="96a94-304">Substitua `TODO 12` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-304">Replace `TODO 12` with the following code.</span></span>

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. <span data-ttu-id="96a94-305">Substitua `TODO 13` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96a94-305">Replace `TODO 13` with the following code.</span></span> <span data-ttu-id="96a94-306">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-306">About this code, note:</span></span> 

    - <span data-ttu-id="96a94-307">Este é o início de um bloco `else` longo, mas o `}` de fechamento não está no final, já que você adicionará mais código a ele.</span><span class="sxs-lookup"><span data-stu-id="96a94-307">This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it.</span></span> 
    - <span data-ttu-id="96a94-308">A cadeia de caracteres `authorization` é um "transportador" seguido pelo token bootstrap, portanto, a primeira linha do bloco `else` está atribuindo o token para `jwt`.</span><span class="sxs-lookup"><span data-stu-id="96a94-308">The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`.</span></span> <span data-ttu-id="96a94-309">("JWT" significa "JSON Web Token".)</span><span class="sxs-lookup"><span data-stu-id="96a94-309">("JWT" stands for "JSON Web Token".)</span></span>
    - <span data-ttu-id="96a94-310">Os dois valores `process.env.*` são as constantes que você atribuiu ao configurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-310">The two `process.env.*` values are the constants that you assigned when you configured the add-in.</span></span> 
    - <span data-ttu-id="96a94-311">O parâmetro de formulário `requested_token_use` está definido como ' on_behalf_of '.</span><span class="sxs-lookup"><span data-stu-id="96a94-311">The `requested_token_use` form parameter is set to 'on_behalf_of'.</span></span> <span data-ttu-id="96a94-312">Isso informa ao Azure AD que o suplemento está solicitando um token de acesso ao Microsoft Graph usando o fluxo On-Behalf-Of.</span><span class="sxs-lookup"><span data-stu-id="96a94-312">This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow.</span></span> <span data-ttu-id="96a94-313">O Azure responderá validando que o token de bootstrap, que é atribuído ao parâmetro de formulário `assertion`, tem uma propriedade `scp` que está definida como `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="96a94-313">Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.</span></span>
    - <span data-ttu-id="96a94-314">O parâmetro de formulário `scope` está definido como "Files.Read.All', que é o único escopo do Microsoft Graph necessário para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="96a94-314">The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.</span></span>

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. <span data-ttu-id="96a94-315">Substitua `TODO 14` pelo código a seguir, que completa o bloco `else`.</span><span class="sxs-lookup"><span data-stu-id="96a94-315">Replace `TODO 14` with the following code, which completes the `else` block.</span></span> <span data-ttu-id="96a94-316">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-316">About this code, note:</span></span>

    - <span data-ttu-id="96a94-317">A constante `tenant` é definida como "comum" porque você configurou o suplemento como multilocatário ao registrá-lo no Azure AD, especificamente quando você define **Tipos de conta com suporte** para **Contas em qualquer diretório corporativo e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="96a94-317">The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span> <span data-ttu-id="96a94-318">Se, em vez disso, você optou por suportar apenas contas no mesmo locatário do Microsoft 365 em que o suplemento está registrado, o código `tenant` seria definido como o GUID do locatário.</span><span class="sxs-lookup"><span data-stu-id="96a94-318">If you had instead chosen to support only accounts in the same Microsoft 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant.</span></span> 
    - <span data-ttu-id="96a94-319">Se a solicitação POST não for recebida, a resposta do Azure AD será convertida para JSON e enviada para o cliente.</span><span class="sxs-lookup"><span data-stu-id="96a94-319">If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client.</span></span> <span data-ttu-id="96a94-320">Esse objeto JSON tem uma propriedade `access_token` à qual o Azure AD atribuiu o token de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96a94-320">This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.</span></span>

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();

            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. <span data-ttu-id="96a94-321">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-321">Save and close the file.</span></span>

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a><span data-ttu-id="96a94-322">Criar o roteiro que obterá os dados do Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="96a94-322">Create the route that will fetch the data from Microsoft Graph</span></span>

1. <span data-ttu-id="96a94-323">Abra o arquivo `app.js` na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="96a94-323">Open the file `app.js` in the root of the project.</span></span> <span data-ttu-id="96a94-324">Logo abaixo da rota para "/dialog.html", adicione a seguinte rota.</span><span class="sxs-lookup"><span data-stu-id="96a94-324">Just below the route for '/dialog.html', add the following route.</span></span> <span data-ttu-id="96a94-325">Esse roteiro é chamado pela função `makeGraphApiCall` que você criou em uma etapa anterior.</span><span class="sxs-lookup"><span data-stu-id="96a94-325">This route is called by the `makeGraphApiCall` function that you created in an earlier step.</span></span>

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. <span data-ttu-id="96a94-326">Substitua `TODO 15` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-326">Replace `TODO 15` with the following.</span></span> <span data-ttu-id="96a94-327">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-327">About this code, note:</span></span>

    - <span data-ttu-id="96a94-328">O chamador dessa rota, `makeGraphApiCall`, adicionou o token de acesso ao Microsoft Graph à solicitação HTTP como um cabeçalho chamado "access_token".</span><span class="sxs-lookup"><span data-stu-id="96a94-328">The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".</span></span>
    - <span data-ttu-id="96a94-329">A função `getGraphData` é definida no arquivo`msgraph-helper.js`.</span><span class="sxs-lookup"><span data-stu-id="96a94-329">The `getGraphData` function is defined in the `msgraph-helper.js` file.</span></span> <span data-ttu-id="96a94-330">(Essa não é a mesma função que a função do lado do cliente`getGraphData` definida no arquivo de `ssoAuthES6.js`.)</span><span class="sxs-lookup"><span data-stu-id="96a94-330">(This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)</span></span>
    - <span data-ttu-id="96a94-331">O último parâmetro, por `queryParamsSegment`, é codificado.</span><span class="sxs-lookup"><span data-stu-id="96a94-331">The last parameter, for `queryParamsSegment`, is hardcoded.</span></span> <span data-ttu-id="96a94-332">Se você reutilizar o código em um suplemento de produção e provenientes de qualquer parte do `queryParamsSegment` de entrada do usuário, certifique-se de que estão limpos para que não possam ser usados em um ataque de inserção de cabeçalho de resposta.</span><span class="sxs-lookup"><span data-stu-id="96a94-332">If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.</span></span>
    - <span data-ttu-id="96a94-333">O código minimiza os dados que devem ser provenientes do Microsoft Graph especificando apenas a propriedade de que precisamos ("nome") e somente os 10 primeiros nomes de pasta ou arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-333">The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.</span></span>

    ```javascript
    const graphToken = req.get('access_token');
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. <span data-ttu-id="96a94-334">Substitua `TODO 16` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="96a94-334">Replace `TODO 16` with the following.</span></span> <span data-ttu-id="96a94-335">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="96a94-335">About this code, note:</span></span>

    - <span data-ttu-id="96a94-336">Se o Microsoft Graph retornar um erro, como um token inválido ou expirado, haverá uma propriedade de código no conjunto de objetos retornados para um status HTTP (por exemplo, 401).</span><span class="sxs-lookup"><span data-stu-id="96a94-336">If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401).</span></span> <span data-ttu-id="96a94-337">O código retransmite o erro para o cliente.</span><span class="sxs-lookup"><span data-stu-id="96a94-337">The code relays the error to the client.</span></span> <span data-ttu-id="96a94-338">Ele será pego na callback `.fail` do `makeGraphApiCall`.</span><span class="sxs-lookup"><span data-stu-id="96a94-338">It will be caught in the `.fail` callback of `makeGraphApiCall`.</span></span>
    - <span data-ttu-id="96a94-339">Os dados do Microsoft Graph incluem metadados OData e eTags que o suplemento não precisa, portanto, o código cria uma nova matriz contendo somente os nomes de arquivos a serem enviados para o cliente.</span><span class="sxs-lookup"><span data-stu-id="96a94-339">Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.</span></span>

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. <span data-ttu-id="96a94-340">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="96a94-340">Save and close the file.</span></span>

## <a name="run-the-project"></a><span data-ttu-id="96a94-341">Executar o projeto</span><span class="sxs-lookup"><span data-stu-id="96a94-341">Run the project</span></span>

1. <span data-ttu-id="96a94-342">Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.</span><span class="sxs-lookup"><span data-stu-id="96a94-342">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="96a94-343">Abra um aviso de comando na raiz da pasta `\Begin`.</span><span class="sxs-lookup"><span data-stu-id="96a94-343">Open a command prompt in the root of the `\Begin` folder.</span></span> 

1. <span data-ttu-id="96a94-344">Execute o comando `npm start`.</span><span class="sxs-lookup"><span data-stu-id="96a94-344">Run the command `npm start`.</span></span> 

1. <span data-ttu-id="96a94-345">Você deve fazer o sideload do suplemento em um aplicativo do Office (Excel, Word ou PowerPoint) para testá-lo.</span><span class="sxs-lookup"><span data-stu-id="96a94-345">You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it.</span></span> <span data-ttu-id="96a94-346">As instruções dependem da plataforma.</span><span class="sxs-lookup"><span data-stu-id="96a94-346">The instructions depend on your platform.</span></span> <span data-ttu-id="96a94-347">Há links para instruções no [Fazer sideload de suplemento para teste](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span><span class="sxs-lookup"><span data-stu-id="96a94-347">There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span></span>

1. <span data-ttu-id="96a94-348">No aplicativo do Office, na faixa de opções **Home**, selecione o botão **Mostrar suplemento** no grupo**SSO Node.js** para abrir o suplemento do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="96a94-348">In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.</span></span>

1. <span data-ttu-id="96a94-349">Clique no botão **Definir Nome de Arquivos do One Drive**.</span><span class="sxs-lookup"><span data-stu-id="96a94-349">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="96a94-350">Se você estiver conectado ao Office com uma conta de educação ou de trabalho do Microsoft 365, ou uma conta da Microsoft, e o SSO estiver funcionando conforme o esperado, os 10 primeiros nomes de arquivos e pastas no OneDrive for Business serão inseridos no documento.</span><span class="sxs-lookup"><span data-stu-id="96a94-350">If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document.</span></span> <span data-ttu-id="96a94-351">Isso pode levar até 15 segundos pela primeira vez. Se você não estiver conectado ou se você estiver em um cenário que não tem suporte para SSO, ou se o SSO não estiver funcionando por nenhum motivo, você será solicitado a fazer logon.</span><span class="sxs-lookup"><span data-stu-id="96a94-351">(It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="96a94-352">Depois de entrar, os nomes de arquivos e pastas serão exibidos.</span><span class="sxs-lookup"><span data-stu-id="96a94-352">After you log in, the file and folder names appear.</span></span>

> [!NOTE]
> <span data-ttu-id="96a94-353">Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso.</span><span class="sxs-lookup"><span data-stu-id="96a94-353">If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so.</span></span> <span data-ttu-id="96a94-354">Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados.</span><span class="sxs-lookup"><span data-stu-id="96a94-354">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="96a94-355">Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter nomes de arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="96a94-355">To prevent this, be sure to *close all other Office applications* before you press **Get OneDrive File Names**.</span></span>
