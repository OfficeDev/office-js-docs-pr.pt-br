---
title: Criar um Suplemento do Office com ASP.NET que use logon único
description: Um guia passo a passo sobre como criar (ou converter) um suplemento do Office com um back-end do ASP.NET para usar o logon único (SSO).
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 591761c9831109cb6900cb418e83d3240ddad29d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292866"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a><span data-ttu-id="bdb3d-103">Criar um Suplemento do Office com ASP.NET que use logon único</span><span class="sxs-lookup"><span data-stu-id="bdb3d-103">Create an ASP.NET Office Add-in that uses single sign-on</span></span>

<span data-ttu-id="bdb3d-104">Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-104">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="bdb3d-105">Confira uma visão geral no artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-105">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="bdb3d-106">Este artigo orienta você durante o processo de habilitação do logon único (SSO) em um suplemento que é criado com o ASP.NET.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET.</span></span>

> [!NOTE]
> <span data-ttu-id="bdb3d-107">Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon único](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-107">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bdb3d-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="bdb3d-108">Prerequisites</span></span>

* <span data-ttu-id="bdb3d-109">Visual Studio 2019 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-109">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="bdb3d-110">Office Developer Tools</span><span class="sxs-lookup"><span data-stu-id="bdb3d-110">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="bdb3d-111">Pelo menos alguns arquivos e pastas armazenados no OneDrive for Business em sua assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-111">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="bdb3d-112">Uma assinatura do Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-112">A Microsoft Azure subscription.</span></span> <span data-ttu-id="bdb3d-113">Este suplemento requer o Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-113">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="bdb3d-114">O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-114">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="bdb3d-115">Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-115">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="bdb3d-116">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="bdb3d-116">Set up the starter project</span></span>

<span data-ttu-id="bdb3d-117">Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-117">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="bdb3d-118">Há duas versões do exemplo:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-118">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="bdb3d-p103">A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="bdb3d-122">A versão **Complete** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-122">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="bdb3d-123">Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Complete" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-123">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="bdb3d-124">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="bdb3d-124">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="bdb3d-125">Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-125">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="bdb3d-126">Entre com as credenciais de ***administrador*** em seu Microsoft 365 locação.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-126">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="bdb3d-127">Por exemplo, MeuNome@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-127">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="bdb3d-128">Selecione **Novo registro**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-128">Select **New registration**.</span></span> <span data-ttu-id="bdb3d-129">Na página **Registrar um aplicativo**, defina os valores da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-129">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="bdb3d-130">Defina **Nome** para `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-130">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="bdb3d-131">Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional (Qualquer diretório do Azure AD – Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-131">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="bdb3d-132">(Se você quiser que o suplemento possa ser usado somente por usuários no locatário em que você está os registrando, escolha **Contas somente neste diretório organizacional...**, mas execute algumas etapas adicionais.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-132">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="bdb3d-133">Confira **Configuração para locatário único** abaixo.)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-133">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="bdb3d-134">Na seção **URI de redirecionamento**, verifique se **Web** está selecionado no menu suspenso e defina o URI como ` https://localhost:44355/AzureADAuth/Authorize`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-134">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="bdb3d-135">Escolha **Registrar**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-135">Choose **Register**.</span></span>

1. <span data-ttu-id="bdb3d-136">Na página **Office-Add-in-ASPNET-SSO** , copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-136">On the **Office-Add-in-ASPNET-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="bdb3d-137">Use ambos os valores nos procedimentos posteriores.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-137">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bdb3d-138">Essa ID é o valor "Audience" quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-138">This ID is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="bdb3d-139">Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-139">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="bdb3d-140">Em **Gerenciar**, selecione **Certificados e segredos**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-140">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="bdb3d-141">Selecione o botão **Novo segredo do cliente**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-141">Select the **New client secret** button.</span></span> <span data-ttu-id="bdb3d-142">Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-142">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="bdb3d-143">*Copiar o valor de segredo do cliente imediatamente e salvá-lo com a ID de aplicativo* antes de prosseguir, pois ele será necessário em um procedimento posterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-143">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="bdb3d-144">Em **Gerenciar**, selecione **Expor uma API**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-144">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="bdb3d-145">Selecione o link **Definir** para gerar o URI da ID de Aplicativo no formato "api: / / $App ID GUID$", em que $App ID GUID$ é **ID do aplicativo (cliente)**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-145">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="bdb3d-146">Insira `localhost:44355/` (Observe a barra "/" anexada ao fim) após o `//` e antes do GUID.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-146">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="bdb3d-147">A ID inteira deve ter o formulário `api://localhost:44355/$App ID GUID$`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-147">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="bdb3d-148">Marque **Salvar** na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-148">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="bdb3d-149">Selecione o botão **Adicionar um escopo**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-149">Select the **Add a scope** button.</span></span> <span data-ttu-id="bdb3d-150">No painel que se abre, insira `access_as_user` como o **Nome de escopo**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-150">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="bdb3d-151">Definir **Quem pode consentir?** aos **Administradores e usuários**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-151">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="bdb3d-152">Preencha os campos para configurar os prompts de consentimento de usuário e administrador com os valores que são apropriados para o `access_as_user` escopo, o que permite que o aplicativo cliente do Office Use as APIs Web do seu suplemento com os mesmos direitos que o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-152">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="bdb3d-153">Sugestões:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-153">Suggestions:</span></span>

    - <span data-ttu-id="bdb3d-154">**Título de autorização de administrador:** Office pode funcionar como o usuário.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-154">**Admin consent title**: Office can act as the user.</span></span>
    - <span data-ttu-id="bdb3d-155">**Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-155">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="bdb3d-156">**Título de autorização de usuário:** O Office pode funcionar como se fosse você.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-156">**User consent title**: Office can act as you.</span></span>
    - <span data-ttu-id="bdb3d-157">**Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que você possui.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-157">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="bdb3d-158">Verifique se o **Estado** está definido como **Habilitado**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-158">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="bdb3d-159">Selecione **Adicionar escopo**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-159">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="bdb3d-160">A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-160">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="bdb3d-161">Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-161">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="bdb3d-162">Cada uma das seguintes IDs precisa ser pré-autorizada.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-162">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="bdb3d-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="bdb3d-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="bdb3d-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="bdb3d-166">`08e18876-6177-487e-b8b5-cf950c1e598c`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-166">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
    - <span data-ttu-id="bdb3d-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="bdb3d-168">Para cada ID, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-168">For each ID, take these steps:</span></span>

    <span data-ttu-id="bdb3d-169">a.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-169">a.</span></span> <span data-ttu-id="bdb3d-170">Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o ID do cliente para o respectivo GUID e marque a caixa `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-170">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="bdb3d-171">b.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-171">b.</span></span> <span data-ttu-id="bdb3d-172">Selecione **Adicionar aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-172">Select **Add application**.</span></span>

1. <span data-ttu-id="bdb3d-173">Em **Gerenciar**, selecione **Permissões para API** e selecione **Adicionar uma permissão**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-173">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="bdb3d-174">No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-174">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="bdb3d-175">Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-175">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="bdb3d-176">Selecione estas opções.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-176">Select the following.</span></span> <span data-ttu-id="bdb3d-177">Somente o primeiro é realmente necessário para o seu suplemento; Mas a `profile` permissão é necessária para que o aplicativo do Office obtenha um token para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-177">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office application to get a token to your add-in web application.</span></span> <span data-ttu-id="bdb3d-178">(Somente Files.Read.All e o perfil são, de fato, necessários para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-178">(Only Files.Read.All and profile are actually needed by the add-in.</span></span> <span data-ttu-id="bdb3d-179">Solicite os outros dois porque a biblioteca MSAL.NET exige.)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-179">You must request the other two because the MSAL.NET library requires them.)</span></span>

    * <span data-ttu-id="bdb3d-180">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="bdb3d-180">Files.Read.All</span></span>
    * <span data-ttu-id="bdb3d-181">offline_access</span><span class="sxs-lookup"><span data-stu-id="bdb3d-181">offline_access</span></span>
    * <span data-ttu-id="bdb3d-182">openid</span><span class="sxs-lookup"><span data-stu-id="bdb3d-182">openid</span></span>
    * <span data-ttu-id="bdb3d-183">perfil</span><span class="sxs-lookup"><span data-stu-id="bdb3d-183">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="bdb3d-184">A permissão `User.Read` pode já estar listada por padrão.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-184">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="bdb3d-185">É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-185">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="bdb3d-186">Marque a caixa de seleção para cada permissão conforme elas forem exibidas.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-186">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="bdb3d-187">Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-187">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="bdb3d-188">Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Aceitar** para a confirmação exibida.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-188">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bdb3d-189">Depois de escolher **Conceder consentimento de administrador para [nome do locatário]**, você verá uma mensagem solicitando que você tente novamente alguns minutos depois, para que a solicitação de consentimento possa ser construída.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-189">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="bdb3d-190">Em caso afirmativo, você pode começar a trabalhar na próxima seção, ***mas não se esqueça de voltar para o portal e pressionar este botão***!</span><span class="sxs-lookup"><span data-stu-id="bdb3d-190">If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="bdb3d-191">Configurar a solução</span><span class="sxs-lookup"><span data-stu-id="bdb3d-191">Configure the solution</span></span>

1. <span data-ttu-id="bdb3d-192">Na raiz da pasta **Before** (antes), abra o arquivo de solução (.sln) no **Visual Studio**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-192">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="bdb3d-193">Clique com o botão direito do mouse no nó superior no **Gerenciador de Soluções** (no nó Solução, não em qualquer um dos nós do projeto) e selecione **Configurar projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-193">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="bdb3d-194">Em **Propriedades Comuns**, selecione **Projeto de Inicialização** e **Vários projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-194">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="bdb3d-195">Verifique se a **Ação** para ambos os projetos está definida como **Iniciar** e se o projeto terminado em "...WebAPI" está listado primeiro.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-195">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="bdb3d-196">Feche a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-196">Close the dialog.</span></span>

1. <span data-ttu-id="bdb3d-197">No **Gerenciador de Soluções**, selecione (não clique com o botão direito) o projeto **Office-Add-in-Microsoft-Graph-ASPNETWebAPI**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-197">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-Microsoft-Graph-ASPNETWebAPI** project.</span></span> <span data-ttu-id="bdb3d-198">O painel **Propriedades** é exibido.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-198">The **Properties** pane opens.</span></span> <span data-ttu-id="bdb3d-199">Verifique se **SSL Habilitado** é **Verdadeiro**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-199">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="bdb3d-200">Verifique se a **URL do SSL** é `http://localhost:44355/`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-200">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="bdb3d-201">Em "Web.config", use os valores copiados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-201">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="bdb3d-202">Defina **ida:ClientID** e **ida:Audience** para sua **ID do aplicativo (cliente)** e defina **ida:Password** para a senha de cliente.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-202">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bdb3d-203">A **ID do aplicativo (cliente)** é o valor de "audiência" quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-203">The **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="bdb3d-204">Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-204">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="bdb3d-205">Se você não tiver escolhido "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, salve e feche o Web.config. Caso contrário, salve, mas deixe-o aberto.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-205">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="bdb3d-206">Ainda no **Gerenciador de Soluções**, escolha o projeto **Office-Add-in-Microsoft-Graph-ASPNET** e abra o arquivo de manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml" e role até a parte inferior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-206">Still in **Solution Explorer**, choose the **Office-Add-in-Microsoft-Graph-ASPNET** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="bdb3d-207">Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-207">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="bdb3d-208">Substitua o espaço reservado "{$application_GUID here$}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-208">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="bdb3d-209">Os sinais "$" não fazem parte da ID, portanto não os inclua.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-209">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="bdb3d-210">Essa é a mesma ID usada para a ClientID e a Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-210">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

  > [!NOTE]
  > <span data-ttu-id="bdb3d-211">O valor **Recurso** é o**URI da ID de aplicativo** que você definiu quando registrou o suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-211">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="bdb3d-212">A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-212">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="bdb3d-213">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-213">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="bdb3d-214">Configuração para locatário único</span><span class="sxs-lookup"><span data-stu-id="bdb3d-214">Setup for single-tenant</span></span>

<span data-ttu-id="bdb3d-215">Se você escolher "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, você execute estas etapas adicionais de configuração:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-215">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="bdb3d-216">Volte para o Portal do Azure e abra a lâmina **Visão geral** do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-216">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="bdb3d-217">Copie a **ID de diretório (locatário)**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-217">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="bdb3d-218">Em Web.config, substitua o "comum" no valor de **ida:Authority** pela GUID copiada na etapa anterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-218">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="bdb3d-219">Ao terminar, o valor deverá ser similar a este: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-219">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="bdb3d-220">Salve e feche o web.config.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-220">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="bdb3d-221">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="bdb3d-221">Code the client side</span></span>

1. <span data-ttu-id="bdb3d-222">Abra o arquivo HomeES6.js na pasta **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-222">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="bdb3d-223">Ele já apresenta alguns códigos:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-223">It already has some code in it:</span></span>

    * <span data-ttu-id="bdb3d-224">Um polyfill que atribui o objeto Office.Promise ao objeto de janela global, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-224">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="bdb3d-225">(Para obter mais detalhes, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-225">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="bdb3d-226">Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-226">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="bdb3d-227">Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-227">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="bdb3d-228">Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-228">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="bdb3d-229">O código implementa o sistema de autorização de fallback que o suplemento usará em situações em que o SSO não é compatível ou gera um erro.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-229">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="bdb3d-230">Abaixo da atribuição a `Office.initialize`, adicione o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-230">Below the assignment to `Office.initialize`, add the code below.</span></span> <span data-ttu-id="bdb3d-231">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-231">Note the following about this code:</span></span>

    * <span data-ttu-id="bdb3d-232">O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-232">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="bdb3d-233">A variável de contador `retryGetAccessToken` é usada para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-233">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="bdb3d-234">A função `getGraphData` é definida com a palavra-chave ES6 `async`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-234">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="bdb3d-235">Usar a sintaxe ES6 facilita o uso da API de SSO em Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-235">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="bdb3d-236">Esse é o único arquivo na solução que usará a sintaxe sem suporte do Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-236">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="bdb3d-237">Colocamos "ES6" no nome do arquivo como um lembrete.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-237">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="bdb3d-238">A solução usa o transcompilador de tsc para transcompilar esse arquivo em ES5, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-238">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="bdb3d-239">(Veja o arquivo tsconfig.json na raiz do projeto.)</span><span class="sxs-lookup"><span data-stu-id="bdb3d-239">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="bdb3d-240">Abaixo da função `getGraphData`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-240">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="bdb3d-241">Observe que você criará a função `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-241">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graph and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. <span data-ttu-id="bdb3d-242">Substitua `TODO 1` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-242">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="bdb3d-243">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-243">About this code, note:</span></span>

    * <span data-ttu-id="bdb3d-244">`getAccessToken` instrui o Office a obter um token de bootstrap do Azure AD e retornar ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-244">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="bdb3d-245">`allowSignInPrompt` indica ao Office para solicitar que o usuário entre caso ele ainda não esteja conectado ao Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-245">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="bdb3d-246">`allowConsentPrompt` informa ao Office para solicitar que o usuário concorde em permitir que o suplemento acesse o perfil AAD do usuário, se o consentimento ainda não tiver sido concedido.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-246">`allowConsentPrompt` tells Office to prompt the user to consent to letting the add-in access the user's AAD profile, if consent has not already been granted.</span></span> <span data-ttu-id="bdb3d-247">(O prompt resultante *não* permite que o usuário concorde com os escopos do Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-247">(The resulting prompt does *not* allow the user to consent to any Microsoft Graph scopes.)</span></span>
    * <span data-ttu-id="bdb3d-248">`forMSGraphAccess` instrui o Office que o suplemento pretende trocar o token de bootstrap por um token de acesso ao Micrsoft Graph, em vez de apenas usar o token de bootstrap como um token de ID.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-248">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="bdb3d-249">A configuração dessa opção dá ao Office a oportunidade de cancelar o processo de obtenção do token de bootstrap (e retornar o código de erro 13012) se o administrador de locatários do usuário não tiver concedido consentimento para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-249">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="bdb3d-250">O código do lado do cliente do suplemento pode responder ao 13012 por meio da ramificação para um sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-250">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="bdb3d-251">Se o `forMSGraphAccess` não for usado e o administrador não tiver concedido o consentimento, o token de inicialização será retornado, mas a tentativa de troca com o fluxo em nome de para resultaria em um erro.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-251">If the `forMSGraphAccess` is not used and the admin has not granted consent, the bootstrap token is returned, but the attempt to exchange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="bdb3d-252">Portanto, a opção `forMSGraphAccess` permite ao suplemento ramificar para o sistema de fallback rapidamente.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-252">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="bdb3d-253">Você criará a função `getData` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-253">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="bdb3d-254">O parâmetro `/api/values` é a URL de um controlador do lado do servidor que fará a troca de tokens e usará o token de acesso recebido para fazer a chamada para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-254">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="bdb3d-255">Abaixo da função `getGraphData`, adicione o seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-255">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="bdb3d-256">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-256">About this code, note:</span></span>

    * <span data-ttu-id="bdb3d-257">Ele é usado pelos sistemas de autorização de fallback e SSO.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-257">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="bdb3d-258">O parâmetro `relativeUrl` é um controlador do lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-258">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="bdb3d-259">O parâmetro `accessToken` pode ser um token de bootstrap ou um token de acesso completo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-259">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="bdb3d-260">O `writeFileNamesToOfficeDocument` já faz parte do projeto.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-260">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="bdb3d-261">Você criará a função `handleServerSideErrors` em uma última etapa.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-261">You create the `handleServerSideErrors` function in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a><span data-ttu-id="bdb3d-262">Tratar erros do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="bdb3d-262">Handle client-side errors</span></span>

1. <span data-ttu-id="bdb3d-263">Abaixo da função `getData`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-263">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="bdb3d-264">Observe que `error.code` é um número, normalmente no intervalo 13xxx.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-264">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="bdb3d-265">Substitua `TODO 2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-265">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="bdb3d-266">Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. <span data-ttu-id="bdb3d-267">Substitua `TODO 3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-267">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="bdb3d-268">Para todos os outros erros, o suplemento ramificará para o sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-268">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="bdb3d-269">Para obter mais informações sobre esses erros, confira [solucionar problemas de SSO nos suplementos do Office](troubleshoot-sso-in-office-add-ins.md). Neste suplemento, o sistema de fallback abre uma caixa de diálogo que requer que o usuário entre, mesmo que o usuário já esteja.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-269">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="bdb3d-270">Resolver erros do lado do servidor</span><span class="sxs-lookup"><span data-stu-id="bdb3d-270">Handle server-side errors</span></span>

1. <span data-ttu-id="bdb3d-271">Abaixo da função `handleClientSideErrors`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-271">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="bdb3d-272">Substitua `TODO 4` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-272">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="bdb3d-273">Sobre esse código, observe que as classes de erro ASP.NET foram criadas antes de haver algo como a MFA.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-273">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="bdb3d-274">Como um efeito colateral de como a lógica do lado do servidor lida com as solicitações de um segundo fator de autenticação, o erro do lado do servidor enviado para o cliente tem uma propriedade **Message**, mas não uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-274">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="bdb3d-275">Mas todos os outros erros terão uma propriedade **ExceptionMessage**, para que o código do cliente precise analisar a resposta para ambos.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-275">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="bdb3d-276">Uma ou outra variável será indefinida.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-276">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="bdb3d-277">Substitua `TODO 5` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-277">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="bdb3d-278">Quando o Microsoft Graph requer uma forma adicional de autenticação, ele envia o erro AADSTS50076.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-278">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="bdb3d-279">Isso inclui informações sobre os requisitos adicionais na propriedade **Message.Claims**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-279">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="bdb3d-280">Para lidar com isso, o código faz uma segunda tentativa de obter o token de bootstrap, mas, desta vez, ele inclui a solicitação de um fator adicional, como o valor da opção `authChallenge`, que informa ao Azure AD a solicitar todos os formulários de autenticação necessários.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-280">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. <span data-ttu-id="bdb3d-281">Substitua `TODO 6` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-281">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="bdb3d-282">Substitua `TODO 7` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-282">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="bdb3d-283">Observe que, em raras ocasiões, o token de bootstrap fica não vencido quando o Office o valida, mas vence no momento em que ele é enviado ao Azure AD para o Exchange.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-283">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="bdb3d-284">O Azure AD responderá com o erro AADSTS500133.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-284">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="bdb3d-285">Quando isso acontece, o código recupera a API de SSO (mas não mais de uma vez).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-285">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="bdb3d-286">Desta vez, o Office retorna um novo token de bootstrap não vencido.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-286">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="bdb3d-287">Substitua `TODO 8` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-287">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="bdb3d-288">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-288">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="bdb3d-289">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="bdb3d-289">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="bdb3d-290">Configurar o middleware OWIN</span><span class="sxs-lookup"><span data-stu-id="bdb3d-290">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="bdb3d-291">Abra o arquivo Startup.cs na raiz do projeto **Office-Add-in-ASPNET-SSO-WebAPI** e adicione o seguinte método à classe **Inicialização**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-291">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="bdb3d-292">Observe que você criará o método `ConfigureAuth` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-292">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="bdb3d-293">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-293">Save and close the file.</span></span>

1. <span data-ttu-id="bdb3d-294">Clique com botão direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-294">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="bdb3d-295">Na caixa de diálogo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-295">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="bdb3d-296">Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-296">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="bdb3d-297">Verifique se todas as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-297">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="bdb3d-p149">Adicione a palavra-chave `partial` à declaração da classe `Startup`, se ainda não estiver lá. A linha deverá ser assim:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p149">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="bdb3d-p150">Adicione o método a seguir à classe `Startup`. Este método especifica como o middleware OWIN validará os tokens de acesso que são transmitidos a ele do método `getData` no arquivo Home.js do lado do cliente. O processo de autorização é disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` é chamado.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p150">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="bdb3d-303">Substitua o `TODO 1` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-303">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="bdb3d-304">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-304">Note about this code:</span></span>

    * <span data-ttu-id="bdb3d-305">O código instrui o OWIN a garantir que a audiência especificada no token de Bootstrap proveniente do aplicativo do Office deve corresponder ao valor especificado na web.config.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-305">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office application must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="bdb3d-306">As contas da Microsoft têm um GUID emissor que é diferente de qualquer GUID de locatário organizacional, portanto, para dar suporte a ambos os tipos de contas, não validamos o emissor.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-306">Microsoft accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="bdb3d-307">`SaveSigninToken`A configuração para `true` faz com que o OWIN salve o token de inicialização bruto do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-307">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office application.</span></span> <span data-ttu-id="bdb3d-308">O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on-behalf-of".</span><span class="sxs-lookup"><span data-stu-id="bdb3d-308">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="bdb3d-309">Os escopos não são validados pelo middleware OWIN.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-309">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="bdb3d-310">Os escopos do token de bootstrap, que devem conter `access_as_user`, são validados no controlador.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-310">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="bdb3d-311">Substitua `TODO 2` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-311">Replace `TODO 2` with the following.</span></span> <span data-ttu-id="bdb3d-312">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-312">Note about this code:</span></span>

    * <span data-ttu-id="bdb3d-313">O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-313">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="bdb3d-314">A URL que é passada para o método é onde o middleware OWIN Obtém instruções para obter a chave necessária para verificar a assinatura no token de inicialização recebido do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-314">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office application.</span></span> <span data-ttu-id="bdb3d-315">O segmento de Autoridade da URL vem do Web.config. Ele é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-315">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="bdb3d-316">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-316">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="bdb3d-317">Criar o controlador /api/values</span><span class="sxs-lookup"><span data-stu-id="bdb3d-317">Create the /api/values controller</span></span>

1. <span data-ttu-id="bdb3d-318">Abra o arquivo **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-318">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="bdb3d-319">Esse controlador é usado quando o sistema SSO obtém um token de bootstrap com êxito.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-319">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="bdb3d-320">Ele não é usado como parte do sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-320">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="bdb3d-321">Esse sistema usou o AzureADAuthController que foi criado para você.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-321">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="bdb3d-322">Verifique se as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-322">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. <span data-ttu-id="bdb3d-p157">Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização configurado no último procedimento sempre que um método controlador for chamado. Apenas os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p157">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="bdb3d-326">Adicione o método a seguir ao `ValuesController`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-326">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="bdb3d-327">Observe que é o valor de retorno é `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um método `GET api/values`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-327">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="bdb3d-328">Este é o efeito colateral deste fato que a lógica de autorização do OAuth deve estar no controlador, em fez de em um filtro ASP.NET.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-328">This is a side effect of that fact that the OAuth  authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="bdb3d-329">Algumas condições de erro na lógica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-329">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. <span data-ttu-id="bdb3d-330">Substitua `TODO1` pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-330">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="bdb3d-331">Observe que o segundo parâmetro do método `SendErrorToClient` é um objeto **Exception**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-331">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="bdb3d-332">Nesse caso, o código passa `null` porque incluir o objeto **Exception** bloqueia a inclusão da propriedade **Message** na resposta HTTP que é gerada.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-332">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="bdb3d-333">Substitua `TODO 2` pelo seguinte código para montar todas as informações necessárias para obter um token do Microsoft Graph usando o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="bdb3d-333">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="bdb3d-334">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-334">About this code, note:</span></span>

    * <span data-ttu-id="bdb3d-335">O suplemento não está mais executando a função de um recurso (ou público) ao qual o aplicativo do Office e o usuário precisam de acesso.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-335">Your add-in is no longer playing the role of a resource (or audience) to which the Office application and user need access.</span></span> <span data-ttu-id="bdb3d-336">Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-336">Now it is itself a client that needs access to Microsoft Graph.</span></span> <span data-ttu-id="bdb3d-337">`ConfidentialClientApplication` é o objeto "client context" da MSAL.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-337">`ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="bdb3d-338">A partir da MSAL.NET 3.x.x, o `bootstrapContext` é apenas o token de bootstrap em si.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-338">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="bdb3d-339">A Autoridade vem do Web.config. Ela é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-339">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="bdb3d-340">A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lança um erro se o código solicitá-los de forma redundante.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-340">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="bdb3d-341">Ele também lançará um erro se o código solicitar `profile` , que realmente é usado apenas quando o aplicativo cliente do Office Obtém o token para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-341">It will also throw an error if your code requests `profile`, which is really only used when the Office client application gets the token to your add-in's web application.</span></span> <span data-ttu-id="bdb3d-342">Então, apenas `Files.Read.All` é explicitamente solicitado.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-342">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri("https://localhost:44355")
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. <span data-ttu-id="bdb3d-343">Substitua `TODO 3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-343">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="bdb3d-344">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-344">Note about this code:</span></span>

    * <span data-ttu-id="bdb3d-345">O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-345">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="bdb3d-346">Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-346">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="bdb3d-347">Quaisquer exceções que não forem do tipo `MsalServiceException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente como mensagens `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-347">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="bdb3d-348">Substitua `TODO 3a` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-348">Replace `TODO 3a` with the following code.</span></span> <span data-ttu-id="bdb3d-349">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-349">About this code, note:</span></span>

    * <span data-ttu-id="bdb3d-350">Se a autenticação multifator for exigida pelo recurso Microsoft Graph e o usuário ainda não a tiver fornecido, o Azure AD retornará "400 Bad Request" com o erro `AADSTS50076` e uma propriedade **Declarações**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-350">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="bdb3d-351">O MSAL exibe **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informações.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-351">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="bdb3d-352">O valor da propriedade **Claims** deve ser passado para o cliente que deve passá-lo para o aplicativo do Office, que o incluirá em uma solicitação para um novo token de inicialização.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-352">The **Claims** property value must be passed to the client which should pass it to the Office application, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="bdb3d-353">O Azure AD solicitará ao usuário todas as formas de autenticação necessárias.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-353">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="bdb3d-p168">As APIs que criam respostas HTTP a partir de exceções não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto de resposta. É necessário criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a criação de uma propriedade **ExceptionMessage**, portanto, a única maneira de obter a ID de erro `AADSTS50076` para o cliente é adicioná-la à **Message** personalizada. O JavaScript no cliente precisará descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p168">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="bdb3d-358">A mensagem personalizada é formatada como JSON para que o JavaScript do cliente possa analisá-la com métodos de objeto `JSON` JavaScript conhecidos.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-358">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="bdb3d-359">Substitua `TODO 3b` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-359">Replace `TODO 3b` with the following code.</span></span> <span data-ttu-id="bdb3d-360">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bdb3d-360">About this code, note:</span></span>

    * <span data-ttu-id="bdb3d-361">Se a chamada para o Azure AD contiver pelo menos um escopo (permissão) que não tenha sido consentido pelo usuário ou por um administrador de locatários (ou se o consentimento foi revogado), o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS65001`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-361">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="bdb3d-362">O MSAL exibe um **MsalUiRequiredException** com essas informações.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-362">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="bdb3d-363">Se a chamada para o Azure AD contiver pelo menos um escopo que Azure AD não reconhece, o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS70011`.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-363">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="bdb3d-364">O MSAL exibe um **MsalUiRequiredException** com essas informações.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-364">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="bdb3d-365">A descrição completa é incluída porque 70011 é retornado em outras condições e ele deverá ser processado neste suplemento somente quando significar que há um escopo inválido.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-365">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="bdb3d-p172">O objeto **MsalUiRequiredException** é passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informações de erro seja incluída na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-p172">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="bdb3d-368">Substitua `TODO 3c` pelo seguinte código para lidar com todas as outras **MsalServiceException**s.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-368">Replace `TODO 3c` with the following code to handle all other **MsalServiceException**s.</span></span> <span data-ttu-id="bdb3d-369">Conforme observado anteriormente,</span><span class="sxs-lookup"><span data-stu-id="bdb3d-369">As noted earlier,</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="bdb3d-370">Substitua `TODO 4` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-370">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="bdb3d-371">O método `GraphApiHelper.GetOneDriveFileNames`, que foi criado para você, faz a solicitação de dados ao Microsoft Graph e inclui o token de acesso.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-371">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="bdb3d-372">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-372">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="bdb3d-373">Executar a solução</span><span class="sxs-lookup"><span data-stu-id="bdb3d-373">Run the solution</span></span>

1. <span data-ttu-id="bdb3d-374">Abra o arquivo de solução do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-374">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="bdb3d-375">No menu **Build**, selecione **Solução Limpa**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-375">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="bdb3d-376">Quando terminar, abra o menu **Build** novamente e selecione **Solução de Build**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-376">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="bdb3d-377">No **Gerenciador de Soluções**, selecione o nó de projeto **Office-Add-in-ASPNET-SSO** (não o nó da solução principal e não o projeto cujo nome termina em "WebAPI").</span><span class="sxs-lookup"><span data-stu-id="bdb3d-377">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="bdb3d-378">No painel **Propriedades**, abra o menu suspenso **Iniciar documento** e escolha uma das três opções (Excel, Word ou PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="bdb3d-378">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![Escolha o aplicativo cliente do Office desejado: Excel, PowerPoint ou Word](../images/SelectHost.JPG)

1. <span data-ttu-id="bdb3d-380">Pressione F5.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-380">Press F5.</span></span>
1. <span data-ttu-id="bdb3d-381">No aplicativo do Office, na faixa de opções **Home**, selecione **Mostrar suplemento** no grupo **SSO ASP.NET** para abrir o suplemento do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-381">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="bdb3d-382">Clique no botão **Definir Nome de Arquivos do One Drive**.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-382">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="bdb3d-383">Se você estiver conectado ao Office com uma conta de educação ou de trabalho do Microsoft 365, ou uma conta da Microsoft, e o SSO estiver funcionando conforme o esperado, os 10 primeiros nomes de arquivos e pastas no OneDrive for Business são exibidos no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-383">If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="bdb3d-384">Se você não estiver conectado ou se você estiver em um cenário que não tem suporte para SSO, ou se o SSO não estiver funcionando por nenhum motivo, você será solicitado a fazer logon.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-384">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="bdb3d-385">Depois de entrar, os nomes de arquivos e pastas serão exibidos.</span><span class="sxs-lookup"><span data-stu-id="bdb3d-385">After you log in, the file and folder names appear.</span></span>
