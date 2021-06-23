---
title: Criar um Suplemento do Office com ASP.NET que use logon único
description: Um guia passo a passo sobre como criar (ou converter) um Office com um back-in ASP.NET para usar o SSO (sign-on único).
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 35e4dcef6d99d5bd3ca204b08a017679684ec2ba
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076452"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a><span data-ttu-id="6a781-103">Criar um Suplemento do Office com ASP.NET que use logon único</span><span class="sxs-lookup"><span data-stu-id="6a781-103">Create an ASP.NET Office Add-in that uses single sign-on</span></span>

<span data-ttu-id="6a781-104">Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="6a781-104">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="6a781-105">Confira uma visão geral no artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6a781-105">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="6a781-106">Este artigo orienta você sobre o processo de habilitação de SSO (login único) em um complemento criado com ASP.NET.</span><span class="sxs-lookup"><span data-stu-id="6a781-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET.</span></span>

> [!NOTE]
> <span data-ttu-id="6a781-107">Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon único](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="6a781-107">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6a781-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6a781-108">Prerequisites</span></span>

* <span data-ttu-id="6a781-109">Visual Studio 2019 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-109">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="6a781-110">Office Developer Tools</span><span class="sxs-lookup"><span data-stu-id="6a781-110">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="6a781-111">Pelo menos alguns arquivos e pastas armazenados em OneDrive for Business em sua assinatura Microsoft 365 assinatura.</span><span class="sxs-lookup"><span data-stu-id="6a781-111">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="6a781-112">Uma assinatura do Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="6a781-112">A Microsoft Azure subscription.</span></span> <span data-ttu-id="6a781-113">Este suplemento requer o Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="6a781-113">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="6a781-114">O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização.</span><span class="sxs-lookup"><span data-stu-id="6a781-114">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="6a781-115">Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="6a781-115">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="6a781-116">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="6a781-116">Set up the starter project</span></span>

<span data-ttu-id="6a781-117">Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="6a781-117">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="6a781-118">Há duas versões do exemplo:</span><span class="sxs-lookup"><span data-stu-id="6a781-118">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="6a781-p103">A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.</span><span class="sxs-lookup"><span data-stu-id="6a781-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="6a781-122">A versão **Complete** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo.</span><span class="sxs-lookup"><span data-stu-id="6a781-122">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="6a781-123">Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Complete" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="6a781-123">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="6a781-124">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="6a781-124">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="6a781-125">Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6a781-125">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="6a781-126">Entre com as ***credenciais de*** administrador no seu Microsoft 365 de adoção.</span><span class="sxs-lookup"><span data-stu-id="6a781-126">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="6a781-127">Por exemplo, MeuNome@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="6a781-127">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="6a781-128">Selecione **Novo registro**.</span><span class="sxs-lookup"><span data-stu-id="6a781-128">Select **New registration**.</span></span> <span data-ttu-id="6a781-129">Na página **Registrar um aplicativo**, defina os valores da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="6a781-129">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="6a781-130">Defina **Nome** para `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="6a781-130">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="6a781-131">Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional (Qualquer diretório do Azure AD – Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**.</span><span class="sxs-lookup"><span data-stu-id="6a781-131">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="6a781-132">(Se você quiser que o suplemento possa ser usado somente por usuários no locatário em que você está os registrando, escolha **Contas somente neste diretório organizacional...**, mas execute algumas etapas adicionais.</span><span class="sxs-lookup"><span data-stu-id="6a781-132">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="6a781-133">Confira **Configuração para locatário único** abaixo.)</span><span class="sxs-lookup"><span data-stu-id="6a781-133">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="6a781-134">Na seção **URI de redirecionamento**, verifique se **Web** está selecionado no menu suspenso e defina o URI como ` https://localhost:44355/AzureADAuth/Authorize`.</span><span class="sxs-lookup"><span data-stu-id="6a781-134">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="6a781-135">Escolha **Registrar**.</span><span class="sxs-lookup"><span data-stu-id="6a781-135">Choose **Register**.</span></span>

1. <span data-ttu-id="6a781-136">Na página **Office-Add-in-ASPNET-SSO,** copie e salve o valor da ID do aplicativo **(cliente).**</span><span class="sxs-lookup"><span data-stu-id="6a781-136">On the **Office-Add-in-ASPNET-SSO** page, copy and save the value for the **Application (client) ID**.</span></span> <span data-ttu-id="6a781-137">Você precisará dele em procedimentos posteriores.</span><span class="sxs-lookup"><span data-stu-id="6a781-137">You'll need it in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a781-138">Essa ID de Aplicativo **(cliente)** é o valor "público" quando outros aplicativos, como o aplicativo cliente Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6a781-138">This **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="6a781-139">Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6a781-139">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="6a781-140">Em **Gerenciar**, selecione **Certificados e segredos**.</span><span class="sxs-lookup"><span data-stu-id="6a781-140">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="6a781-141">Selecione o botão **Novo segredo do cliente**.</span><span class="sxs-lookup"><span data-stu-id="6a781-141">Select the **New client secret** button.</span></span> <span data-ttu-id="6a781-142">Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="6a781-142">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="6a781-143">Copie o valor do segredo do cliente (não a *ID Secreta)* imediatamente e salve-o com a ID do aplicativo antes de prosseguir, pois você precisará dele em um procedimento posterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-143">*Copy the client secret value (not the Secret ID) immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="6a781-144">Em **Gerenciar**, selecione **Expor uma API**.</span><span class="sxs-lookup"><span data-stu-id="6a781-144">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="6a781-145">Selecione o link **Definir** para gerar o URI da ID de Aplicativo no formato "api: / / $App ID GUID$", em que $App ID GUID$ é **ID do aplicativo (cliente)**.</span><span class="sxs-lookup"><span data-stu-id="6a781-145">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="6a781-146">Insira `localhost:44355/` (Observe a barra "/" anexada ao fim) após o `//` e antes do GUID.</span><span class="sxs-lookup"><span data-stu-id="6a781-146">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="6a781-147">A ID inteira deve ter o formulário `api://localhost:44355/$App ID GUID$`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="6a781-147">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="6a781-148">Marque **Salvar** na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="6a781-148">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="6a781-149">Selecione o botão **Adicionar um escopo**.</span><span class="sxs-lookup"><span data-stu-id="6a781-149">Select the **Add a scope** button.</span></span> <span data-ttu-id="6a781-150">No painel que se abre, insira `access_as_user` como o **Nome de escopo**.</span><span class="sxs-lookup"><span data-stu-id="6a781-150">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="6a781-151">Definir **Quem pode consentir?** aos **Administradores e usuários**.</span><span class="sxs-lookup"><span data-stu-id="6a781-151">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="6a781-152">Preencha os campos para configurar os prompts de consentimento do administrador e do usuário com valores apropriados para o escopo que permite que o aplicativo cliente Office use as APIs da Web do seu complemento com os mesmos direitos do usuário `access_as_user` atual.</span><span class="sxs-lookup"><span data-stu-id="6a781-152">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="6a781-153">Sugestões:</span><span class="sxs-lookup"><span data-stu-id="6a781-153">Suggestions:</span></span>

    * <span data-ttu-id="6a781-154">**Nome de exibição** de consentimento do administrador : Office pode atuar como o usuário.</span><span class="sxs-lookup"><span data-stu-id="6a781-154">**Admin consent display name**: Office can act as the user.</span></span>
    * <span data-ttu-id="6a781-155">**Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="6a781-155">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    * <span data-ttu-id="6a781-156">**Nome de exibição** de consentimento do usuário : Office pode atuar como você.</span><span class="sxs-lookup"><span data-stu-id="6a781-156">**User consent display name**: Office can act as you.</span></span>
    * <span data-ttu-id="6a781-157">**Descrição do** consentimento do usuário : Office para chamar as APIs da Web do complemento com os mesmos direitos que você tem.</span><span class="sxs-lookup"><span data-stu-id="6a781-157">**User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="6a781-158">Verifique se o **Estado** está definido como **Habilitado**.</span><span class="sxs-lookup"><span data-stu-id="6a781-158">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="6a781-159">Selecione **Adicionar escopo**.</span><span class="sxs-lookup"><span data-stu-id="6a781-159">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a781-160">A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="6a781-160">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="6a781-161">Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-161">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="6a781-162">Cada uma das seguintes IDs precisa ser pré-autorizada.</span><span class="sxs-lookup"><span data-stu-id="6a781-162">Each of the following IDs needs to be pre-authorized.</span></span>

    * <span data-ttu-id="6a781-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="6a781-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="6a781-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="6a781-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    * <span data-ttu-id="6a781-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="6a781-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    * <span data-ttu-id="6a781-166">`08e18876-6177-487e-b8b5-cf950c1e598c`(Office na Web)</span><span class="sxs-lookup"><span data-stu-id="6a781-166">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
    * <span data-ttu-id="6a781-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)</span><span class="sxs-lookup"><span data-stu-id="6a781-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="6a781-168">Para cada ID, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="6a781-168">For each ID, take these steps:</span></span>

    <span data-ttu-id="6a781-169">a.</span><span class="sxs-lookup"><span data-stu-id="6a781-169">a.</span></span> <span data-ttu-id="6a781-170">Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o ID do cliente para o respectivo GUID e marque a caixa `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="6a781-170">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="6a781-171">b.</span><span class="sxs-lookup"><span data-stu-id="6a781-171">b.</span></span> <span data-ttu-id="6a781-172">Selecione **Adicionar aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="6a781-172">Select **Add application**.</span></span>

1. <span data-ttu-id="6a781-173">Em **Gerenciar**, selecione **Permissões para API** e selecione **Adicionar uma permissão**.</span><span class="sxs-lookup"><span data-stu-id="6a781-173">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="6a781-174">No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.</span><span class="sxs-lookup"><span data-stu-id="6a781-174">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="6a781-175">Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="6a781-175">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="6a781-176">Selecione estas opções.</span><span class="sxs-lookup"><span data-stu-id="6a781-176">Select the following.</span></span> <span data-ttu-id="6a781-177">Somente o primeiro é realmente necessário pelo seu próprio complemento; mas a permissão é necessária para que o aplicativo Office para obter um `profile` token para seu aplicativo Web de complemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-177">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office application to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="6a781-178">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="6a781-178">Files.Read.All</span></span>
    * <span data-ttu-id="6a781-179">perfil</span><span class="sxs-lookup"><span data-stu-id="6a781-179">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a781-180">A permissão `User.Read` pode já estar listada por padrão.</span><span class="sxs-lookup"><span data-stu-id="6a781-180">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="6a781-181">É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.</span><span class="sxs-lookup"><span data-stu-id="6a781-181">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="6a781-182">Marque a caixa de seleção para cada permissão conforme elas forem exibidas.</span><span class="sxs-lookup"><span data-stu-id="6a781-182">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="6a781-183">Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.</span><span class="sxs-lookup"><span data-stu-id="6a781-183">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="6a781-184">Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Aceitar** para a confirmação exibida.</span><span class="sxs-lookup"><span data-stu-id="6a781-184">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a781-185">Depois de escolher **Conceder consentimento de administrador para [nome do locatário]**, você verá uma mensagem solicitando que você tente novamente alguns minutos depois, para que a solicitação de consentimento possa ser construída.</span><span class="sxs-lookup"><span data-stu-id="6a781-185">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="6a781-186">Nesse caso, você pode começar a trabalhar na próxima seção, mas não se esqueça de voltar ao **_portal e pressionar este botão_**!</span><span class="sxs-lookup"><span data-stu-id="6a781-186">If so, you can start work on the next section, **_but don't forget to come back to the portal and press this button_**!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="6a781-187">Configurar a solução</span><span class="sxs-lookup"><span data-stu-id="6a781-187">Configure the solution</span></span>

1. <span data-ttu-id="6a781-188">Na raiz da pasta **Before** (antes), abra o arquivo de solução (.sln) no **Visual Studio**.</span><span class="sxs-lookup"><span data-stu-id="6a781-188">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="6a781-189">Clique com o botão direito do mouse no nó superior no **Gerenciador de Soluções** (no nó Solução, não em qualquer um dos nós do projeto) e selecione **Configurar projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="6a781-189">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="6a781-190">Em **Propriedades Comuns**, selecione **Projeto de Inicialização** e **Vários projetos de inicialização**.</span><span class="sxs-lookup"><span data-stu-id="6a781-190">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="6a781-191">Verifique se a **Ação** para ambos os projetos está definida como **Iniciar** e se o projeto terminado em "...WebAPI" está listado primeiro.</span><span class="sxs-lookup"><span data-stu-id="6a781-191">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="6a781-192">Feche a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="6a781-192">Close the dialog.</span></span>

1. <span data-ttu-id="6a781-193">No **Explorador de** Soluções, selecione (não clique com o botão direito do mouse) no projeto **Office-Add-in-ASPNET-SSO-WebAPI.**</span><span class="sxs-lookup"><span data-stu-id="6a781-193">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="6a781-194">O painel **Propriedades** é exibido.</span><span class="sxs-lookup"><span data-stu-id="6a781-194">The **Properties** pane opens.</span></span> <span data-ttu-id="6a781-195">Verifique se **SSL Habilitado** é **Verdadeiro**.</span><span class="sxs-lookup"><span data-stu-id="6a781-195">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="6a781-196">Verifique se a **URL do SSL** é `http://localhost:44355/`.</span><span class="sxs-lookup"><span data-stu-id="6a781-196">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="6a781-197">Em "Web.config", use os valores copiados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="6a781-197">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="6a781-198">Defina **ida:ClientID** e **ida:Audience** para sua **ID do aplicativo (cliente)** e defina **ida:Password** para a senha de cliente.</span><span class="sxs-lookup"><span data-stu-id="6a781-198">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span> <span data-ttu-id="6a781-199">Além disso, de definir **ida:Domain** como `http://localhost:44355` (sem barra de encaminhamento "/" no final).</span><span class="sxs-lookup"><span data-stu-id="6a781-199">Also, set **ida:Domain** to `http://localhost:44355` (no forward slash "/" at the end).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6a781-200">A ID de aplicativo **(cliente)** é o valor "público" quando outros aplicativos, como o aplicativo cliente Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6a781-200">The **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="6a781-201">Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6a781-201">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="6a781-202">Se você não tiver escolhido "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, salve e feche o Web.config. Caso contrário, salve, mas deixe-o aberto.</span><span class="sxs-lookup"><span data-stu-id="6a781-202">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="6a781-203">Ainda no **Explorador** de Soluções, escolha o projeto **Office-Add-in-ASPNET-SSO** e abra o arquivo de manifesto do complemento "Office-Add-in-ASPNET-SSO.xml" e role até a parte inferior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-203">Still in **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="6a781-204">Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="6a781-204">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

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

1. <span data-ttu-id="6a781-205">Substitua o espaço reservado "{$application_GUID here$}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-205">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="6a781-206">Os sinais "$" não fazem parte da ID, portanto não os inclua.</span><span class="sxs-lookup"><span data-stu-id="6a781-206">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="6a781-207">Essa é a mesma ID usada para a ClientID e a Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="6a781-207">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a781-208">O valor **Recurso** é o **URI da ID de aplicativo** que você definiu quando registrou o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-208">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="6a781-209">A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="6a781-209">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="6a781-210">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-210">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="6a781-211">Configuração para locatário único</span><span class="sxs-lookup"><span data-stu-id="6a781-211">Setup for single-tenant</span></span>

<span data-ttu-id="6a781-212">Se você escolher "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, você execute estas etapas adicionais de configuração:</span><span class="sxs-lookup"><span data-stu-id="6a781-212">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="6a781-213">Volte para o Portal do Azure e abra a lâmina **Visão geral** do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-213">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="6a781-214">Copie a **ID de diretório (locatário)**.</span><span class="sxs-lookup"><span data-stu-id="6a781-214">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="6a781-215">Em Web.config, substitua o "comum" no valor de **ida:Authority** pela GUID copiada na etapa anterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-215">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="6a781-216">Ao terminar, o valor deverá ser similar a este: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span><span class="sxs-lookup"><span data-stu-id="6a781-216">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="6a781-217">Salve e feche o web.config.</span><span class="sxs-lookup"><span data-stu-id="6a781-217">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="6a781-218">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="6a781-218">Code the client side</span></span>

1. <span data-ttu-id="6a781-219">Abra o arquivo HomeES6.js na pasta **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="6a781-219">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="6a781-220">Ele já apresenta alguns códigos:</span><span class="sxs-lookup"><span data-stu-id="6a781-220">It already has some code in it:</span></span>

    * <span data-ttu-id="6a781-221">Um polyfill que atribui o objeto Office.Promise ao objeto de janela global, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="6a781-221">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="6a781-222">(Para obter mais detalhes, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="6a781-222">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="6a781-223">Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="6a781-223">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="6a781-224">Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6a781-224">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="6a781-225">Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.</span><span class="sxs-lookup"><span data-stu-id="6a781-225">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="6a781-226">O código implementa o sistema de autorização de fallback que o suplemento usará em situações em que o SSO não é compatível ou gera um erro.</span><span class="sxs-lookup"><span data-stu-id="6a781-226">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="6a781-p134">Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="6a781-p134">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="6a781-229">O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções.</span><span class="sxs-lookup"><span data-stu-id="6a781-229">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="6a781-230">A variável de contador `retryGetAccessToken` é usada para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="6a781-230">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="6a781-231">A função `getGraphData` é definida com a palavra-chave ES6 `async`.</span><span class="sxs-lookup"><span data-stu-id="6a781-231">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="6a781-232">Usar a sintaxe ES6 facilita o uso da API de SSO em Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="6a781-232">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="6a781-233">Esse é o único arquivo na solução que usará a sintaxe sem suporte do Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="6a781-233">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="6a781-234">Colocamos "ES6" no nome do arquivo como um lembrete.</span><span class="sxs-lookup"><span data-stu-id="6a781-234">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="6a781-235">A solução usa o transcompilador de tsc para transcompilar esse arquivo em ES5, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="6a781-235">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="6a781-236">(Veja o arquivo tsconfig.json na raiz do projeto.)</span><span class="sxs-lookup"><span data-stu-id="6a781-236">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="6a781-237">Abaixo da função `getGraphData`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-237">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="6a781-238">Observe que você criará a função `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-238">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

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

1. <span data-ttu-id="6a781-239">Substitua `TODO 1` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-239">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="6a781-240">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="6a781-240">About this code, note:</span></span>

    * <span data-ttu-id="6a781-241">`getAccessToken` instrui o Office a obter um token de bootstrap do Azure AD e retornar ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-241">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="6a781-242">`allowSignInPrompt` indica ao Office para solicitar que o usuário entre caso ele ainda não esteja conectado ao Office.</span><span class="sxs-lookup"><span data-stu-id="6a781-242">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="6a781-243">`allowConsentPrompt`informa Office solicitar que o usuário consenta em permitir que o complemento acesse o perfil AAD do usuário, se o consentimento ainda não tiver sido concedido.</span><span class="sxs-lookup"><span data-stu-id="6a781-243">`allowConsentPrompt` tells Office to prompt the user to consent to letting the add-in access the user's AAD profile, if consent has not already been granted.</span></span> <span data-ttu-id="6a781-244">(O prompt resultante não *permite* que o usuário consenta com quaisquer escopos Graph Microsoft.)</span><span class="sxs-lookup"><span data-stu-id="6a781-244">(The resulting prompt does *not* allow the user to consent to any Microsoft Graph scopes.)</span></span>
    * <span data-ttu-id="6a781-245">`forMSGraphAccess` instrui o Office que o suplemento pretende trocar o token de bootstrap por um token de acesso ao Micrsoft Graph, em vez de apenas usar o token de bootstrap como um token de ID.</span><span class="sxs-lookup"><span data-stu-id="6a781-245">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="6a781-246">A configuração dessa opção dá ao Office a oportunidade de cancelar o processo de obtenção do token de bootstrap (e retornar o código de erro 13012) se o administrador de locatários do usuário não tiver concedido consentimento para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-246">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="6a781-247">O código do lado do cliente do suplemento pode responder ao 13012 por meio da ramificação para um sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="6a781-247">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="6a781-248">Se o não for usado e o administrador não tiver concedido consentimento, o token bootstrap será retornado, mas a tentativa de trocar com o fluxo on-behalf-of resultaria em `forMSGraphAccess` um erro.</span><span class="sxs-lookup"><span data-stu-id="6a781-248">If the `forMSGraphAccess` is not used and the admin has not granted consent, the bootstrap token is returned, but the attempt to exchange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="6a781-249">Portanto, a opção `forMSGraphAccess` permite ao suplemento ramificar para o sistema de fallback rapidamente.</span><span class="sxs-lookup"><span data-stu-id="6a781-249">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="6a781-250">Você criará a função `getData` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-250">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="6a781-251">O parâmetro `/api/values` é a URL de um controlador do lado do servidor que fará a troca de tokens e usará o token de acesso recebido para fazer a chamada para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6a781-251">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="6a781-252">Abaixo da função `getGraphData`, adicione o seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-252">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="6a781-253">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="6a781-253">About this code, note:</span></span>

    * <span data-ttu-id="6a781-254">Ele é usado pelos sistemas de autorização de fallback e SSO.</span><span class="sxs-lookup"><span data-stu-id="6a781-254">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="6a781-255">O parâmetro `relativeUrl` é um controlador do lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="6a781-255">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="6a781-256">O parâmetro `accessToken` pode ser um token de bootstrap ou um token de acesso completo.</span><span class="sxs-lookup"><span data-stu-id="6a781-256">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="6a781-257">O `writeFileNamesToOfficeDocument` já faz parte do projeto.</span><span class="sxs-lookup"><span data-stu-id="6a781-257">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="6a781-258">Você criará a função `handleServerSideErrors` em uma última etapa.</span><span class="sxs-lookup"><span data-stu-id="6a781-258">You create the `handleServerSideErrors` function in a later step.</span></span>

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

### <a name="handle-client-side-errors"></a><span data-ttu-id="6a781-259">Tratar erros do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="6a781-259">Handle client-side errors</span></span>

1. <span data-ttu-id="6a781-260">Abaixo da função `getData`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-260">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="6a781-261">Observe que `error.code` é um número, normalmente no intervalo 13xxx.</span><span class="sxs-lookup"><span data-stu-id="6a781-261">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

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

1. <span data-ttu-id="6a781-262">Substitua `TODO 2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-262">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="6a781-263">Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6a781-263">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
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

1. <span data-ttu-id="6a781-264">Substitua `TODO 3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-264">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="6a781-265">Para todos os outros erros, o suplemento ramificará para o sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="6a781-265">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="6a781-266">Para obter mais informações sobre esses erros, consulte [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). Nesse complemento, o sistema de fallback abre uma caixa de diálogo que exige que o usuário entre, mesmo que o usuário já tenha.</span><span class="sxs-lookup"><span data-stu-id="6a781-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="6a781-267">Resolver erros do lado do servidor</span><span class="sxs-lookup"><span data-stu-id="6a781-267">Handle server-side errors</span></span>

1. <span data-ttu-id="6a781-268">Abaixo da função `handleClientSideErrors`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-268">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="6a781-269">Substitua `TODO 4` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-269">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="6a781-270">Sobre esse código, observe que as classes de erro ASP.NET foram criadas antes de haver algo como a MFA.</span><span class="sxs-lookup"><span data-stu-id="6a781-270">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="6a781-271">Como um efeito colateral de como a lógica do lado do servidor lida com as solicitações de um segundo fator de autenticação, o erro do lado do servidor enviado para o cliente tem uma propriedade **Message**, mas não uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="6a781-271">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="6a781-272">Mas todos os outros erros terão uma propriedade **ExceptionMessage**, para que o código do cliente precise analisar a resposta para ambos.</span><span class="sxs-lookup"><span data-stu-id="6a781-272">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="6a781-273">Uma ou outra variável será indefinida.</span><span class="sxs-lookup"><span data-stu-id="6a781-273">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="6a781-274">Substitua `TODO 5` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-274">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="6a781-275">Quando o Microsoft Graph requer uma forma adicional de autenticação, ele envia o erro AADSTS50076.</span><span class="sxs-lookup"><span data-stu-id="6a781-275">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="6a781-276">Isso inclui informações sobre os requisitos adicionais na propriedade **Message.Claims**.</span><span class="sxs-lookup"><span data-stu-id="6a781-276">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="6a781-277">Para lidar com isso, o código faz uma segunda tentativa de obter o token de bootstrap, mas, desta vez, ele inclui a solicitação de um fator adicional, como o valor da opção `authChallenge`, que informa ao Azure AD a solicitar todos os formulários de autenticação necessários.</span><span class="sxs-lookup"><span data-stu-id="6a781-277">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

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

1. <span data-ttu-id="6a781-278">Substitua `TODO 6` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-278">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="6a781-279">Substitua `TODO 7` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-279">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="6a781-280">Observe que, em raras ocasiões, o token de bootstrap fica não vencido quando o Office o valida, mas vence no momento em que ele é enviado ao Azure AD para o Exchange.</span><span class="sxs-lookup"><span data-stu-id="6a781-280">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="6a781-281">O Azure AD responderá com o erro AADSTS500133.</span><span class="sxs-lookup"><span data-stu-id="6a781-281">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="6a781-282">Quando isso acontece, o código recupera a API de SSO (mas não mais de uma vez).</span><span class="sxs-lookup"><span data-stu-id="6a781-282">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="6a781-283">Desta vez, o Office retorna um novo token de bootstrap não vencido.</span><span class="sxs-lookup"><span data-stu-id="6a781-283">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="6a781-284">Substitua `TODO 8` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-284">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="6a781-285">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-285">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="6a781-286">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="6a781-286">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="6a781-287">Configurar o middleware OWIN</span><span class="sxs-lookup"><span data-stu-id="6a781-287">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="6a781-288">Abra o arquivo Startup.cs na raiz do projeto **Office-Add-in-ASPNET-SSO-WebAPI** e adicione o seguinte método à classe **Inicialização**.</span><span class="sxs-lookup"><span data-stu-id="6a781-288">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="6a781-289">Observe que você criará o método `ConfigureAuth` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6a781-289">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="6a781-290">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-290">Save and close the file.</span></span>

1. <span data-ttu-id="6a781-291">Clique com botão direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.</span><span class="sxs-lookup"><span data-stu-id="6a781-291">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="6a781-292">Na caixa de diálogo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="6a781-292">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="6a781-293">Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="6a781-293">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="6a781-294">Verifique se todas as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-294">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="6a781-p149">Adicione a palavra-chave `partial` à declaração da classe `Startup`, se ainda não estiver lá. A linha deverá ser assim:</span><span class="sxs-lookup"><span data-stu-id="6a781-p149">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="6a781-p150">Adicione o método a seguir à classe `Startup`. Este método especifica como o middleware OWIN validará os tokens de acesso que são transmitidos a ele do método `getData` no arquivo Home.js do lado do cliente. O processo de autorização é disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` é chamado.</span><span class="sxs-lookup"><span data-stu-id="6a781-p150">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="6a781-300">Substitua o `TODO 1` pelo seguinte.</span><span class="sxs-lookup"><span data-stu-id="6a781-300">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="6a781-301">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="6a781-301">Note about this code:</span></span>

    * <span data-ttu-id="6a781-302">O código instrui o OWIN a garantir que o público especificado no token bootstrap que vem do aplicativo Office deve corresponder ao valor especificado no web.config.</span><span class="sxs-lookup"><span data-stu-id="6a781-302">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office application must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="6a781-303">As contas da Microsoft têm um GUID de emissor diferente de qualquer GUID de locatário organizacional, portanto, para dar suporte a ambos os tipos de contas, não validamos o emissor.</span><span class="sxs-lookup"><span data-stu-id="6a781-303">Microsoft accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="6a781-304">A `SaveSigninToken` `true` configuração como faz com que o OWIN salve o token de inicialização bruto do aplicativo Office aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6a781-304">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office application.</span></span> <span data-ttu-id="6a781-305">O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on-behalf-of".</span><span class="sxs-lookup"><span data-stu-id="6a781-305">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="6a781-306">Os escopos não são validados pelo middleware OWIN.</span><span class="sxs-lookup"><span data-stu-id="6a781-306">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="6a781-307">Os escopos do token de bootstrap, que devem conter `access_as_user`, são validados no controlador.</span><span class="sxs-lookup"><span data-stu-id="6a781-307">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="6a781-p154">Substitua `TODO 2` pelo seguinte. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="6a781-p154">Replace `TODO 2` with the following. Note about this code:</span></span>

    * <span data-ttu-id="6a781-310">O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="6a781-310">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="6a781-311">A URL passada para o método é onde o middleware OWIN obtém instruções para obter a chave que precisa para verificar a assinatura no token bootstrap recebido do aplicativo Office.</span><span class="sxs-lookup"><span data-stu-id="6a781-311">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office application.</span></span> <span data-ttu-id="6a781-312">O segmento de Autoridade da URL vem do Web.config. Ele é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.</span><span class="sxs-lookup"><span data-stu-id="6a781-312">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="6a781-313">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-313">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="6a781-314">Criar o controlador /api/values</span><span class="sxs-lookup"><span data-stu-id="6a781-314">Create the /api/values controller</span></span>

1. <span data-ttu-id="6a781-315">Abra o arquivo **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="6a781-315">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="6a781-316">Esse controlador é usado quando o sistema SSO obtém um token de bootstrap com êxito.</span><span class="sxs-lookup"><span data-stu-id="6a781-316">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="6a781-317">Ele não é usado como parte do sistema de autorização de fallback.</span><span class="sxs-lookup"><span data-stu-id="6a781-317">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="6a781-318">Esse sistema usou o AzureADAuthController que foi criado para você.</span><span class="sxs-lookup"><span data-stu-id="6a781-318">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="6a781-319">Verifique se as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-319">Ensure that the following `using` statements are at the top of the file.</span></span>

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

1. <span data-ttu-id="6a781-p157">Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização configurado no último procedimento sempre que um método controlador for chamado. Apenas os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.</span><span class="sxs-lookup"><span data-stu-id="6a781-p157">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="6a781-323">Adicione o método a seguir ao `ValuesController`.</span><span class="sxs-lookup"><span data-stu-id="6a781-323">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="6a781-324">Observe que é o valor de retorno é `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um método `GET api/values`.</span><span class="sxs-lookup"><span data-stu-id="6a781-324">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="6a781-325">Esse é um efeito colateral do fato de que a lógica de autorização OAuth deve estar no controlador, em vez de em um filtro ASP.NET.</span><span class="sxs-lookup"><span data-stu-id="6a781-325">This is a side effect of that fact that the OAuth authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="6a781-326">Algumas condições de erro na lógica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-326">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //         token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. <span data-ttu-id="6a781-327">Substitua `TODO1` pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="6a781-327">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="6a781-328">Observe que o segundo parâmetro do método `SendErrorToClient` é um objeto **Exception**.</span><span class="sxs-lookup"><span data-stu-id="6a781-328">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="6a781-329">Nesse caso, o código passa `null` porque incluir o objeto **Exception** bloqueia a inclusão da propriedade **Message** na resposta HTTP que é gerada.</span><span class="sxs-lookup"><span data-stu-id="6a781-329">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="6a781-330">Substitua `TODO 2` pelo seguinte código para montar todas as informações necessárias para obter um token do Microsoft Graph usando o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="6a781-330">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="6a781-331">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="6a781-331">About this code, note:</span></span>

    * <span data-ttu-id="6a781-332">Seu add-in não está mais desempenhando a função de um recurso (ou audiência) ao qual o aplicativo Office usuário precisa de acesso.</span><span class="sxs-lookup"><span data-stu-id="6a781-332">Your add-in is no longer playing the role of a resource (or audience) to which the Office application and user need access.</span></span> <span data-ttu-id="6a781-333">Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6a781-333">Now it is itself a client that needs access to Microsoft Graph.</span></span> <span data-ttu-id="6a781-334">`ConfidentialClientApplication` é o objeto "client context" da MSAL.</span><span class="sxs-lookup"><span data-stu-id="6a781-334">`ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="6a781-335">A partir da MSAL.NET 3.x.x, o `bootstrapContext` é apenas o token de bootstrap em si.</span><span class="sxs-lookup"><span data-stu-id="6a781-335">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="6a781-336">A Autoridade vem do Web.config. Ela é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.</span><span class="sxs-lookup"><span data-stu-id="6a781-336">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="6a781-337">O MSAL lançará um erro se seu código solicitar , que é realmente usado apenas quando o aplicativo cliente Office obtém o token para o aplicativo Web do `profile` seu complemento.</span><span class="sxs-lookup"><span data-stu-id="6a781-337">MSAL will throw an error if your code requests `profile`, which is really only used when the Office client application gets the token to your add-in's web application.</span></span> <span data-ttu-id="6a781-338">Então, apenas `Files.Read.All` é explicitamente solicitado.</span><span class="sxs-lookup"><span data-stu-id="6a781-338">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. <span data-ttu-id="6a781-p163">Substitua `TODO 3` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="6a781-p163">Replace `TODO 3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6a781-341">O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso.</span><span class="sxs-lookup"><span data-stu-id="6a781-341">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="6a781-342">Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="6a781-342">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="6a781-343">Quaisquer exceções que não forem do tipo `MsalServiceException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente como mensagens `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="6a781-343">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

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

1. <span data-ttu-id="6a781-p165">Substitua `TODO 3a` pelo código a seguir. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="6a781-p165">Replace `TODO 3a` with the following code. About this code, note:</span></span>

    * <span data-ttu-id="6a781-346">Se a autenticação multifator for exigida pelo recurso Microsoft Graph e o usuário ainda não a tiver fornecido, o Azure AD retornará "400 Bad Request" com o erro `AADSTS50076` e uma propriedade **Declarações**.</span><span class="sxs-lookup"><span data-stu-id="6a781-346">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="6a781-347">O MSAL exibe **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informações.</span><span class="sxs-lookup"><span data-stu-id="6a781-347">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="6a781-348">O **valor da** propriedade Claims deve ser passado para o cliente, que deve passá-lo para o aplicativo Office, que o inclui em uma solicitação para um novo token bootstrap.</span><span class="sxs-lookup"><span data-stu-id="6a781-348">The **Claims** property value must be passed to the client which should pass it to the Office application, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="6a781-349">O Azure AD solicitará ao usuário todas as formas de autenticação necessárias.</span><span class="sxs-lookup"><span data-stu-id="6a781-349">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="6a781-p168">As APIs que criam respostas HTTP a partir de exceções não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto de resposta. É necessário criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a criação de uma propriedade **ExceptionMessage**, portanto, a única maneira de obter a ID de erro `AADSTS50076` para o cliente é adicioná-la à **Message** personalizada. O JavaScript no cliente precisará descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.</span><span class="sxs-lookup"><span data-stu-id="6a781-p168">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="6a781-354">A mensagem personalizada é formatada como JSON para que o JavaScript do cliente possa analisá-la com métodos de objeto `JSON` JavaScript conhecidos.</span><span class="sxs-lookup"><span data-stu-id="6a781-354">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="6a781-p169">Substitua `TODO 3b` pelo código a seguir. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="6a781-p169">Replace `TODO 3b` with the following code. About this code, note:</span></span>

    * <span data-ttu-id="6a781-357">Se a chamada para o Azure AD contiver pelo menos um escopo (permissão) que não tenha sido consentido pelo usuário ou por um administrador de locatários (ou se o consentimento foi revogado), o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS65001`.</span><span class="sxs-lookup"><span data-stu-id="6a781-357">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="6a781-358">O MSAL exibe um **MsalUiRequiredException** com essas informações.</span><span class="sxs-lookup"><span data-stu-id="6a781-358">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    * <span data-ttu-id="6a781-359">Se a chamada para o Azure AD contiver pelo menos um escopo que Azure AD não reconhece, o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS70011`.</span><span class="sxs-lookup"><span data-stu-id="6a781-359">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="6a781-360">O MSAL exibe um **MsalUiRequiredException** com essas informações.</span><span class="sxs-lookup"><span data-stu-id="6a781-360">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    * <span data-ttu-id="6a781-361">A descrição completa é incluída porque 70011 é retornado em outras condições e ele deverá ser processado neste suplemento somente quando significar que há um escopo inválido.</span><span class="sxs-lookup"><span data-stu-id="6a781-361">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    * <span data-ttu-id="6a781-p172">O objeto **MsalUiRequiredException** é passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informações de erro seja incluída na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="6a781-p172">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="6a781-364">Substitua `TODO 3c` pelo seguinte código para lidar com todas as outras **MsalServiceException** s.</span><span class="sxs-lookup"><span data-stu-id="6a781-364">Replace `TODO 3c` with the following code to handle all other **MsalServiceException** s.</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="6a781-365">Substitua `TODO 4` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-365">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="6a781-366">O método `GraphApiHelper.GetOneDriveFileNames`, que foi criado para você, faz a solicitação de dados ao Microsoft Graph e inclui o token de acesso.</span><span class="sxs-lookup"><span data-stu-id="6a781-366">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="6a781-367">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6a781-367">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="6a781-368">Executar a solução</span><span class="sxs-lookup"><span data-stu-id="6a781-368">Run the solution</span></span>

1. <span data-ttu-id="6a781-369">Abra o arquivo de solução do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6a781-369">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="6a781-370">No menu **Build**, selecione **Solução Limpa**.</span><span class="sxs-lookup"><span data-stu-id="6a781-370">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="6a781-371">Quando terminar, abra o menu **Build** novamente e selecione **Solução de Build**.</span><span class="sxs-lookup"><span data-stu-id="6a781-371">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="6a781-372">No **Gerenciador de Soluções**, selecione o nó de projeto **Office-Add-in-ASPNET-SSO** (não o nó da solução principal e não o projeto cujo nome termina em "WebAPI").</span><span class="sxs-lookup"><span data-stu-id="6a781-372">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="6a781-373">No painel **Propriedades**, abra o menu suspenso **Iniciar documento** e escolha uma das três opções (Excel, Word ou PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="6a781-373">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![Escolha o aplicativo cliente Office desejado: Excel, PowerPoint ou Word.](../images/SelectHost.JPG)

1. <span data-ttu-id="6a781-375">Pressione F5.</span><span class="sxs-lookup"><span data-stu-id="6a781-375">Press F5.</span></span>
1. <span data-ttu-id="6a781-376">No aplicativo do Office, na faixa de opções **Home**, selecione **Mostrar suplemento** no grupo **SSO ASP.NET** para abrir o suplemento do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6a781-376">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="6a781-377">Clique no botão **Definir Nome de Arquivos do One Drive**.</span><span class="sxs-lookup"><span data-stu-id="6a781-377">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="6a781-378">Se você estiver conectado ao Office com uma conta de Microsoft 365 Education ou de trabalho, ou uma conta da Microsoft, e o SSO estiver funcionando conforme esperado, os primeiros 10 nomes de arquivo e pasta em seu OneDrive for Business serão exibidos no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6a781-378">If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="6a781-379">Se você não estiver conectado ou estiver em um cenário que não dá suporte ao SSO, ou o SSO não estiver funcionando por qualquer motivo, você será solicitado a entrar.</span><span class="sxs-lookup"><span data-stu-id="6a781-379">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to sign in.</span></span> <span data-ttu-id="6a781-380">Depois de entrar, os nomes de arquivo e pasta aparecem.</span><span class="sxs-lookup"><span data-stu-id="6a781-380">After you sign in, the file and folder names appear.</span></span>

### <a name="testing-the-fallback-path"></a><span data-ttu-id="6a781-381">Testar o caminho de fallback</span><span class="sxs-lookup"><span data-stu-id="6a781-381">Testing the fallback path</span></span>

<span data-ttu-id="6a781-382">Para testar o caminho de autorização de fallback, force o caminho do SSO a falhar com as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="6a781-382">To test the fallback authorization path, force the SSO path to fail with the following steps.</span></span>

1. <span data-ttu-id="6a781-383">Adicione o código a seguir à parte superior do `getDataWithToken` método no arquivo HomeES6.js.</span><span class="sxs-lookup"><span data-stu-id="6a781-383">Add the following code to the very top of the `getDataWithToken` method in the HomeES6.js file.</span></span>

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. <span data-ttu-id="6a781-384">Em seguida, adicione a seguinte linha à parte superior do `try` bloco nesse mesmo método, logo acima da chamada para `getAccessToken` .</span><span class="sxs-lookup"><span data-stu-id="6a781-384">Then add the following line to the top of the `try` block in that same method, just above the call to `getAccessToken`.</span></span>

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a><span data-ttu-id="6a781-385">Atualizando o complemento quando você vai para preparação e produção</span><span class="sxs-lookup"><span data-stu-id="6a781-385">Updating the add-in when you go to staging and production</span></span>

<span data-ttu-id="6a781-386">Como todos os Office Web, quando você estiver pronto para mover para um servidor de preparação ou produção, você deve atualizar o domínio no manifesto com o `localhost:44355` novo domínio.</span><span class="sxs-lookup"><span data-stu-id="6a781-386">Like all Office Web Add-ins, when you are ready to move to a staging or production server, you must update the `localhost:44355` domain in the manifest with the new domain.</span></span> <span data-ttu-id="6a781-387">Da mesma forma, você deve atualizar o domínio no arquivo web.config.</span><span class="sxs-lookup"><span data-stu-id="6a781-387">Similarly, you must update the domain in the web.config file.</span></span>

<span data-ttu-id="6a781-388">Como o domínio aparece no registro do AAD, você precisa atualizar esse registro para usar o novo domínio no lugar de onde `localhost:44355` quer que ele apareça.</span><span class="sxs-lookup"><span data-stu-id="6a781-388">Since the domain appears in the AAD registration, you need to update that registration to use the new domain in place of `localhost:44355` wherever it appears.</span></span>
