---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: e8bc5f09b3e9d401fdba992d87ec3ef4faf7fc08
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217083"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="09283-102">Habilitar o logon único para Suplementos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="09283-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="09283-103">Os usuários entram no Office (online, em dispositivos móveis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365).</span><span class="sxs-lookup"><span data-stu-id="09283-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="09283-104">Você pode tirar proveito disso e usar o logon único (SSO) para autorizar usuário para suplemento, sem exigir que o usuário entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="09283-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Imagem mostrando o processo de logon de um suplemento](../images/sso-for-office-addins.png)

## <a name="preview-status"></a><span data-ttu-id="09283-106">Status de visualização</span><span class="sxs-lookup"><span data-stu-id="09283-106">Preview Status</span></span>

<span data-ttu-id="09283-107">A API de logon único tem suporte somente na visualização.</span><span class="sxs-lookup"><span data-stu-id="09283-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="09283-108">Está disponível para os desenvolvedores para experimentação; mas não deve ser usado em um suplemento de produção.</span><span class="sxs-lookup"><span data-stu-id="09283-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="09283-109">Além disso, os suplementos que usam o SSO não são aceitos no [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="09283-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="09283-110">O SSO exige o Office 365 (a versão de assinatura do Office).</span><span class="sxs-lookup"><span data-stu-id="09283-110">SSO requires Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="09283-111">Você deve usar o build e a versão mensal mais recente do canal Insiders.</span><span class="sxs-lookup"><span data-stu-id="09283-111">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="09283-112">É necessário ingressar no programa Office Insider para obter essa versão.</span><span class="sxs-lookup"><span data-stu-id="09283-112">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="09283-113">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="09283-113">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="09283-114">Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.</span><span class="sxs-lookup"><span data-stu-id="09283-114">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

<span data-ttu-id="09283-115">Nem todos os aplicativos do Office oferecem suporte a visualização de SSO.</span><span class="sxs-lookup"><span data-stu-id="09283-115">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="09283-116">Está disponível no Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="09283-116">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="09283-117">Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="09283-117">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span>

## <a name="requirements-and-best-practices"></a><span data-ttu-id="09283-118">Requisitos e as práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="09283-118">Requirements and Best Practices</span></span>

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

<span data-ttu-id="09283-119">Se você estiver trabalhando com um suplemento do **Outlook**, certifique-se de habilitar a Autenticação Moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="09283-119">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="09283-120">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="09283-120">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="09283-121">Você *não* deve confiar no SSO como único método do suplemento de autenticação.</span><span class="sxs-lookup"><span data-stu-id="09283-121">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="09283-122">Devem implementar um sistema de autenticação alternativo que o suplemento possa se enquadrar em determinadas situações de erro.</span><span class="sxs-lookup"><span data-stu-id="09283-122">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="09283-123">Você pode usar um sistema de autenticação e tabelas de usuário ou utilizar um dos provedores de logon de redes sociais.</span><span class="sxs-lookup"><span data-stu-id="09283-123">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="09283-124">Para saber mais sobre como fazer isso com um suplemento do Office, confira [Autorizar serviços externos nos suplementos do Office](/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="09283-124">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="09283-125">Para *Outlook*, há um sistema de fallback recomendado.</span><span class="sxs-lookup"><span data-stu-id="09283-125">For *Outlook*, there is a recommended fallback system.</span></span> <span data-ttu-id="09283-126">Para mais informações, confira [Cenário: implementar o logon único no serviço em um Suplemento do Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="09283-126">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span> <span data-ttu-id="09283-127">Para exemplos que usam o Azure Active Directory como o sistema de fallback, confira [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) e [SSO com Suplemento ASP.NET do Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="09283-127">For samples that use Azure Active Directory as the fallback system, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>

## <a name="how-sso-works-at-runtime"></a><span data-ttu-id="09283-128">Como o SSO funciona em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="09283-128">How SSO works at runtime</span></span>

<span data-ttu-id="09283-129">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="09283-129">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="09283-131">No suplemento, o JavaScript chama uma nova API do Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="09283-131">In the add-in, JavaScript calls a new Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span> <span data-ttu-id="09283-132">Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-132">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="09283-133">Confira [Token de acesso de amostra](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="09283-133">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="09283-134">Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.</span><span class="sxs-lookup"><span data-stu-id="09283-134">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="09283-135">Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.</span><span class="sxs-lookup"><span data-stu-id="09283-135">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="09283-136">O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual. </span><span class="sxs-lookup"><span data-stu-id="09283-136">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="09283-137">O Azure AD envia o token do suplemento ao aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="09283-137">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="09283-138">O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="09283-138">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessToken` call.</span></span>
7. <span data-ttu-id="09283-139">O JavaScript no suplemento pode analisar o token e extrair informações necessárias, como endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="09283-139">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span>
8. <span data-ttu-id="09283-140">Opcionalmente, o suplemento pode enviar solicitação HTTP para o servidor para obter mais dados sobre o usuário; como as preferências do usuário.</span><span class="sxs-lookup"><span data-stu-id="09283-140">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="09283-141">Como alternativa, o próprio token de acesso pode ser enviado para o servidor para análise e validação.</span><span class="sxs-lookup"><span data-stu-id="09283-141">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="09283-142">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="09283-142">Develop an SSO add-in</span></span>

<span data-ttu-id="09283-143">Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO.</span><span class="sxs-lookup"><span data-stu-id="09283-143">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="09283-144">Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agnóstica.</span><span class="sxs-lookup"><span data-stu-id="09283-144">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="09283-145">Para orientações detalhadas, confira:</span><span class="sxs-lookup"><span data-stu-id="09283-145">For detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="09283-146">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="09283-146">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="09283-147">Criar um Suplemento do Office com ASP.NET que usa logon único</span><span class="sxs-lookup"><span data-stu-id="09283-147">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> <span data-ttu-id="09283-148">Você pode usar o gerador Yeoman para criar um Suplemento do Office com Node.js habilitado para SSO.</span><span class="sxs-lookup"><span data-stu-id="09283-148">You can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="09283-149">O gerador Yeoman simplifica o processo de criação de um suplemento habilitado para SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO.</span><span class="sxs-lookup"><span data-stu-id="09283-149">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="09283-150">Para obter mais informações, confira [Início rápido de logon único (SSO)](../quickstarts/sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="09283-150">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

### <a name="create-the-service-application"></a><span data-ttu-id="09283-151">Criar o aplicativo de serviço</span><span class="sxs-lookup"><span data-stu-id="09283-151">Create the service application</span></span>

<span data-ttu-id="09283-p111">Registre o suplemento no portal de registro para o ponto de extremidade do Azure v 2.0. Esse é um processo que leva entre 5 e 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="09283-p111">Register the add-in at the registration portal for the Azure v2.0 endpoint. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="09283-154">Obter um ID de cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-154">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="09283-155">Especificar as permissões que seu suplemento precisa de AAD v.</span><span class="sxs-lookup"><span data-stu-id="09283-155">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="09283-156">ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="09283-156">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="09283-157">A permissão "perfil" sempre é necessária.</span><span class="sxs-lookup"><span data-stu-id="09283-157">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="09283-158">Conceder a confiança do aplicativo host do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-158">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="09283-159">Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="09283-159">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="09283-160">Para mais detalhes sobre esse processo, confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="09283-160">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="09283-161">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="09283-161">Configure the add-in</span></span>

<span data-ttu-id="09283-162">Adicione novas marcações ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="09283-162">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="09283-163">**WebApplicationInfo** – o pai dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="09283-163">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="09283-164">**ID** - O ID do cliente do suplemento Este é um ID do aplicativo que você obtém como parte do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-164">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="09283-165">Confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="09283-165">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="09283-166">**Resource** – A URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-166">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="09283-167">Esse é o mesmo URI (incluindo o protocolo `api:`) que você usou ao registrar o suplemento no AAD.</span><span class="sxs-lookup"><span data-stu-id="09283-167">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="09283-168">Parte de domínio deste URI deve coincidir com o domínio, incluindo qualquer subdomínio, que o usado nas URLs na seção `<Resources>` do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-168">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="09283-169">**Scopes** – O pai de uma ou mais elementos **Scope**.</span><span class="sxs-lookup"><span data-stu-id="09283-169">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="09283-170">**Scope** – Especifica uma permissão que seu suplemento precisa para o AAD.</span><span class="sxs-lookup"><span data-stu-id="09283-170">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="09283-171">A permissão `profile` sempre é necessária, e pode ser a única permissão necessária, se o suplemento não acessar o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="09283-171">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="09283-172">Se isso acontecer, você também precisa de elementos **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="09283-172">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="09283-173">Bibliotecas que você usa no seu código para acessar o Microsoft Graph pode precisar de permissões adicionais.</span><span class="sxs-lookup"><span data-stu-id="09283-173">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="09283-174">Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="09283-174">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="09283-175">Para saber mais, confira [autorizar o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="09283-175">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="09283-p116">Para hosts do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="09283-p116">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="09283-178">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="09283-178">The following is an example of the markup:</span></span>

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a><span data-ttu-id="09283-179">Adicionar código do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="09283-179">Add client-side code</span></span>

<span data-ttu-id="09283-180">Adicione o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="09283-180">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="09283-181">Chamar [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="09283-181">Call [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span>

* <span data-ttu-id="09283-182">Analisar o token de acesso ou encaminhá-lo ao código de servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-182">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="09283-183">Aqui está um exemplo simples de uma chamada para `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="09283-183">Here's a simple example of a call to `getAccessToken`.</span></span>

> [!NOTE]
> <span data-ttu-id="09283-184">Este exemplo lida explicitamente com apenas um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="09283-184">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="09283-185">Para exemplos de tratamento de erro mais elaborados, confira [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) e [SSO com Suplemento ASP.NET do Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="09283-185">For examples of more elaborate error handling, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>


```js
async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });

        // The /api/values controller will make the token exchange and use the
        // access token it gets back to make the call to MS Graph.
        getData("/api/DoSomething", bootstrapToken);
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
}
```

<span data-ttu-id="09283-186">Aqui está um exemplo simples de como passar o token de suplemento para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="09283-186">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="09283-187">O token é incluído como um cabeçalho de `Authorization`ao enviar uma solicitação para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="09283-187">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="09283-188">Este exemplo prevê enviar dados JSON, para que ele tenha o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver escrevendo no servidor.</span><span class="sxs-lookup"><span data-stu-id="09283-188">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + bootstrapToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a><span data-ttu-id="09283-189">Quando chamar o método</span><span class="sxs-lookup"><span data-stu-id="09283-189">When to call the method</span></span>

<span data-ttu-id="09283-190">Se o seu suplemento não puder ser usado quando nenhum usuário estiver logado no Office, então será necessário chamar o `getAccessToken` *quando o suplemento for iniciado* e passar `allowSignInPrompt: true` no parâmetro de `options` do `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="09283-190">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessToken` *when the add-in launches* and pass `allowSignInPrompt: true` in the `options` parameter of `getAccessToken`.</span></span>

<span data-ttu-id="09283-191">Se o complemento tiver alguma funcionalidade que não exija um usuário conectado, então chame `getAccessToken` \* quando o usuário fizer uma ação que exija acesso a um usuário logado\*.</span><span class="sxs-lookup"><span data-stu-id="09283-191">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessToken` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="09283-192">Não há uma degradação significativa do desempenho com chamadas redundantes de `getAccessToken` porque o Office armazena em cache o token de inicialização e o reutilizará, até que ele expire, sem fazer outra chamada para o AAD v.</span><span class="sxs-lookup"><span data-stu-id="09283-192">There is no significant performance degradation with redundant calls of `getAccessToken` because Office caches the bootstrap token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="09283-193">Ponto de extremidade 2.0 sempre que `getAccessToken` for chamado.</span><span class="sxs-lookup"><span data-stu-id="09283-193">2.0 endpoint whenever `getAccessToken` is called.</span></span> <span data-ttu-id="09283-194">Portanto, você pode adicionar chamadas de `getAccessToken` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.</span><span class="sxs-lookup"><span data-stu-id="09283-194">So you can add calls of `getAccessToken` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="09283-195">Adicionar código no lado do servidor</span><span class="sxs-lookup"><span data-stu-id="09283-195">Add server-side code</span></span>

<span data-ttu-id="09283-196">Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá.</span><span class="sxs-lookup"><span data-stu-id="09283-196">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="09283-197">Algumas tarefas de servidor que o suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="09283-197">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="09283-198">Criar um ou mais métodos de Web API com informações sobre o usuário que são extraídas do token; Por exemplo, uma forma que procura preferências do usuário em seu banco de dados hospedado.</span><span class="sxs-lookup"><span data-stu-id="09283-198">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="09283-199">(Confira **usando o token SSO, como uma identidade** abaixo.)Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código que você precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="09283-199">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="09283-200">Obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="09283-200">Get Microsoft Graph data.</span></span> <span data-ttu-id="09283-201">O código do lado do servidor precisa fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="09283-201">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="09283-202">Iniciar o fluxo "on behalf of" com uma chamada para o ponto de extremidade v 2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo).</span><span class="sxs-lookup"><span data-stu-id="09283-202">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="09283-203">O token de acesso nesse contexto é chamado de bootstrap token.</span><span class="sxs-lookup"><span data-stu-id="09283-203">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="09283-204">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="09283-204">Get data from Microsoft Graph by using the new token.</span></span>
    * <span data-ttu-id="09283-205">Opcionalmente, valide o token de acesso antes de iniciar o fluxo (confira **Validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="09283-205">Optionally, before initiating the flow, validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="09283-206">Opcionalmente, após a conclusão do fluxo on-behalf-of, armazene em cache o novo token de acesso retornado do fluxo, de forma que ele seja reutilizado em outras chamadas para o Microsoft Graph até que ele expire.</span><span class="sxs-lookup"><span data-stu-id="09283-206">Optionally, after the on-behalf-of flow completes, cache the new access token that is returned from the flow so that it an be reused in other calls to Microsoft Graph until it expires.</span></span>

 <span data-ttu-id="09283-207">Para saber mais sobre como obter acesso autorizado aos dados do usuário Microsoft Graph, confira [Autorizar o Microsoft Graph nos suplementos do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="09283-207">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="09283-208">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="09283-208">Validate the access token</span></span>

<span data-ttu-id="09283-209">Após a API da Web receber o token de acesso, ela deve validá-lo antes que ele possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="09283-209">Once the Web API receives the access token, it can validate it before using it.</span></span> <span data-ttu-id="09283-210">O token é um Token Web JSON (JWT) e isso significa que validação funciona como uma validação de token na maioria dos fluxos padrão do OAuth.</span><span class="sxs-lookup"><span data-stu-id="09283-210">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="09283-211">Há diversas bibliotecas disponíveis que podem lidar com a validação de JWT. No entanto, as noções básicas incluem:</span><span class="sxs-lookup"><span data-stu-id="09283-211">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="09283-212">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="09283-212">Checking that the token is well-formed</span></span>
- <span data-ttu-id="09283-213">Verificar se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="09283-213">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="09283-214">Verificar se o token está direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="09283-214">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="09283-215">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="09283-215">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="09283-216">Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="09283-216">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="09283-217">A declaração `iss` no token deve começar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="09283-217">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="09283-218">O parâmetro `aud` do token será configurado como a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-218">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="09283-219">O parâmetro `scp` do token será definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="09283-219">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="09283-220">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="09283-220">Using the SSO token as an identity</span></span>

<span data-ttu-id="09283-221">Se o suplemento precisar verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade.</span><span class="sxs-lookup"><span data-stu-id="09283-221">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="09283-222">As seguintes declarações no token estão relacionadas à identidade.</span><span class="sxs-lookup"><span data-stu-id="09283-222">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="09283-223">`name` – O nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="09283-223">`name` - The user's display name.</span></span>
- <span data-ttu-id="09283-224">`preferred_username`O endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="09283-224">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="09283-225">`oid` – Um GUID que representa a ID do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="09283-225">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="09283-226">`tid` – Um GUID que representa a ID da organização do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="09283-226">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="09283-227">Como os valores `name` e `preferred_username` podem alterar, recomendamos que os valores `oid` e `tid` sejam usados para correlacionar a identidade com o serviço de autorização do back-end.</span><span class="sxs-lookup"><span data-stu-id="09283-227">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="09283-228">Por exemplo, o serviço poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados do usuário interno.</span><span class="sxs-lookup"><span data-stu-id="09283-228">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="09283-229">Em seguida, nas solicitações subsequentes, o usuário poderia ser recuperado usando o mesmo valor e o acesso a recursos específicos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.</span><span class="sxs-lookup"><span data-stu-id="09283-229">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="09283-230">Token de acesso de exemplo</span><span class="sxs-lookup"><span data-stu-id="09283-230">Example access token</span></span>

<span data-ttu-id="09283-231">A seguir está uma carga decodificada típica do token de acesso.</span><span class="sxs-lookup"><span data-stu-id="09283-231">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="09283-232">Para saber mais sobre as propriedades, confira [Referência de tokens de versão do Azure Active Directory 2.0](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="09283-232">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>

```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",
    iat: 1521143967,
    nbf: 1521143967,
    exp: 1521147867,
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",
    azpacr: "0",
    e_exp: 262800,
    name: "Mila Nikolova",
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",
    preferred_username: "milan@contoso.com",
    scp: "access_as_user",
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",
    uti: "MICAQyhrH02ov54bCtIDAA",
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="09283-233">Usando o SSO com um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="09283-233">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="09283-234">Há algumas diferenças pequenas, mas importantes entre usar o SSO em um suplemento do Outlook e em um suplemento do Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="09283-234">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="09283-235">Não deixe de ler [Autenticar o usuário com um token de logon único em um suplemento do Outlook](/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar o logon único ao serviço em um suplemento do Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="09283-235">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="09283-236">Referência da API do SSO</span><span class="sxs-lookup"><span data-stu-id="09283-236">SSO API reference</span></span>

### <a name="getaccesstoken"></a><span data-ttu-id="09283-237">getAccessToken</span><span class="sxs-lookup"><span data-stu-id="09283-237">getAccessToken</span></span>

<span data-ttu-id="09283-238">O namespace [Auth](/javascript/api/office-runtime/officeruntime.auth) do OfficeRuntime, `OfficeRuntime.Auth`, fornece um método, `getAccessToken` que permite com que o host do Office obtenha um token de acesso para o aplicativo web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-238">The OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) namespace, `OfficeRuntime.Auth`, provides a method, `getAccessToken` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="09283-239">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="09283-239">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessToken(options?: AuthOptions: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="09283-240">O método chama o ponto de extremidade do Azure Active Directory V 2.0 para obter um token de acesso para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="09283-240">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="09283-241">Isso permite que os suplementos identifiquem usuários.</span><span class="sxs-lookup"><span data-stu-id="09283-241">This enables add-ins to identify users.</span></span> <span data-ttu-id="09283-242">O código do lado do servidor pode usar esse token para acessar o Microsoft Graph do aplicativo Web do suplemento usando o [fluxo OAuth "em nome de"](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="09283-242">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="09283-243">No Outlook, não há suporte para esse API se o suplemento for carregado em uma caixa de correio do Gmail ou do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="09283-243">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="09283-244">Hosts</span><span class="sxs-lookup"><span data-stu-id="09283-244">Hosts</span></span>|<span data-ttu-id="09283-245">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="09283-245">Excel, OneNote, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="09283-246">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="09283-246">Requirement sets</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)|[<span data-ttu-id="09283-247">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="09283-247">IdentityAPI</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)|

#### <a name="parameters"></a><span data-ttu-id="09283-248">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="09283-248">Parameters</span></span>

<span data-ttu-id="09283-249">`options` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="09283-249">`options` - Optional.</span></span> <span data-ttu-id="09283-250">Aceite um objeto [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) (veja abaixo) para definir comportamentos de logon.</span><span class="sxs-lookup"><span data-stu-id="09283-250">Accepts an [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="09283-251">`callback` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="09283-251">`callback` - Optional.</span></span> <span data-ttu-id="09283-252">Aceita um método de retorno que possa analisar o token de ID de usuário ou usar o token fluxo "em nome de" para obter acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="09283-252">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="09283-253">Se [AsyncResult](/javascript/api/office/office.asyncresult) `.status` é "bem-sucedido", em seguida, `AsyncResult.value` é o v AAD bruto.</span><span class="sxs-lookup"><span data-stu-id="09283-253">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="09283-254">token de acesso 2.0 formatado.</span><span class="sxs-lookup"><span data-stu-id="09283-254">2.0-formatted access token.</span></span>

<span data-ttu-id="09283-255">A interface [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) fornece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v.</span><span class="sxs-lookup"><span data-stu-id="09283-255">The [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="09283-256">2.0 com o método`getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="09283-256">2.0 with the `getAccessToken` method.</span></span>
