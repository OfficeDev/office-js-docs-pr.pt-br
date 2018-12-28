---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: a823c1c58a79d97094e97ffe66f082b45601c10c
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457870"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="3789c-102">Habilitar o logon único para Suplementos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="3789c-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="3789c-103">Os usuários entram no Office (online, em dispositivos móveis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365).</span><span class="sxs-lookup"><span data-stu-id="3789c-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="3789c-104">Você pode tirar proveito disso e usar o logon único (SSO) para autorizar usuário para suplemento, sem exigir que o usuário entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="3789c-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="3789c-106">Status de visualização</span><span class="sxs-lookup"><span data-stu-id="3789c-106">Preview Status</span></span>

<span data-ttu-id="3789c-107">A API de logon único tem suporte somente na visualização.</span><span class="sxs-lookup"><span data-stu-id="3789c-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="3789c-108">Está disponível para os desenvolvedores para experimentação; mas não deve ser usado em um suplemento de produção.</span><span class="sxs-lookup"><span data-stu-id="3789c-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="3789c-109">Além disso, os suplementos que usam o SSO não são aceitos no [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="3789c-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="3789c-110">Nem todos os aplicativos do Office oferecem suporte a visualização SSO.</span><span class="sxs-lookup"><span data-stu-id="3789c-110">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="3789c-111">Está disponível no Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="3789c-111">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="3789c-112">Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="3789c-112">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="3789c-113">Requisitos e as práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="3789c-113">Requirements and Best Practices</span></span>

<span data-ttu-id="3789c-114">Para usar o SSO, você deve carregar a versão beta da biblioteca de JavaScript do Office de `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` na página de inicialização HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="3789c-115">Se você estiver trabalhando com um suplemento do **Outlook**, certifique-se de habilitar a Autenticação Moderna para o locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="3789c-115">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="3789c-116">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="3789c-116">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="3789c-117">Você *não* deve confiar no SSO como único método do suplemento de autenticação.</span><span class="sxs-lookup"><span data-stu-id="3789c-117">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="3789c-118">Devem implementar um sistema de autenticação alternativo que o suplemento possa se enquadrar em determinadas situações de erro.</span><span class="sxs-lookup"><span data-stu-id="3789c-118">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="3789c-119">Você pode usar um sistema de autenticação e tabelas de usuário ou utilizar um dos provedores de logon de redes sociais.</span><span class="sxs-lookup"><span data-stu-id="3789c-119">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="3789c-120">Para saber mais sobre como fazer isso com um suplemento do Office, confira [Autorizar serviços externos nos suplementos do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="3789c-120">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="3789c-121">Para *Outlook*, há um sistema de retorno recomendado.</span><span class="sxs-lookup"><span data-stu-id="3789c-121">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="3789c-122">Para mais informações, confira [Cenário: implementar o logon único no serviço em um Suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="3789c-122">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="3789c-123">Como o SSO funciona em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="3789c-123">How SSO works at runtime</span></span>

<span data-ttu-id="3789c-124">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="3789c-124">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="3789c-126">No suplemento, o JavaScript chama uma nova API Office.js [getAccessTokenAsync](#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="3789c-126">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference).</span></span> <span data-ttu-id="3789c-127">Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-127">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="3789c-128">Confira [Token de acesso de amostra](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="3789c-128">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="3789c-129">Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.</span><span class="sxs-lookup"><span data-stu-id="3789c-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="3789c-130">Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.</span><span class="sxs-lookup"><span data-stu-id="3789c-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="3789c-131">O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual. </span><span class="sxs-lookup"><span data-stu-id="3789c-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="3789c-132">O Azure AD envia o token do suplemento ao aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3789c-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="3789c-133">O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="3789c-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="3789c-134">O JavaScript no suplemento pode analisar o token e extrair informações necessárias, como endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="3789c-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="3789c-135">Opcionalmente, o suplemento pode enviar solicitação HTTP para o servidor para obter mais dados sobre o usuário; como as preferências do usuário.</span><span class="sxs-lookup"><span data-stu-id="3789c-135">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="3789c-136">Como alternativa, o próprio token de acesso pode ser enviado para o servidor para análise e validação.</span><span class="sxs-lookup"><span data-stu-id="3789c-136">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="3789c-137">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="3789c-137">Develop an SSO add-in</span></span>

<span data-ttu-id="3789c-138">Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO.</span><span class="sxs-lookup"><span data-stu-id="3789c-138">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="3789c-139">Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agnóstica.</span><span class="sxs-lookup"><span data-stu-id="3789c-139">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="3789c-140">Confira exemplos de explicações detalhadas em:</span><span class="sxs-lookup"><span data-stu-id="3789c-140">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="3789c-141">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="3789c-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="3789c-142">Criar um Suplemento do Office com ASP.NET que usa logon único</span><span class="sxs-lookup"><span data-stu-id="3789c-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="3789c-143">Criar o aplicativo de serviço</span><span class="sxs-lookup"><span data-stu-id="3789c-143">Create the service application</span></span>

<span data-ttu-id="3789c-144">Registre o suplemento no portal de registro para o ponto de extremidade v 2.0 do Azure:https://apps.dev.microsoft.com</span><span class="sxs-lookup"><span data-stu-id="3789c-144">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com.</span></span> <span data-ttu-id="3789c-145">Esse é um processo que leva entre 5 e 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="3789c-145">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="3789c-146">Obter um ID de cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="3789c-147">Especificar as permissões que seu suplemento precisa de AAD v.</span><span class="sxs-lookup"><span data-stu-id="3789c-147">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="3789c-148">ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="3789c-148">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="3789c-149">A permissão "perfil" sempre é necessária.</span><span class="sxs-lookup"><span data-stu-id="3789c-149">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="3789c-150">Conceder a confiança do aplicativo host do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="3789c-151">Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="3789c-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="3789c-152">Para mais detalhes sobre esse processo, confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="3789c-152">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="3789c-153">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3789c-153">Configure the add-in</span></span>

<span data-ttu-id="3789c-154">Adicione novas marcações ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="3789c-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="3789c-155">**WebApplicationInfo** – o pai dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="3789c-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="3789c-156">**ID** - O ID do cliente do suplemento Este é um ID do aplicativo que você obtém como parte do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-156">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="3789c-157">Confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="3789c-157">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="3789c-158">**Resource** – A URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-158">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="3789c-159">Esse é o mesmo URI (incluindo o protocolo `api:`) que você usou ao registrar o suplemento no AAD.</span><span class="sxs-lookup"><span data-stu-id="3789c-159">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="3789c-160">Parte de domínio deste URI deve coincidir com o domínio, incluindo qualquer subdomínio, que o usado nas URLs na seção `<Resources>` do manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-160">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="3789c-161">**Scopes** – O pai de uma ou mais elementos **Scope**.</span><span class="sxs-lookup"><span data-stu-id="3789c-161">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="3789c-162">**Scope** – Especifica uma permissão que seu suplemento precisa para o AAD.</span><span class="sxs-lookup"><span data-stu-id="3789c-162">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="3789c-163">A permissão `profile` sempre é necessária, e pode ser a única permissão necessária, se o suplemento não acessar o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3789c-163">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="3789c-164">Se isso acontecer, você também precisa de elementos **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="3789c-164">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="3789c-165">Bibliotecas que você usa no seu código para acessar o Microsoft Graph pode precisar de permissões adicionais.</span><span class="sxs-lookup"><span data-stu-id="3789c-165">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="3789c-166">Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="3789c-166">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="3789c-167">Para saber mais, confira [autorizar o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="3789c-167">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="3789c-p114">Para hosts do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="3789c-p114">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="3789c-170">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="3789c-170">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="3789c-171">Adicionar código do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="3789c-171">Add client-side code</span></span>

<span data-ttu-id="3789c-172">Adicione o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="3789c-172">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="3789c-173">Ligue [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="3789c-173">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="3789c-174">Analisar o token de acesso ou encaminhá-lo ao código de servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-174">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="3789c-175">Aqui está um exemplo simples de uma chamada para `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="3789c-175">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="3789c-176">Este exemplo lida explicitamente com apenas um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="3789c-176">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="3789c-177">Exemplos de tratamento de erro mais elaborados, confira [Home.js Office – adicionar-em-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js Office – adicionar-em-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="3789c-177">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="3789c-178">E veja [solucionar problemas de mensagens de erro no logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="3789c-178">And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

<span data-ttu-id="3789c-179">Aqui está um exemplo simples de como passar o token de suplemento para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="3789c-179">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="3789c-180">O token é incluído como um cabeçalho de `Authorization`ao enviar uma solicitação para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="3789c-180">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="3789c-181">Este exemplo prevê enviar dados JSON, para que ele tenha o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver escrevendo no servidor.</span><span class="sxs-lookup"><span data-stu-id="3789c-181">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="3789c-182">Quando chamar o método</span><span class="sxs-lookup"><span data-stu-id="3789c-182">When to call the method</span></span>

<span data-ttu-id="3789c-183">Se o seu suplemento não puder ser usado quando nenhum usuário estiver logado no Office, então você deve chamar `getAccessTokenAsync` \* quando o suplemento for iniciado \*.</span><span class="sxs-lookup"><span data-stu-id="3789c-183">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="3789c-184">Se o complemento tiver alguma funcionalidade que não exija um usuário conectado, então chame `getAccessTokenAsync` \* quando o usuário fizer uma ação que exija acesso a um usuário logado\*.</span><span class="sxs-lookup"><span data-stu-id="3789c-184">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="3789c-185">Não há uma degradação significativa do desempenho com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutilizará, até que ele expire, sem fazer outra chamada para o AAD V.</span><span class="sxs-lookup"><span data-stu-id="3789c-185">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="3789c-186">Ponto de extremidade 2.0 sempre que `getAccessTokenAsync` for chamado.</span><span class="sxs-lookup"><span data-stu-id="3789c-186">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="3789c-187">Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.</span><span class="sxs-lookup"><span data-stu-id="3789c-187">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="3789c-188">Adicionar código no lado do servidor</span><span class="sxs-lookup"><span data-stu-id="3789c-188">Add server-side code</span></span>

<span data-ttu-id="3789c-189">Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá.</span><span class="sxs-lookup"><span data-stu-id="3789c-189">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="3789c-190">Algumas tarefas de servidor que o suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="3789c-190">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="3789c-191">Criar um ou mais métodos de Web API com informações sobre o usuário que são extraídas do token; Por exemplo, uma forma que procura preferências do usuário em seu banco de dados hospedado.</span><span class="sxs-lookup"><span data-stu-id="3789c-191">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="3789c-192">(Confira **usando o token SSO, como uma identidade** abaixo.)Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código que você precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="3789c-192">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="3789c-193">Obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3789c-193">Get Microsoft Graph data.</span></span> <span data-ttu-id="3789c-194">O código do lado do servidor precisa fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="3789c-194">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="3789c-195">Validar o token de acesso (consulte **validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="3789c-195">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="3789c-196">Iniciar o fluxo "on behalf of" com uma chamada para o ponto de extremidade v 2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo).</span><span class="sxs-lookup"><span data-stu-id="3789c-196">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="3789c-197">O token de acesso nesse contexto é chamado de bootstrap token.</span><span class="sxs-lookup"><span data-stu-id="3789c-197">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="3789c-198">Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.</span><span class="sxs-lookup"><span data-stu-id="3789c-198">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="3789c-199">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="3789c-199">Get data from Microsoft Graph by using the new token.</span></span>

 <span data-ttu-id="3789c-200">Para saber mais sobre como obter acesso autorizado aos dados do usuário Microsoft Graph, confira [Autorizar o Microsoft Graph nos suplementos do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="3789c-200">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="3789c-201">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="3789c-201">Validate the access token</span></span>

<span data-ttu-id="3789c-202">Após a API Web receber o token de acesso, ela deve validá-lo antes que ele possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="3789c-202">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="3789c-203">O token é um Token Web JSON (JWT) e isso significa que validação funciona como uma validação de token na maioria dos fluxos padrão do OAuth.</span><span class="sxs-lookup"><span data-stu-id="3789c-203">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="3789c-204">Há diversas bibliotecas disponíveis que podem lidar com a validação de JWT. No entanto, as noções básicas incluem:</span><span class="sxs-lookup"><span data-stu-id="3789c-204">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="3789c-205">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="3789c-205">Checking that the token is well-formed</span></span>
- <span data-ttu-id="3789c-206">Verificar se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="3789c-206">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="3789c-207">Verificar se o token está direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="3789c-207">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="3789c-208">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="3789c-208">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="3789c-209">Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="3789c-209">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="3789c-210">A declaração `iss` no token deve começar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="3789c-210">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="3789c-211">O parâmetro `aud` do token será configurado como a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-211">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="3789c-212">O parâmetro `scp` do token será definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="3789c-212">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="3789c-213">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="3789c-213">Using the SSO token as an identity</span></span>

<span data-ttu-id="3789c-214">Se o suplemento precisar verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade.</span><span class="sxs-lookup"><span data-stu-id="3789c-214">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="3789c-215">As seguintes declarações no token estão relacionadas à identidade.</span><span class="sxs-lookup"><span data-stu-id="3789c-215">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="3789c-216">`name` – O nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="3789c-216">`name` - The user's display name.</span></span>
- <span data-ttu-id="3789c-217">`preferred_username`O endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="3789c-217">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="3789c-218">`oid` – Um GUID que representa a ID do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="3789c-218">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="3789c-219">`tid` – Um GUID que representa a ID da organização do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="3789c-219">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="3789c-220">Como os valores `name` e `preferred_username` podem alterar, recomendamos que os valores `oid` e `tid` sejam usados para correlacionar a identidade com o serviço de autorização do back-end.</span><span class="sxs-lookup"><span data-stu-id="3789c-220">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="3789c-221">Por exemplo, o serviço poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados do usuário interno.</span><span class="sxs-lookup"><span data-stu-id="3789c-221">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="3789c-222">Em seguida, nas solicitações subsequentes, o usuário poderia ser recuperado usando o mesmo valor e o acesso a recursos específicos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.</span><span class="sxs-lookup"><span data-stu-id="3789c-222">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="3789c-223">Token de acesso de exemplo</span><span class="sxs-lookup"><span data-stu-id="3789c-223">Example access token</span></span>

<span data-ttu-id="3789c-224">A seguir está uma carga decodificada típica do token de acesso.</span><span class="sxs-lookup"><span data-stu-id="3789c-224">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="3789c-225">Para saber mais sobre as propriedades, confira [Referência de tokens de versão do Azure Active Directory 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="3789c-225">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="3789c-226">Usando o SSO com um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="3789c-226">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="3789c-227">Há algumas diferenças pequenas, mas importantes entre usar o SSO em um suplemento do Outlook e em um suplemento do Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="3789c-227">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="3789c-228">Não deixe de ler [Autenticar o usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar o logon único ao serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="3789c-228">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="3789c-229">Referência da API do SSO</span><span class="sxs-lookup"><span data-stu-id="3789c-229">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="3789c-230">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3789c-230">getAccessTokenAsync</span></span>

<span data-ttu-id="3789c-231">O namespace de autenticação do Office `Office.context.auth`, fornece um método `getAccessTokenAsync` que permite que o host do Office obtenha um token de acesso para aplicativo web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-231">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="3789c-232">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="3789c-232">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="3789c-233">O método chama o ponto de extremidade do Azure Active Directory V 2.0 para obter um token de acesso para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3789c-233">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="3789c-234">Isso permite que os suplementos identifiquem usuários.</span><span class="sxs-lookup"><span data-stu-id="3789c-234">This enables add-ins to identify users.</span></span> <span data-ttu-id="3789c-235">O código do lado do servidor pode usar esse token para acessar o Microsoft Graph do aplicativo Web do suplemento usando o [fluxo OAuth "em nome de"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="3789c-235">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="3789c-236">No Outlook, não há suporte para esse API se o suplemento for carregado em uma caixa de correio do Gmail ou do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="3789c-236">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

<table><tr><td><span data-ttu-id="3789c-237">Hosts</span><span class="sxs-lookup"><span data-stu-id="3789c-237">Hosts</span></span></td><td><span data-ttu-id="3789c-238">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="3789c-238">Excel, OneNote, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td>[<span data-ttu-id="3789c-239">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="3789c-239">Requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[<span data-ttu-id="3789c-240">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="3789c-240">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="3789c-241">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="3789c-241">Parameters</span></span>

<span data-ttu-id="3789c-242">`options` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="3789c-242">`options` - Optional.</span></span> <span data-ttu-id="3789c-243">Aceitar um objeto `AuthOptions` (veja abaixo) para definir comportamentos logon.</span><span class="sxs-lookup"><span data-stu-id="3789c-243">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="3789c-244">`callback` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="3789c-244">`callback` - Optional.</span></span> <span data-ttu-id="3789c-245">Aceita um método de retorno que possa analisar o token de ID de usuário ou usar o token fluxo "em nome de" para obter acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3789c-245">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="3789c-246">Se [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult) `.status` é "bem-sucedido", em seguida, `AsyncResult.value` é o v AAD bruto.</span><span class="sxs-lookup"><span data-stu-id="3789c-246">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="3789c-247">token de acesso 2.0 formatado.</span><span class="sxs-lookup"><span data-stu-id="3789c-247">2.0-formatted access token.</span></span>

<span data-ttu-id="3789c-248">A interface `AuthOptions`fornece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v.</span><span class="sxs-lookup"><span data-stu-id="3789c-248">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="3789c-249">2.0 com o método`getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="3789c-249">2.0 with the `getAccessTokenAsync` method.</span></span>

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



