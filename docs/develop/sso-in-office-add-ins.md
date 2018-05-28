---
title: Habilitar o logon ?nico para Suplementos do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 45bd63150ffa8e46bf9c0fa54711ac907b8490ce
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="5e03d-102">Habilitar o logon ?nico para Suplementos do Office (visualiza??o)</span><span class="sxs-lookup"><span data-stu-id="5e03d-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="5e03d-103">Os usu?rios entram no Office (online, em dispositivos m?veis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365).</span><span class="sxs-lookup"><span data-stu-id="5e03d-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="5e03d-104">Voc? pode aproveitar isso e usar o logon ?nico (SSO) para autorizar o usu?rio ao seu suplemento sem exigir que o usu?rio fa?a login uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="5e03d-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="5e03d-106">Atualmente a API de logon ?nico tem suporte para Word, Excel e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5e03d-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="5e03d-107">Confira mais informa??es sobre os programas para os quais a API de logon ?nico tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="5e03d-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span></span>
> <span data-ttu-id="5e03d-108">Se voc? estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autentica??o Moderna para a loca??o do Office 365.</span><span class="sxs-lookup"><span data-stu-id="5e03d-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="5e03d-109">Confira mais informa??es sobre como fazer isso em [Exchange Online: como habilitar seu locat?rio para autentica??o moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="5e03d-109">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="5e03d-110">Para os usu?rios, isso torna a experi?ncia de execu??o do suplemento mais f?cil com apenas um ?nico logon.</span><span class="sxs-lookup"><span data-stu-id="5e03d-110">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="5e03d-111">Para os desenvolvedores, isso significa que o suplemento n?o precisa manter suas pr?prias tabelas de usu?rio com senhas criptografadas.</span><span class="sxs-lookup"><span data-stu-id="5e03d-111">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="5e03d-112">Como ele funciona em tempo de execu??o</span><span class="sxs-lookup"><span data-stu-id="5e03d-112">How it works at runtime</span></span>

<span data-ttu-id="5e03d-113">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="5e03d-113">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="5e03d-115">No suplemento, o JavaScript chama uma nova API Office.js `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-115">In the add-in, JavaScript calls a new Office.js API `getAccessTokenAsync`.</span></span> <span data-ttu-id="5e03d-116">Isso notifica o aplicativo host do Office para que obtenha um token de acesso para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-116">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="5e03d-117">Veja [Exemplo de token de acesso](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="5e03d-117">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="5e03d-118">Se o usu?rio n?o estiver conectado, o aplicativo host do Office abrir? uma janela pop-up para o usu?rio entrar.</span><span class="sxs-lookup"><span data-stu-id="5e03d-118">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="5e03d-119">Se essa ? a primeira vez que o usu?rio atual usa seu suplemento, ser? solicitado que ele d? o consentimento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-119">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="5e03d-120">O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usu?rio atual.</span><span class="sxs-lookup"><span data-stu-id="5e03d-120">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="5e03d-121">O Azure AD envia o token do suplemento ao aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="5e03d-121">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="5e03d-122">O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-122">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="5e03d-123">O JavaScript no suplemento pode analisar o token e extrair as informa??es necess?rias, como o endere?o de email do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5e03d-123">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="5e03d-124">Opcionalmente, o suplemento pode enviar uma solicita??o HTTP para o servidor para obter mais dados sobre o usu?rio, tais como as prefer?ncias do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5e03d-124">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="5e03d-125">Ou ent?o, o pr?prio token de acesso pode ser enviado para o servidor para an?lise e valida??o.</span><span class="sxs-lookup"><span data-stu-id="5e03d-125">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="5e03d-126">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="5e03d-126">Develop an SSO add-in</span></span>

<span data-ttu-id="5e03d-127">Esta se??o descreve as tarefas envolvidas na cria??o de um suplemento do Office que usa SSO.</span><span class="sxs-lookup"><span data-stu-id="5e03d-127">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="5e03d-128">Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agn?stica.</span><span class="sxs-lookup"><span data-stu-id="5e03d-128">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="5e03d-129">Confira exemplos de explica??es detalhadas em:</span><span class="sxs-lookup"><span data-stu-id="5e03d-129">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="5e03d-130">Criar um Suplemento do Office com Node.js que usa logon ?nico</span><span class="sxs-lookup"><span data-stu-id="5e03d-130">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="5e03d-131">Criar um Suplemento do Office com ASP.NET que usa logon ?nico</span><span class="sxs-lookup"><span data-stu-id="5e03d-131">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="5e03d-132">Criar o aplicativo de servi?o</span><span class="sxs-lookup"><span data-stu-id="5e03d-132">Create the service application</span></span>

<span data-ttu-id="5e03d-133">Registre o suplemento no portal de registro para o ponto de extremidade v2.0 do Azure: https://apps.dev.microsoft.com. Esse ? um processo que leva de 5 a 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="5e03d-133">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5?10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="5e03d-134">Obter uma ID do cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-134">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="5e03d-135">Especificar as permiss?es que seu suplemento precisa para o AAD v.</span><span class="sxs-lookup"><span data-stu-id="5e03d-135">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="5e03d-136">Ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="5e03d-136">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="5e03d-137">A permiss?o "perfil" ? sempre necess?ria.</span><span class="sxs-lookup"><span data-stu-id="5e03d-137">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="5e03d-138">Conceder a rela??o de confian?a do aplicativo host do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-138">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="5e03d-139">Autorizar previamente o aplicativo host do Office para o suplemento com a permiss?o padr?o *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="5e03d-139">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="5e03d-140">Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="5e03d-140">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="5e03d-141">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="5e03d-141">Configure the add-in</span></span>

<span data-ttu-id="5e03d-142">Adicione novas marca??es ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="5e03d-142">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="5e03d-143">**WebApplicationInfo** ? O respons?vel dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="5e03d-143">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="5e03d-144">**Id** - A ID do cliente do suplemento. Esta ? uma ID do aplicativo que voc? obt?m como parte do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-144">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="5e03d-145">Veja [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="5e03d-145">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="5e03d-146">**Recurso** ? O URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-146">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="5e03d-147">**Escopos** ? O respons?vel de um ou mais elementos de **Escopo**.</span><span class="sxs-lookup"><span data-stu-id="5e03d-147">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="5e03d-148">**Escopo** ? Especifica uma permiss?o que o suplemento precisa para o AAD.</span><span class="sxs-lookup"><span data-stu-id="5e03d-148">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="5e03d-149">A permiss?o `profile` ? sempre necess?ria e pode ser a ?nica permiss?o necess?ria, se seu suplemento n?o acessar o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5e03d-149">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="5e03d-150">Se isso acontecer, voc? tamb?m precisar? dos elementos do **Escopo** para as permiss?es necess?rias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-150">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="5e03d-151">Bibliotecas que voc? usa no seu c?digo para acessar o Microsoft Graph podem precisar de permiss?es adicionais.</span><span class="sxs-lookup"><span data-stu-id="5e03d-151">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="5e03d-152">Por exemplo, a Microsoft Authentication Library (MSAL) para .NET requer a permiss?o `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-152">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="5e03d-153">Para mais informa??es, veja [Autorizar para o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="5e03d-153">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="5e03d-p110">Para hosts do Office diferentes do Outlook, adicione a marca??o no final da se??o `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marca??o no final da se??o `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-p110">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="5e03d-156">Veja a seguir um exemplo da marca??o:</span><span class="sxs-lookup"><span data-stu-id="5e03d-156">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="5e03d-157">Adicionar c?digo do cliente</span><span class="sxs-lookup"><span data-stu-id="5e03d-157">Add client-side code</span></span>

<span data-ttu-id="5e03d-158">Adicionar o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="5e03d-158">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="5e03d-159">Chamar [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span><span class="sxs-lookup"><span data-stu-id="5e03d-159">Call [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span></span>
* <span data-ttu-id="5e03d-160">Analisar o token de acesso ou pass?-lo para o c?digo do servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-160">Parse the access token or pass it to the add-in?s server-side code.</span></span> 

<span data-ttu-id="5e03d-161">Aqui est? um exemplo simples de uma chamada para `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-161">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!Note]
> <span data-ttu-id="5e03d-162">Este exemplo manipula apenas um tipo de erro explicitamente.</span><span class="sxs-lookup"><span data-stu-id="5e03d-162">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="5e03d-163">Para exemplos de manipula??o de erro mais elaborados, veja [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="5e03d-163">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="5e03d-164">E veja [Solucionar problemas de mensagens de erro no logon ?nico (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5e03d-164">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="5e03d-165">Aqui est? um exemplo simples de passagem do token do suplemento para o servidor.</span><span class="sxs-lookup"><span data-stu-id="5e03d-165">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="5e03d-166">O token ? inclu?do como um cabe?alho `Authorization` ao enviar uma solicita??o de volta para o servidor.</span><span class="sxs-lookup"><span data-stu-id="5e03d-166">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="5e03d-167">Este exemplo prev? o envio de dados JSON e, portanto, ele usa o m?todo `POST`, mas `GET` ? suficiente para enviar o token de acesso quando voc? n?o estiver gravando no servidor.</span><span class="sxs-lookup"><span data-stu-id="5e03d-167">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="5e03d-168">Quando chamar o m?todo</span><span class="sxs-lookup"><span data-stu-id="5e03d-168">When to call the method</span></span>

<span data-ttu-id="5e03d-169">Se o seu suplemento n?o puder ser usado quando nenhum usu?rio estiver conectado no Office, voc? dever? chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.</span><span class="sxs-lookup"><span data-stu-id="5e03d-169">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="5e03d-170">Se o suplemento tiver alguma funcionalidade que n?o exija um usu?rio conectado, voc? poder? chamar `getAccessTokenAsync` *quando o usu?rio realizar uma a??o que exija um usu?rio conectado*.</span><span class="sxs-lookup"><span data-stu-id="5e03d-170">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="5e03d-171">N?o h? uma degrada??o do desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutiliza at? que ele expire, sem fazer outra chamada para o ponto de extremidade v2.0 do AAD</span><span class="sxs-lookup"><span data-stu-id="5e03d-171">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="5e03d-172">sempre que o `getAccessTokenAsync` for chamado.</span><span class="sxs-lookup"><span data-stu-id="5e03d-172">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="5e03d-173">Portanto, voc? pode adicionar chamadas de `getAccessTokenAsync` para todas as fun??es e manipuladores que iniciam uma a??o onde o token ? necess?rio.</span><span class="sxs-lookup"><span data-stu-id="5e03d-173">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="5e03d-174">Adicionar c?digo do servidor</span><span class="sxs-lookup"><span data-stu-id="5e03d-174">Add server-side code</span></span>

<span data-ttu-id="5e03d-175">Na maioria dos cen?rios, n?o haver? muitas raz?es para obter o token de acesso, se o suplemento n?o o passar no lado do servidor e o utilizar l?.</span><span class="sxs-lookup"><span data-stu-id="5e03d-175">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="5e03d-176">Algumas tarefas do servidor que seu suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="5e03d-176">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="5e03d-177">Criar um ou mais m?todos da API da Web que usem informa??es sobre o usu?rio extra?do do token; por exemplo, um m?todo que procura as prefer?ncias do usu?rio em sua base de dados hospedada.</span><span class="sxs-lookup"><span data-stu-id="5e03d-177">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="5e03d-178">(Veja **Usar o token SSO como uma identidade** abaixo.) Dependendo do seu idioma e estrutura, as bibliotecas podem estar dispon?veis para simplificar o c?digo que voc? precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="5e03d-178">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="5e03d-179">Obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5e03d-179">Get Microsoft Graph data.</span></span> <span data-ttu-id="5e03d-180">O c?digo do servidor precisa fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="5e03d-180">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="5e03d-181">Validar o token de acesso (veja **Validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="5e03d-181">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="5e03d-182">Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usu?rio e as credenciais do suplemento (sua ID e segredo).</span><span class="sxs-lookup"><span data-stu-id="5e03d-182">Initiate the ?on behalf of? flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="5e03d-183">Nesse contexto, o token de acesso ? chamado de token de inicializa??o.</span><span class="sxs-lookup"><span data-stu-id="5e03d-183">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="5e03d-184">Armazenar em cache o novo token de acesso que ? retornado do fluxo em nome de.</span><span class="sxs-lookup"><span data-stu-id="5e03d-184">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="5e03d-185">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="5e03d-185">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="5e03d-186">Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usu?rio, veja [Autorizar para o Microsoft Graph no seu Suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="5e03d-186">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="5e03d-187">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="5e03d-187">Validate the token</span></span>

<span data-ttu-id="5e03d-188">Ap?s a API Web receber o token de acesso, ela deve valid?-lo antes que ele possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="5e03d-188">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="5e03d-189">O token ? um Token Web JSON (JWT) e isso significa que valida??o funciona como uma valida??o de token na maioria dos fluxos padr?o do OAuth.</span><span class="sxs-lookup"><span data-stu-id="5e03d-189">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="5e03d-190">H? diversas bibliotecas dispon?veis que podem lidar com a valida??o de JWT. No entanto, as no??es b?sicas incluem:</span><span class="sxs-lookup"><span data-stu-id="5e03d-190">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="5e03d-191">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="5e03d-191">Checking that the token is well-formed</span></span>
- <span data-ttu-id="5e03d-192">Verificar se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="5e03d-192">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="5e03d-193">Verificar se o token est? direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="5e03d-193">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="5e03d-194">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="5e03d-194">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="5e03d-195">Os tokens SSO v?lidos ser?o emitidos pela autoridade do Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-195">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="5e03d-196">A declara??o `iss` no token deve come?ar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="5e03d-196">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="5e03d-197">O par?metro `aud` do token ser? configurado como a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e03d-197">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="5e03d-198">O par?metro `scp` do token ser? definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="5e03d-198">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="5e03d-199">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="5e03d-199">Using the SSO token as an identity</span></span>

<span data-ttu-id="5e03d-200">Se o suplemento precisar verificar a identidade do usu?rio, o token SSO cont?m informa??es que podem ser usadas para estabelecer a identidade.</span><span class="sxs-lookup"><span data-stu-id="5e03d-200">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="5e03d-201">As seguintes declara??es no token est?o relacionadas ? identidade.</span><span class="sxs-lookup"><span data-stu-id="5e03d-201">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="5e03d-202">`name` ? O nome para exibi??o do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5e03d-202">`name` - The user's display name.</span></span>
- <span data-ttu-id="5e03d-203">`preferred_username` O endere?o de email do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5e03d-203">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="5e03d-204">`oid` ? Um GUID que representa a ID do usu?rio no Active Directory do Azure.</span><span class="sxs-lookup"><span data-stu-id="5e03d-204">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="5e03d-205">`tid` ? Um GUID que representa a ID da organiza??o do usu?rio no Active Directory do Azure.</span><span class="sxs-lookup"><span data-stu-id="5e03d-205">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="5e03d-206">Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ??para correlacionar a identidade com o servi?o de autoriza??o do back-end.</span><span class="sxs-lookup"><span data-stu-id="5e03d-206">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="5e03d-207">Por exemplo, o servi?o poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazen?-los como um valor no registro do usu?rio no banco de dados do usu?rio interno.</span><span class="sxs-lookup"><span data-stu-id="5e03d-207">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="5e03d-208">Em seguida, nas solicita??es subsequentes, o usu?rio poderia ser recuperado usando o mesmo valor e o acesso a recursos espec?ficos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.</span><span class="sxs-lookup"><span data-stu-id="5e03d-208">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="5e03d-209">Exemplo de token de acesso</span><span class="sxs-lookup"><span data-stu-id="5e03d-209">Example access token</span></span>

<span data-ttu-id="5e03d-210">A seguir, um conte?do decodificado t?pico de um token de acesso.</span><span class="sxs-lookup"><span data-stu-id="5e03d-210">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="5e03d-211">Para mais informa??es sobre as propriedades, veja [Refer?ncia de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="5e03d-211">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-and-outlook-add-in"></a><span data-ttu-id="5e03d-212">Usar o SSO com o suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="5e03d-212">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="5e03d-213">Existem algumas diferen?as pequenas, mas importantes, no uso do SSO com o suplemento do Outlook para us?-lo como suplemento do Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="5e03d-213">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="5e03d-214">Certifique-se de ler [Autenticar um usu?rio com um token de logon ?nico em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cen?rio: implementar o logon ?nico no servi?o em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="5e03d-214">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>