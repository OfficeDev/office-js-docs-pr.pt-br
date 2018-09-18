---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 534ac41e7518756a2aa5b4408ce7adb0f434e27d
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945338"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="c5d47-102">Habilitar o logon único para Suplementos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="c5d47-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="c5d47-103">Os usuários entram no Office (online, em dispositivos móveis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365).</span><span class="sxs-lookup"><span data-stu-id="c5d47-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="c5d47-104">Você pode aproveitar isso e usar o logon único (SSO) para autorizar o usuário ao seu suplemento sem exigir que o usuário faça login uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="c5d47-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="c5d47-106">Atualmente a API de logon único tem suporte para Word, Excel e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c5d47-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="c5d47-107">Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="c5d47-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>
> <span data-ttu-id="c5d47-108">Para usar SSO, você precisa carregar a versão beta da Biblioteca JavaScript para Office de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página HTML de inicialização do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-108">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="c5d47-109">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365.</span><span class="sxs-lookup"><span data-stu-id="c5d47-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="c5d47-110">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="c5d47-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="c5d47-111">Para os usuários, isso torna a experiência de execução do suplemento mais fácil com apenas um único logon.</span><span class="sxs-lookup"><span data-stu-id="c5d47-111">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="c5d47-112">Para os desenvolvedores, isso significa que o suplemento não precisa manter suas próprias tabelas de usuário com senhas criptografadas.</span><span class="sxs-lookup"><span data-stu-id="c5d47-112">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="c5d47-113">Como ele funciona em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="c5d47-113">How it works at runtime</span></span>

<span data-ttu-id="c5d47-114">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="c5d47-114">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="c5d47-p104">No suplemento, o JavaScript chama uma nova API Office.js[getAccessTokenAsync](#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. Consulte [Exemplo de token de acesso](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="c5d47-p104">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. (Hereafter, this is called the [add-in token](#example-access-token).)</span></span>
2. <span data-ttu-id="c5d47-119">Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.</span><span class="sxs-lookup"><span data-stu-id="c5d47-119">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="c5d47-120">Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-120">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="c5d47-121">O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="c5d47-121">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="c5d47-122">O Azure AD envia o token do suplemento ao aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="c5d47-122">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="c5d47-123">O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-123">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="c5d47-124">O JavaScript no suplemento pode analisar o token e extrair as informações necessárias, como o endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5d47-124">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="c5d47-125">Opcionalmente, o suplemento pode enviar uma solicitação HTTP para o servidor para obter mais dados sobre o usuário, tais como as preferências do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5d47-125">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="c5d47-126">Ou então, o próprio token de acesso pode ser enviado para o servidor para análise e validação.</span><span class="sxs-lookup"><span data-stu-id="c5d47-126">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="c5d47-127">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="c5d47-127">Develop an SSO add-in</span></span>

<span data-ttu-id="c5d47-128">Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO.</span><span class="sxs-lookup"><span data-stu-id="c5d47-128">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="c5d47-129">Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agnóstica.</span><span class="sxs-lookup"><span data-stu-id="c5d47-129">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="c5d47-130">Confira exemplos de explicações detalhadas em:</span><span class="sxs-lookup"><span data-stu-id="c5d47-130">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="c5d47-131">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="c5d47-131">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="c5d47-132">Criar um Suplemento do Office com ASP.NET que usa logon único</span><span class="sxs-lookup"><span data-stu-id="c5d47-132">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="c5d47-133">Criar o aplicativo de serviço</span><span class="sxs-lookup"><span data-stu-id="c5d47-133">Create the service application</span></span>

<span data-ttu-id="c5d47-134">Registrar o suplemento no portal de registro para o ponto de extremidade do Azure v 2.0: https://apps.dev.microsoft.com.</span><span class="sxs-lookup"><span data-stu-id="c5d47-134">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="c5d47-135">Esse processo leva de 5 a 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="c5d47-135">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="c5d47-136">Obter uma ID de cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-136">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="c5d47-137">Especificar as permissões que o suplemento precisa para o AAD v.</span><span class="sxs-lookup"><span data-stu-id="c5d47-137">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="c5d47-138">Ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="c5d47-138">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="c5d47-139">A permissão "perfil" é sempre necessária.</span><span class="sxs-lookup"><span data-stu-id="c5d47-139">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="c5d47-140">Conceder a confiança do aplicativo host do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-140">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="c5d47-141">Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="c5d47-141">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="c5d47-142">Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="c5d47-142">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="c5d47-143">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="c5d47-143">Configure the add-in</span></span>

<span data-ttu-id="c5d47-144">Adicione novas marcações ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="c5d47-144">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="c5d47-145">**WebApplicationInfo** – o pai dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="c5d47-145">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="c5d47-146">**Id** - A ID do cliente do suplemento. Esta é uma ID do aplicativo que você obtém como parte do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-146">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="c5d47-147">Veja [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do D do Azure v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="c5d47-147">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="c5d47-148">**Resource** – A URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-148">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="c5d47-149">**Scopes** – O pai de uma ou mais elementos **Scope**.</span><span class="sxs-lookup"><span data-stu-id="c5d47-149">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="c5d47-150">**Scope** – Especifica uma permissão que o suplemento precisa para o AAD.</span><span class="sxs-lookup"><span data-stu-id="c5d47-150">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="c5d47-151">A permissão `profile` é sempre necessária e pode ser a única permissão necessária, se seu suplemento não acessar o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c5d47-151">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="c5d47-152">Se isso acontecer, você também precisará dos elementos do **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-152">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="c5d47-153">Bibliotecas que você usa no seu código para acessar o Microsoft Graph podem precisar de permissões adicionais.</span><span class="sxs-lookup"><span data-stu-id="c5d47-153">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="c5d47-154">Por exemplo, a Microsoft Authentication Library (MSAL) para .NET requer a permissão `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-154">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="c5d47-155">Para mais informações, veja [Autorizar para o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="c5d47-155">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="c5d47-p111">Para hosts do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-p111">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="c5d47-158">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="c5d47-158">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="c5d47-159">Adicionar código do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="c5d47-159">Add client-side code</span></span>

<span data-ttu-id="c5d47-160">Adicione o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="c5d47-160">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="c5d47-161">Chame [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="c5d47-161">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="c5d47-162">Analisar o token de acesso ou passá-lo para o código do servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-162">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="c5d47-163">Aqui está um exemplo simples de uma chamada para `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-163">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="c5d47-164">Este exemplo manipula apenas um tipo de erro explicitamente.</span><span class="sxs-lookup"><span data-stu-id="c5d47-164">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="c5d47-165">Para exemplos de manipulação de erro mais elaborados, veja [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="c5d47-165">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="c5d47-166">E veja [Solucionar problemas de mensagens de erro para logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="c5d47-166">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="c5d47-167">Aqui está um exemplo simples de passagem do token do suplemento para o servidor.</span><span class="sxs-lookup"><span data-stu-id="c5d47-167">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="c5d47-168">O token é incluído como um cabeçalho `Authorization` ao enviar uma solicitação de volta para o servidor.</span><span class="sxs-lookup"><span data-stu-id="c5d47-168">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="c5d47-169">Este exemplo prevê o envio de dados JSON e, portanto, ele usa o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver gravando no servidor.</span><span class="sxs-lookup"><span data-stu-id="c5d47-169">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="c5d47-170">Quando chamar o método</span><span class="sxs-lookup"><span data-stu-id="c5d47-170">When to call the method</span></span>

<span data-ttu-id="c5d47-171">Se o seu suplemento não puder ser usado quando nenhum usuário está conectado ao Office, chame `getAccessTokenAsync` *quando o suplemento for iniciado*.</span><span class="sxs-lookup"><span data-stu-id="c5d47-171">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="c5d47-172">Se o suplemento tiver alguma funcionalidade que não exija um usuário conectado, você poderá chamar `getAccessTokenAsync` *quando o usuário realizar uma ação que exija um usuário conectado*.</span><span class="sxs-lookup"><span data-stu-id="c5d47-172">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="c5d47-173">Não há uma degradação do desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutiliza até que ele expire, sem fazer outra chamada para o AAD v.</span><span class="sxs-lookup"><span data-stu-id="c5d47-173">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="c5d47-174">sempre que o `getAccessTokenAsync` for chamado.</span><span class="sxs-lookup"><span data-stu-id="c5d47-174">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="c5d47-175">Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.</span><span class="sxs-lookup"><span data-stu-id="c5d47-175">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="c5d47-176">Adicionar código no lado do servidor</span><span class="sxs-lookup"><span data-stu-id="c5d47-176">Add server-side code</span></span>

<span data-ttu-id="c5d47-177">Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá.</span><span class="sxs-lookup"><span data-stu-id="c5d47-177">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="c5d47-178">Algumas tarefas do servidor que seu suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="c5d47-178">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="c5d47-179">Criar um ou mais métodos da API da Web que usem informações sobre o usuário extraído do token; por exemplo, um método que procura as preferências do usuário em sua base de dados hospedada.</span><span class="sxs-lookup"><span data-stu-id="c5d47-179">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="c5d47-180">(Veja **Usar o token SSO como uma identidade** abaixo.) Dependendo do seu idioma e estrutura, as bibliotecas podem estar disponíveis para simplificar o código que você precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="c5d47-180">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="c5d47-181">Obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c5d47-181">Get Microsoft Graph data.</span></span> <span data-ttu-id="c5d47-182">O código do lado do servidor precisa fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="c5d47-182">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="c5d47-183">Validar o token de acesso (veja **Validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="c5d47-183">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="c5d47-184">Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade do AD do Azure  v2.0 que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo).</span><span class="sxs-lookup"><span data-stu-id="c5d47-184">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="c5d47-185">Nesse contexto, o token de acesso é chamado de token de inicialização.</span><span class="sxs-lookup"><span data-stu-id="c5d47-185">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="c5d47-186">Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.</span><span class="sxs-lookup"><span data-stu-id="c5d47-186">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="c5d47-187">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="c5d47-187">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="c5d47-188">Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usuário, veja [Autorizar para o Microsoft Graph no seu Suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="c5d47-188">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="c5d47-189">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="c5d47-189">For more information, see Validate the access token.</span></span>

<span data-ttu-id="c5d47-190">Após a API Web receber o token de acesso, ela deve validá-lo antes que ele possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="c5d47-190">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="c5d47-191">O token é um Token Web JSON (JWT) e isso significa que validação funciona como uma validação de token na maioria dos fluxos padrão do OAuth.</span><span class="sxs-lookup"><span data-stu-id="c5d47-191">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="c5d47-192">Há diversas bibliotecas disponíveis que podem lidar com a validação de JWT. No entanto, as noções básicas incluem:</span><span class="sxs-lookup"><span data-stu-id="c5d47-192">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="c5d47-193">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="c5d47-193">Checking that the token is well-formed</span></span>
- <span data-ttu-id="c5d47-194">Verificar se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="c5d47-194">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="c5d47-195">Verificar se o token está direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="c5d47-195">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="c5d47-196">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="c5d47-196">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="c5d47-197">Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-197">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="c5d47-198">A declaração `iss` no token deve começar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="c5d47-198">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="c5d47-199">O parâmetro `aud` do token será configurado como a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-199">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="c5d47-200">O parâmetro `scp` do token será definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-200">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="c5d47-201">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="c5d47-201">Using the SSO token as an identity</span></span>

<span data-ttu-id="c5d47-202">Se o suplemento precisar verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade.</span><span class="sxs-lookup"><span data-stu-id="c5d47-202">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="c5d47-203">As seguintes declarações no token estão relacionadas à identidade.</span><span class="sxs-lookup"><span data-stu-id="c5d47-203">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="c5d47-204">`name` – O nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5d47-204">`name` - The user's display name.</span></span>
- <span data-ttu-id="c5d47-205">`preferred_username` - O endereço de e-mail do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5d47-205">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="c5d47-206">`oid` – Um GUID que representa a ID do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="c5d47-206">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="c5d47-207">`tid` – Um GUID que representa a ID da organização do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="c5d47-207">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="c5d47-208">Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ​​para correlacionar a identidade com o serviço de autorização do seu back-end.</span><span class="sxs-lookup"><span data-stu-id="c5d47-208">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="c5d47-209">Por exemplo, o serviço poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados do usuário interno.</span><span class="sxs-lookup"><span data-stu-id="c5d47-209">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="c5d47-210">Em seguida, nas solicitações subsequentes, o usuário poderia ser recuperado usando o mesmo valor e o acesso a recursos específicos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.</span><span class="sxs-lookup"><span data-stu-id="c5d47-210">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="c5d47-211">Exemplo de token de acesso</span><span class="sxs-lookup"><span data-stu-id="c5d47-211">Example access token</span></span>

<span data-ttu-id="c5d47-212">A seguir, um conteúdo decodificado típico de um token de acesso.</span><span class="sxs-lookup"><span data-stu-id="c5d47-212">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="c5d47-213">Para mais informações sobre as propriedades, veja [Referência de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="c5d47-213">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="c5d47-214">Usar o SSO com um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="c5d47-214">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="c5d47-p124">Existem algumas diferenças pequenas, mas importantes, entre usar o SSO em um suplemento do Outlook e usá-lo em um suplemento do Excel, PowerPoint ou Word. Leia [Autenticar um usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="c5d47-p124">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="c5d47-217">Referência da API de SSO</span><span class="sxs-lookup"><span data-stu-id="c5d47-217">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="c5d47-218">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c5d47-218">getAccessTokenAsync</span></span>

<span data-ttu-id="c5d47-219">O namespace Office Auth, `Office.context.auth`, fornece um método `getAccessTokenAsync` que permite que o host do Office obtenha um token de acesso para o aplicativo da Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-219">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="c5d47-220">Indiretamente, isso também permite que o suplemento acesse os dados do Microsoft Graph do usuário conectado sem exigir que o usuário entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="c5d47-220">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="c5d47-221">O método chama o ponto de extremidade do Active Directory do Azure V 2.0 para obter um token de acesso para o aplicativo da Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5d47-221">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="c5d47-222">Isso permite que os suplementos identifiquem usuários.</span><span class="sxs-lookup"><span data-stu-id="c5d47-222">This enables add-ins to identify users.</span></span> <span data-ttu-id="c5d47-223">O código do lado do servidor pode usar esse token para acessar o Microsoft Graph do aplicativo Web do suplemento usando o [fluxo OAuth "em nome de"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="c5d47-223">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="c5d47-224">No Outlook, essa API não é suportada se o suplemento for carregado em uma caixa de correio do Outlook.com ou do Gmail.</span><span class="sxs-lookup"><span data-stu-id="c5d47-224">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

<table><tr><td><span data-ttu-id="c5d47-225">Hosts</span><span class="sxs-lookup"><span data-stu-id="c5d47-225">Hosts</span></span></td><td><span data-ttu-id="c5d47-226">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="c5d47-226">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="c5d47-227">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="c5d47-227">Requirement sets</span></span></td><td>[<span data-ttu-id="c5d47-228">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="c5d47-228">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="c5d47-229">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c5d47-229">Parameters</span></span>

<span data-ttu-id="c5d47-230">`options` - Opcional.</span><span class="sxs-lookup"><span data-stu-id="c5d47-230">`options` - Optional.</span></span> <span data-ttu-id="c5d47-231">Aceita um objeto `AuthOptions` (veja abaixo) para definir os comportamentos de logon.</span><span class="sxs-lookup"><span data-stu-id="c5d47-231">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="c5d47-232">`callback` - Opcional.</span><span class="sxs-lookup"><span data-stu-id="c5d47-232">`callback` - Optional.</span></span> <span data-ttu-id="c5d47-233">Aceita um método de retorno de chamada que pode analisar o token para a ID do usuário ou usar o token no fluxo de "em nome de" para obter acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c5d47-233">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="c5d47-234">Se [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` for "succeeded", então `AsyncResult.value` é o AAD v não processado.</span><span class="sxs-lookup"><span data-stu-id="c5d47-234">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="c5d47-235">Token de acesso formatado para 2.0.</span><span class="sxs-lookup"><span data-stu-id="c5d47-235">2.0-formatted access token.</span></span>

<span data-ttu-id="c5d47-236">A interface `AuthOptions` oferece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v.</span><span class="sxs-lookup"><span data-stu-id="c5d47-236">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="c5d47-237">2.0 com o método `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="c5d47-237">2.0 with the `getAccessTokenAsync` method.</span></span>

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



