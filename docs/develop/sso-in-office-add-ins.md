---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: fb4eacee9419339116e15ef3fccc03b291faf3ec
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506025"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="bf261-102">Habilitar o logon único para Suplementos do Office (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="bf261-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="bf261-p101">Os usuários entram no Office (plataformas online, móveis e desktop) usando uma conta pessoal da Microsoft ou contas do trabalho ou da escola (Office 365). Você pode aproveitar isso e usar o logon único (SSO) para autorizar que o usuário use o seu suplemento sem exigir que ele entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="bf261-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="bf261-106">Status da versão prévia</span><span class="sxs-lookup"><span data-stu-id="bf261-106">Preview Status</span></span>

<span data-ttu-id="bf261-p102">A API de logon único é suportada somente no modo de visualização neste momento. Ela está disponível para experimentação dos desenvolvedores; mas não deve ser usada em um suplemento de produção. Além disso, os suplementos que usam SSO não são aceitos no [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="bf261-p102">The Single Sign-on API is currently supported in preview only. It is available to developers for experimentation; but it should not be used in a production add-in. In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="bf261-p103">Nem todos os aplicativos do Office oferecem suporte para versão prévia do SSO. Ele está disponível no Word, Excel, Outlook e PowerPoint. Para obter mais informações sobre onde a API de logon único é suportada no momento, confira [Conjuntos de requisitos IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="bf261-p103">Not all Office applications support the SSO preview. It is available in Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="bf261-113">Requisitos e melhores práticas</span><span class="sxs-lookup"><span data-stu-id="bf261-113">Requirements and Best Practices</span></span>

<span data-ttu-id="bf261-114">Para usar o SSO, você deve carregar a versão beta da Biblioteca JavaScript do Office em `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` na página HTML de inicialização do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="bf261-p104">Se estiver usando um suplemento do **Outlook** , você deve habilitar a autenticação moderna para os locatários do Office 365. Para obter mais informações sobre isso, confira  [Exchange Online: como habilitar o seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="bf261-p104">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="bf261-p105">Você *não* deve depender do SSO como único método de autenticação do seu suplemento. Você deve implementar um sistema alternativo de autenticação ao qual seu suplemento possa recorrer em determinadas situações de erro. Você pode usar um sistema de autenticação e de tabelas de usuário, ou você pode aproveitar um dos provedores de logon social. Para mais informações sobre como fazer isso com um suplemento do Office, confira [Autorizar serviços externos no seu suplemento do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Para o *Outlook*, existe um sistema alternativo recomendado. Para mais informações, confira [Cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="bf261-p105">You should *not* rely on SSO as your add-in's only method of authentication. You should implement an alternate authentication system that your add-in can fall back to in certain error situations. You can use a system of user tables and authentication, or you can leverage one of the social login providers. For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). For *Outlook*, there is a recommended fall back system. For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="bf261-123">Funcionamento do SSO em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="bf261-123">How it works at runtime</span></span>

<span data-ttu-id="bf261-124">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="bf261-124">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="bf261-p106">No suplemento, o JavaScript chama uma nova API Office.js[getAccessTokenAsync](#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. Consulte [Exemplo de token de acesso](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="bf261-p106">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="bf261-129">Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.</span><span class="sxs-lookup"><span data-stu-id="bf261-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="bf261-130">Se essa é a primeira vez que o usuário atual usa o suplemento, será solicitado que informe seu consentimento.</span><span class="sxs-lookup"><span data-stu-id="bf261-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="bf261-131">O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual.</span><span class="sxs-lookup"><span data-stu-id="bf261-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="bf261-132">O Azure AD envia o token do suplemento para o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="bf261-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="bf261-133">O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="bf261-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="bf261-134">O JavaScript no suplemento pode analisar o token e extrair as informações necessárias, como o endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="bf261-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="bf261-p107">Opcionalmente, o suplemento pode enviar a solicitação HTTP para o seu servidor visando coletar mais dados sobre o usuário, como as preferências do usuário, por exemplo. Ou o próprio token de acesso pode ser enviado para o servidor visando a análise e a validação.</span><span class="sxs-lookup"><span data-stu-id="bf261-p107">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences. Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="bf261-137">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="bf261-137">Develop an SSO add-in</span></span>

<span data-ttu-id="bf261-p108">Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO. Essas tarefas são descritas aqui de forma independente de idioma e estrutura. Para ver exemplos de passo a passo detalhado, confira:</span><span class="sxs-lookup"><span data-stu-id="bf261-p108">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="bf261-141">Criar um suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="bf261-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="bf261-142">Criar um Suplemento do Office com ASP.NET que usa logon único</span><span class="sxs-lookup"><span data-stu-id="bf261-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="bf261-143">Criar o aplicativo de serviço</span><span class="sxs-lookup"><span data-stu-id="bf261-143">Create the service application</span></span>

<span data-ttu-id="bf261-p109">Registre o suplemento no portal de registro para o ponto de extremidade v2.0 do Azure: https://apps.dev.microsoft.com. Esse é um processo que leva de 5 a 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="bf261-p109">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="bf261-146">Obter uma ID do cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="bf261-p110">Especifique as permissões que seu suplemento precisa para o ponto de extremidade AAD v. 2.0 (e, opcionalmente, para o Microsoft Graph). A permissão "perfil" sempre será necessária.</span><span class="sxs-lookup"><span data-stu-id="bf261-p110">Specify the permissions that your add-in needs to AAD v. 2.0 endpoint (and optionally to Microsoft Graph). The "profile" permission is always needed.</span></span>
* <span data-ttu-id="bf261-150">Conceda a relação de confiança do aplicativo host do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="bf261-151">Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="bf261-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="bf261-152">Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="bf261-152">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="bf261-153">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="bf261-153">Configure the add-in</span></span>

<span data-ttu-id="bf261-154">Adicione novas marcações ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="bf261-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="bf261-155">**WebApplicationInfo** – O pai dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="bf261-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="bf261-p111">**Id** - ID do cliente do suplemento. É uma ID de aplicativo que você obtém como parte do processo de registro do suplemento. Confira [Registrar um suplemento do Office que usa o SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="bf261-p111">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="bf261-158">**Recurso** – A URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="bf261-159">**Escopos** – O pai de um ou mais elementos **Escopo**.</span><span class="sxs-lookup"><span data-stu-id="bf261-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="bf261-p112">**Escopo** - Especifica uma permissão que o suplemento precisa para AAD. A permissão `profile` sempre é necessária e pode ser a única permissão necessária se seu suplemento não acessar o Microsoft Graph. Se ele tiver acesso, elementos **Escopo** também são necessários para as permissões do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. As bibliotecas que você usar em seu código para acessar o Microsoft Graph podem precisar de permissões adicionais. Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`. Para mais informações, confira [Autorizar para o Microsoft Graph a partir de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="bf261-p112">**Scope** - Specifies a permission that the add-in needs to AAD. The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="bf261-p113">Para hosts do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="bf261-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="bf261-168">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="bf261-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="bf261-169">Adicionar código do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="bf261-169">Add client-side code</span></span>

<span data-ttu-id="bf261-170">Adicione o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="bf261-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="bf261-171">Chame [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="bf261-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="bf261-172">Analisar o token de acesso ou passá-lo para o código do servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="bf261-173">Aqui está um exemplo simples de uma chamada para `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="bf261-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="bf261-p114">Este exemplo trata apenas de um tipo de erro explicitamente. Para obter exemplos de manipulação de erro mais elaborada, confira [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js no Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). Confira [Solucionar mensagens de erro para logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bf261-p114">This example handles only one kind of error explicitly. For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

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

<span data-ttu-id="bf261-p115">Veja um exemplo simples de como passar o token do suplemento para o servidor. O token é incluído como um cabeçalho `Authorization` ao enviar uma solicitação de volta para o servidor. Este exemplo visualiza o envio de dados JSON. Portanto, ele usa o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver gravando para o servidor.</span><span class="sxs-lookup"><span data-stu-id="bf261-p115">Here's a simple example of passing the add-in token to the server-side. The token is included as an `Authorization` header when sending a request back to the server-side. This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="bf261-180">Quando chamar o método</span><span class="sxs-lookup"><span data-stu-id="bf261-180">When to call the method</span></span>

<span data-ttu-id="bf261-181">Se o seu suplemento não puder ser usado quando nenhum usuário estiver conectado ao Office, você deverá chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.</span><span class="sxs-lookup"><span data-stu-id="bf261-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="bf261-p116">Se o suplemento tiver algumas funcionalidades que não exijam um usuário conectado, você deverá chamar `getAccessTokenAsync` *quando o usuário executar uma ação onde seja necessário que o usuário esteja conectado*. Não há nenhuma degradação de desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e irá reutilizá-lo até que ele expire, sem fazer outra chamada para o ponto de extremidade AAD v. 2.0 sempre que `getAccessTokenAsync` for chamado. Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.</span><span class="sxs-lookup"><span data-stu-id="bf261-p116">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="bf261-186">Adicionar código do servidor</span><span class="sxs-lookup"><span data-stu-id="bf261-186">Add server-side code</span></span>

<span data-ttu-id="bf261-p117">Na maioria dos cenários, não há razão para obter o token de acesso, caso o seu suplemento não o passe para um servidor e use-o. Veja algumas tarefas do servidor que seu suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="bf261-p117">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="bf261-p118">Criar um ou mais métodos da API da Web que usam as informações sobre o usuário extraídas do token; por exemplo, um método que procura as preferências do usuário em sua base de dados hospedada. (Confira **Usar o token de SSO como uma identidade** abaixo). Dependendo do seu idioma e da estrutura, as bibliotecas podem estar disponíveis, simplificando o código que você precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="bf261-p118">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base. (See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="bf261-p119">Obtenha os dados do Microsoft Graph. Seu código do servidor deve fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="bf261-p119">Get Microsoft Graph data. Your server-side code should do the following:</span></span>

    * <span data-ttu-id="bf261-193">Validar o token de acesso (confira **Validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="bf261-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="bf261-p120">Inicie o fluxo "em nome de" com uma chamada para o ponto de extremidade do Azure AD v2.0 que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (seu ID e segredo). Nesse contexto, o token de acesso é chamado token de inicialização.</span><span class="sxs-lookup"><span data-stu-id="bf261-p120">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="bf261-196">Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.</span><span class="sxs-lookup"><span data-stu-id="bf261-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="bf261-197">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="bf261-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="bf261-198">Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usuário, veja [Autorizar o Microsoft Graph no Suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="bf261-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="bf261-199">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="bf261-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="bf261-p121">Depois que a API da Web receber o token de acesso, ela deverá validá-lo para utilizá-lo. O token é um JSON Web Token (JWT) e isso significa que a validação funciona como validação do token nos fluxos padrão de OAuth. Há um número de bibliotecas disponíveis que pode manipular a validação de JWT, mas os fundamentos incluem:</span><span class="sxs-lookup"><span data-stu-id="bf261-p121">Once the Web API receives the access token, it must validate it before using it. The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows. There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="bf261-203">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="bf261-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="bf261-204">Verificando se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="bf261-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="bf261-205">Verificar se o token está direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="bf261-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="bf261-206">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="bf261-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="bf261-p122">Tokens válidos de SSO serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`. A declaração `iss` no token deve começar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="bf261-p122">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`. The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="bf261-209">O parâmetro `aud` do token será configurado para a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf261-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="bf261-210">O parâmetro `scp` do token será definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="bf261-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="bf261-211">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="bf261-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="bf261-p123">Se seu suplemento precisa verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade. As seguintes declarações no token relacionam-se com a identidade.</span><span class="sxs-lookup"><span data-stu-id="bf261-p123">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity. The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="bf261-214">`name` – O nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="bf261-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="bf261-215">`preferred_username` - O endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="bf261-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="bf261-216">`oid` – Uma GUID que representa a ID do usuário no Active Directory do Azure.</span><span class="sxs-lookup"><span data-stu-id="bf261-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="bf261-217">`tid` – Uma GUID que representa a ID da organização do usuário no Active Directory do Azure.</span><span class="sxs-lookup"><span data-stu-id="bf261-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="bf261-218">Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ​​para correlacionar a identidade com o serviço de autorização do seu back-end.</span><span class="sxs-lookup"><span data-stu-id="bf261-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="bf261-p124">Por exemplo, o seu serviço pode formatar esses valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário em seu banco de dados de usuário interno. Nas solicitações subsequentes, o usuário pode ser recuperado usando o mesmo valor, enquanto o acesso a recursos específicos pode ser determinado com base nos seus mecanismos existentes de controle de acesso.</span><span class="sxs-lookup"><span data-stu-id="bf261-p124">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database. Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="bf261-221">Exemplo de token de acesso</span><span class="sxs-lookup"><span data-stu-id="bf261-221">Example access token</span></span>

<span data-ttu-id="bf261-p125">A seguir está uma carga decodificada típica de um token de acesso. Para obter informações sobre as propriedades, confira [Referência de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="bf261-p125">The following is a typical decoded payload of an access token. For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="bf261-224">Como usar o SSO com um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="bf261-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="bf261-p126">Existem algumas diferenças pequenas, mas importantes, entre usar o SSO em um suplemento do Outlook em lugar de usá-lo em um suplemento do Excel, PowerPoint ou Word. Leia [Autenticar um usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar único logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="bf261-p126">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="bf261-227">Referência da API de SSO</span><span class="sxs-lookup"><span data-stu-id="bf261-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="bf261-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="bf261-228">getAccessTokenAsync</span></span>

<span data-ttu-id="bf261-p127">O namespace do Office Auth, `Office.context.auth`, fornece um método, `getAccessTokenAsync` que permite ao host do Office obter um token de acesso para o aplicativo da web do suplemento. Indiretamente, isso também permite que o suplemento acesse dados do Microsoft Graph do usuário conectado sem exigir que o usuário entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="bf261-p127">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application. Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="bf261-p128">O método chama o ponto de extremidade do Active Directory do Azure V 2.0 para obter um token de acesso para o aplicativo da web do seu suplemento. Isso permite que os suplementos identifiquem usuários. O código do servidor pode usar este token para acessar o Microsoft Graph para o aplicativo da web do suplemento usando o [fluxo de OAuth "em nome de"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="bf261-p128">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. This enables add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="bf261-234">No Outlook, essa API não é suportada se o suplemento for carregado em uma caixa de correio do Outlook.com ou do Gmail.</span><span class="sxs-lookup"><span data-stu-id="bf261-234">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

<table><tr><td><span data-ttu-id="bf261-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="bf261-235">Hosts</span></span></td><td><span data-ttu-id="bf261-236">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="bf261-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="bf261-237">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="bf261-237">Requirement sets</span></span></td><td>[<span data-ttu-id="bf261-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="bf261-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="bf261-239">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bf261-239">Parameters</span></span>

<span data-ttu-id="bf261-p129">`options` - Opcional. Aceita um objeto `AuthOptions` (veja abaixo) para definir os comportamentos de logon.</span><span class="sxs-lookup"><span data-stu-id="bf261-p129">`options` - Optional. Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="bf261-p130">`callback` - Opcional. Aceita um método de retorno de chamada que pode analisar o token para a ID do usuário ou usar o token no fluxo de "em nome de" para obter acesso ao Microsoft Graph. Se [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` tiver "êxito", `AsyncResult.value` será o token de acesso formatado do AAD v. 2.0 bruto.</span><span class="sxs-lookup"><span data-stu-id="bf261-p130">`callback` - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph. If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.</span></span>

<span data-ttu-id="bf261-p131">A interface `AuthOptions` oferece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v. 2.0 com o método `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="bf261-p131">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessTokenAsync` method.</span></span>

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



