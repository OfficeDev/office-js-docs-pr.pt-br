---
title: Habilitar o logon único para Suplementos do Office
description: Saiba como habilitar o logon único para suplementos do Office usando contas pessoais, corporativas ou de estudante da Microsoft.
ms.date: 07/30/2020
localization_priority: Priority
ms.openlocfilehash: 104a64fa5a761e06711e9c5f850bba0267830809
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237816"
---
# <a name="enable-single-sign-on-for-office-add-ins"></a><span data-ttu-id="72cdd-103">Habilitar o logon único para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="72cdd-103">Enable single sign-on for Office Add-ins</span></span>


<span data-ttu-id="72cdd-104">Os usuários entram no Office (plataformas online, de dispositivos móveis e de área de trabalho) usando contas pessoais da Microsoft, contas corporativas ou do Microsoft 365 Education.</span><span class="sxs-lookup"><span data-stu-id="72cdd-104">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their Microsoft 365 Education or work account.</span></span> <span data-ttu-id="72cdd-105">Você pode tirar proveito disso e usar o logon único (SSO) para autorizar usuário para suplemento, sem exigir que o usuário entre uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="72cdd-105">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Imagem mostrando o processo de logon de um suplemento](../images/sso-for-office-addins.png)

## <a name="requirements-and-best-practices"></a><span data-ttu-id="72cdd-107">Requisitos e as práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="72cdd-107">Requirements and Best Practices</span></span>

<span data-ttu-id="72cdd-108">Se você estiver trabalhando com um suplemento do **Outlook**, certifique-se de habilitar a Autenticação Moderna para a locação do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="72cdd-108">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="72cdd-109">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="72cdd-109">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="72cdd-110">Você *não* deve confiar no SSO como único método do suplemento de autenticação.</span><span class="sxs-lookup"><span data-stu-id="72cdd-110">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="72cdd-111">Devem implementar um sistema de autenticação alternativo que o suplemento possa se enquadrar em determinadas situações de erro.</span><span class="sxs-lookup"><span data-stu-id="72cdd-111">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="72cdd-112">Você pode usar um sistema de autenticação e tabelas de usuário ou utilizar um dos provedores de logon de redes sociais.</span><span class="sxs-lookup"><span data-stu-id="72cdd-112">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="72cdd-113">Para mais informações sobre como fazer isso com um Suplemento do Office, consulte [Autorizar serviços externos no Suplemento do Office](auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-113">For more information about how to do this with an Office Add-in, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).</span></span> <span data-ttu-id="72cdd-114">Para *Outlook*, há um sistema de fallback recomendado.</span><span class="sxs-lookup"><span data-stu-id="72cdd-114">For *Outlook*, there is a recommended fallback system.</span></span> <span data-ttu-id="72cdd-115">Para mais informações, confira [Cenário: implementar o logon único no serviço em um Suplemento do Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-115">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span> <span data-ttu-id="72cdd-116">Para exemplos que usam o Azure Active Directory como o sistema de fallback, confira [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) e [SSO com Suplemento ASP.NET do Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="72cdd-116">For samples that use Azure Active Directory as the fallback system, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>



## <a name="how-sso-works-at-runtime"></a><span data-ttu-id="72cdd-117">Como o SSO funciona em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="72cdd-117">How SSO works at runtime</span></span>

<span data-ttu-id="72cdd-118">O diagrama a seguir mostra como funciona o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="72cdd-118">The following diagram shows how the SSO process works.</span></span>

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="72cdd-120">No suplemento, o JavaScript chama uma nova API do Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="72cdd-120">In the add-in, JavaScript calls a new Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span> <span data-ttu-id="72cdd-121">Isso informa ao aplicativo cliente do Office para obter um token de acesso para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-121">This tells the Office client application to obtain an access token to the add-in.</span></span> <span data-ttu-id="72cdd-122">Confira [Token de acesso de amostra](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="72cdd-122">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="72cdd-123">Se o usuário não estiver conectado, o aplicativo cliente do Office abrirá uma janela pop-up para o usuário entrar.</span><span class="sxs-lookup"><span data-stu-id="72cdd-123">If the user is not signed in, the Office client application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="72cdd-124">Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-124">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="72cdd-125">O aplicativo cliente do Office solicita o **token do suplemento** do ponto de extremidade v2.0 do Azure AD para o usuário atual. </span><span class="sxs-lookup"><span data-stu-id="72cdd-125">The Office client application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="72cdd-126">O Azure AD envia o token do suplemento ao aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="72cdd-126">Azure AD sends the add-in token to the Office client application.</span></span>
6. <span data-ttu-id="72cdd-127">O aplicativo cliente do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-127">The Office client application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessToken` call.</span></span>
7. <span data-ttu-id="72cdd-128">O JavaScript no suplemento pode analisar o token e extrair informações necessárias, como endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="72cdd-128">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span>
8. <span data-ttu-id="72cdd-129">Opcionalmente, o suplemento pode enviar solicitação HTTP para o servidor para obter mais dados sobre o usuário; como as preferências do usuário.</span><span class="sxs-lookup"><span data-stu-id="72cdd-129">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="72cdd-130">Como alternativa, o próprio token de acesso pode ser enviado para o servidor para análise e validação.</span><span class="sxs-lookup"><span data-stu-id="72cdd-130">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="72cdd-131">Desenvolver um suplemento com SSO</span><span class="sxs-lookup"><span data-stu-id="72cdd-131">Develop an SSO add-in</span></span>

<span data-ttu-id="72cdd-132">Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO.</span><span class="sxs-lookup"><span data-stu-id="72cdd-132">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="72cdd-133">Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agnóstica.</span><span class="sxs-lookup"><span data-stu-id="72cdd-133">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="72cdd-134">Para orientações detalhadas, confira:</span><span class="sxs-lookup"><span data-stu-id="72cdd-134">For detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="72cdd-135">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="72cdd-135">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="72cdd-136">Criar um Suplemento do Office com ASP.NET que usa logon único</span><span class="sxs-lookup"><span data-stu-id="72cdd-136">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> <span data-ttu-id="72cdd-137">Você pode usar o gerador Yeoman para criar um Suplemento do Office com Node.js habilitado para SSO.</span><span class="sxs-lookup"><span data-stu-id="72cdd-137">You can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="72cdd-138">O gerador Yeoman simplifica o processo de criação de um suplemento habilitado para SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO.</span><span class="sxs-lookup"><span data-stu-id="72cdd-138">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="72cdd-139">Para obter mais informações, confira [Início rápido de logon único (SSO)](../quickstarts/sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-139">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

### <a name="create-the-service-application"></a><span data-ttu-id="72cdd-140">Criar o aplicativo de serviço</span><span class="sxs-lookup"><span data-stu-id="72cdd-140">Create the service application</span></span>

<span data-ttu-id="72cdd-p108">Registre o suplemento no portal de registro para o ponto de extremidade do Azure v 2.0. Esse é um processo que leva entre 5 e 10 minutos e inclui as seguintes tarefas:</span><span class="sxs-lookup"><span data-stu-id="72cdd-p108">Register the add-in at the registration portal for the Azure v2.0 endpoint. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="72cdd-143">Obter um ID de cliente e o segredo para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-143">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="72cdd-144">Especificar as permissões que seu suplemento precisa de AAD v.</span><span class="sxs-lookup"><span data-stu-id="72cdd-144">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="72cdd-145">ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="72cdd-145">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="72cdd-146">As permissões "perfil" e "openid" são sempre necessárias.</span><span class="sxs-lookup"><span data-stu-id="72cdd-146">The "profile" and "openid" permissions are always needed.</span></span>
* <span data-ttu-id="72cdd-147">Conceder a confiança do aplicativo cliente do Office para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-147">Grant the Office client application trust to the add-in.</span></span>
* <span data-ttu-id="72cdd-148">Autorizar previamente o aplicativo cliente do Office para o suplemento com a permissão padrão *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="72cdd-148">Preauthorize the Office client application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="72cdd-149">Para mais detalhes sobre esse processo, confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-149">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="72cdd-150">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="72cdd-150">Configure the add-in</span></span>

<span data-ttu-id="72cdd-151">Adicione novas marcações ao manifesto do suplemento:</span><span class="sxs-lookup"><span data-stu-id="72cdd-151">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="72cdd-152">**WebApplicationInfo** – o pai dos seguintes elementos.</span><span class="sxs-lookup"><span data-stu-id="72cdd-152">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="72cdd-153">**ID** - O ID do cliente do suplemento Este é um ID do aplicativo que você obtém como parte do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-153">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="72cdd-154">Confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-154">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="72cdd-155">**Resource** – A URL do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-155">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="72cdd-156">Esse é o mesmo URI (incluindo o protocolo `api:`) que você usou ao registrar o suplemento no AAD.</span><span class="sxs-lookup"><span data-stu-id="72cdd-156">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="72cdd-157">Parte do domínio deste URI deve corresponder ao domínio, incluindo quaisquer subdomínios, usados nos URLs na seção `<Resources>` do manifesto do suplemento e o URI deve terminar com o ID do cliente no `<Id>`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-157">The domain part of this URI must match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest and the URI must end with the client ID in the `<Id>`.</span></span>
* <span data-ttu-id="72cdd-158">**Scopes** – O pai de uma ou mais elementos **Scope**.</span><span class="sxs-lookup"><span data-stu-id="72cdd-158">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="72cdd-159">**Scope** – Especifica uma permissão que seu suplemento precisa para o AAD.</span><span class="sxs-lookup"><span data-stu-id="72cdd-159">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="72cdd-160">As permissões `profile` e `openID` são sempre necessárias e podem ser as únicas permissões necessárias, se o suplemento não acessar o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="72cdd-160">The `profile` and `openID` permissions are always needed and may be the only permissions needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="72cdd-161">Se isso acontecer, você também precisa de elementos **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-161">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="72cdd-162">Bibliotecas que você usa no seu código para acessar o Microsoft Graph pode precisar de permissões adicionais.</span><span class="sxs-lookup"><span data-stu-id="72cdd-162">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="72cdd-163">Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-163">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="72cdd-164">Para saber mais, confira [autorizar o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-164">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="72cdd-p113">Para aplicativos do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-p113">For Office applications other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="72cdd-167">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="72cdd-167">The following is an example of the markup:</span></span>

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>openid</Scope>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```
> [!NOTE]
> <span data-ttu-id="72cdd-168">O não cumprimento dos requisitos de formato no manifesto para SSO fará com que seu suplemento seja rejeitado do AppSource até que atenda ao formato exigido.</span><span class="sxs-lookup"><span data-stu-id="72cdd-168">Not following the format requirements in the manifest for SSO will cause your add-in to be rejected from AppSource until it meets the required format.</span></span>

### <a name="add-client-side-code"></a><span data-ttu-id="72cdd-169">Adicionar código do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="72cdd-169">Add client-side code</span></span>

<span data-ttu-id="72cdd-170">Adicione o JavaScript ao suplemento para:</span><span class="sxs-lookup"><span data-stu-id="72cdd-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="72cdd-171">Chamar [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="72cdd-171">Call [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span>

* <span data-ttu-id="72cdd-172">Analisar o token de acesso ou encaminhá-lo ao código de servidor do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-172">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="72cdd-173">Aqui está um exemplo simples de uma chamada para `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-173">Here's a simple example of a call to `getAccessToken`.</span></span>

> [!NOTE]
> <span data-ttu-id="72cdd-174">Este exemplo lida explicitamente com apenas um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="72cdd-174">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="72cdd-175">Para exemplos de tratamento de erro mais elaborados, confira [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) e [SSO com Suplemento ASP.NET do Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="72cdd-175">For examples of more elaborate error handling, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>


```js
async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken();

        // The /api/DoSomething controller will make the token exchange and use the
        // access token it gets back to make the call to MS Graph.
        getData("/api/DoSomething", bootstrapToken);
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // Microsoft 365 Education or work account, or a Microsoft account.
        } else {
            // Handle error
        }
    }
}
```

<span data-ttu-id="72cdd-176">Aqui está um exemplo simples de como passar o token de suplemento para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="72cdd-176">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="72cdd-177">O token é incluído como um cabeçalho de `Authorization`ao enviar uma solicitação para o lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="72cdd-177">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="72cdd-178">Este exemplo prevê enviar dados JSON, para que ele tenha o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver escrevendo no servidor.</span><span class="sxs-lookup"><span data-stu-id="72cdd-178">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="72cdd-179">Quando chamar o método</span><span class="sxs-lookup"><span data-stu-id="72cdd-179">When to call the method</span></span>

<span data-ttu-id="72cdd-180">Se o seu suplemento não puder ser usado quando não houver usuário conectado ao Office, você deve chamar `getAccessToken` *quando o suplemento for iniciado* e passar `allowSignInPrompt: true` no `options` parâmetro `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-180">If your add-in cannot be used when there is no user currently logged into Office, then you should call `getAccessToken` *when the add-in launches* and pass `allowSignInPrompt: true` in the `options` parameter of `getAccessToken`.</span></span> <span data-ttu-id="72cdd-181">Por exemplo: `OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });`</span><span class="sxs-lookup"><span data-stu-id="72cdd-181">For example; `OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });`</span></span>

<span data-ttu-id="72cdd-182">Se o complemento tiver alguma funcionalidade que não exija um usuário conectado, então chame `getAccessToken` *quando o usuário fizer uma ação que exija acesso a um usuário logado*.</span><span class="sxs-lookup"><span data-stu-id="72cdd-182">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessToken` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="72cdd-183">Não há uma degradação significativa do desempenho com chamadas redundantes de `getAccessToken` porque o Office armazena em cache o token de inicialização e o reutilizará, até que ele expire, sem fazer outra chamada para o AAD v.</span><span class="sxs-lookup"><span data-stu-id="72cdd-183">There is no significant performance degradation with redundant calls of `getAccessToken` because Office caches the bootstrap token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="72cdd-184">Ponto de extremidade 2.0 sempre que `getAccessToken` for chamado.</span><span class="sxs-lookup"><span data-stu-id="72cdd-184">2.0 endpoint whenever `getAccessToken` is called.</span></span> <span data-ttu-id="72cdd-185">Portanto, você pode adicionar chamadas de `getAccessToken` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.</span><span class="sxs-lookup"><span data-stu-id="72cdd-185">So you can add calls of `getAccessToken` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="72cdd-186">Adicionar código no lado do servidor</span><span class="sxs-lookup"><span data-stu-id="72cdd-186">Add server-side code</span></span>

<span data-ttu-id="72cdd-187">Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá.</span><span class="sxs-lookup"><span data-stu-id="72cdd-187">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="72cdd-188">Algumas tarefas de servidor que o suplemento pode fazer:</span><span class="sxs-lookup"><span data-stu-id="72cdd-188">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="72cdd-189">Criar um ou mais métodos de Web API com informações sobre o usuário que são extraídas do token; Por exemplo, uma forma que procura preferências do usuário em seu banco de dados hospedado.</span><span class="sxs-lookup"><span data-stu-id="72cdd-189">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="72cdd-190">(Confira **usando o token SSO, como uma identidade** abaixo.)Dependendo do seu idioma e da estrutura, podem estar disponíveis bibliotecas que simplificarão o código que você precisa escrever.</span><span class="sxs-lookup"><span data-stu-id="72cdd-190">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="72cdd-191">Obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="72cdd-191">Get Microsoft Graph data.</span></span> <span data-ttu-id="72cdd-192">O código do lado do servidor precisa fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="72cdd-192">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="72cdd-193">Iniciar o fluxo "on behalf of" com uma chamada para o ponto de extremidade v 2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo).</span><span class="sxs-lookup"><span data-stu-id="72cdd-193">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="72cdd-194">O token de acesso nesse contexto é chamado de bootstrap token.</span><span class="sxs-lookup"><span data-stu-id="72cdd-194">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="72cdd-195">Obter os dados do Microsoft Graph usando o novo token.</span><span class="sxs-lookup"><span data-stu-id="72cdd-195">Get data from Microsoft Graph by using the new token.</span></span>
    * <span data-ttu-id="72cdd-196">Opcionalmente, valide o token de acesso antes de iniciar o fluxo (confira **Validar o token de acesso** abaixo).</span><span class="sxs-lookup"><span data-stu-id="72cdd-196">Optionally, before initiating the flow, validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="72cdd-197">Opcionalmente, após a conclusão do fluxo on-behalf-of, armazene em cache o novo token de acesso retornado do fluxo, de forma que ele seja reutilizado em outras chamadas para o Microsoft Graph até que ele expire.</span><span class="sxs-lookup"><span data-stu-id="72cdd-197">Optionally, after the on-behalf-of flow completes, cache the new access token that is returned from the flow so that it an be reused in other calls to Microsoft Graph until it expires.</span></span>

 <span data-ttu-id="72cdd-198">Para saber mais sobre como obter acesso autorizado aos dados do usuário Microsoft Graph, confira [Autorizar o Microsoft Graph nos suplementos do Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="72cdd-199">Validar o token de acesso</span><span class="sxs-lookup"><span data-stu-id="72cdd-199">Validate the access token</span></span>

<span data-ttu-id="72cdd-200">Após a API da Web receber o token de acesso, ela deve validá-lo antes que ele possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="72cdd-200">Once the Web API receives the access token, it can validate it before using it.</span></span> <span data-ttu-id="72cdd-201">O token é um Token Web JSON (JWT) e isso significa que validação funciona como uma validação de token na maioria dos fluxos padrão do OAuth.</span><span class="sxs-lookup"><span data-stu-id="72cdd-201">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="72cdd-202">Há diversas bibliotecas disponíveis que podem lidar com a validação de JWT. No entanto, as noções básicas incluem:</span><span class="sxs-lookup"><span data-stu-id="72cdd-202">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="72cdd-203">Verificar se o token foi bem formado</span><span class="sxs-lookup"><span data-stu-id="72cdd-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="72cdd-204">Verificar se o token foi emitido pela autoridade desejada</span><span class="sxs-lookup"><span data-stu-id="72cdd-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="72cdd-205">Verificar se o token está direcionado para a API Web</span><span class="sxs-lookup"><span data-stu-id="72cdd-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="72cdd-206">Ao validar o token, lembre-se das seguintes diretrizes:</span><span class="sxs-lookup"><span data-stu-id="72cdd-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="72cdd-207">Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-207">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="72cdd-208">A declaração `iss` no token deve começar com esse valor.</span><span class="sxs-lookup"><span data-stu-id="72cdd-208">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="72cdd-209">O parâmetro `aud` do token será configurado como a ID de aplicativo do registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="72cdd-210">O parâmetro `scp` do token será definido como `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="72cdd-211">Usar o token SSO como uma identidade</span><span class="sxs-lookup"><span data-stu-id="72cdd-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="72cdd-212">Se o suplemento precisar verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade.</span><span class="sxs-lookup"><span data-stu-id="72cdd-212">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="72cdd-213">As seguintes declarações no token estão relacionadas à identidade.</span><span class="sxs-lookup"><span data-stu-id="72cdd-213">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="72cdd-214">`name` – O nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="72cdd-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="72cdd-215">`preferred_username`O endereço de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="72cdd-215">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="72cdd-216">`oid` – Um GUID que representa a ID do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="72cdd-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="72cdd-217">`tid` – Um GUID que representa a ID da organização do usuário no Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="72cdd-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="72cdd-218">Como os valores `name` e `preferred_username` podem alterar, recomendamos que os valores `oid` e `tid` sejam usados para correlacionar a identidade com o serviço de autorização do back-end.</span><span class="sxs-lookup"><span data-stu-id="72cdd-218">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="72cdd-219">Por exemplo, o serviço poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados do usuário interno.</span><span class="sxs-lookup"><span data-stu-id="72cdd-219">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="72cdd-220">Em seguida, nas solicitações subsequentes, o usuário poderia ser recuperado usando o mesmo valor e o acesso a recursos específicos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.</span><span class="sxs-lookup"><span data-stu-id="72cdd-220">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="72cdd-221">Token de acesso de exemplo</span><span class="sxs-lookup"><span data-stu-id="72cdd-221">Example access token</span></span>

<span data-ttu-id="72cdd-222">A seguir está uma carga decodificada típica do token de acesso.</span><span class="sxs-lookup"><span data-stu-id="72cdd-222">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="72cdd-223">Para saber mais sobre as propriedades, confira [Referência de tokens de versão do Azure Active Directory 2.0](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="72cdd-223">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>

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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="72cdd-224">Usando o SSO com um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="72cdd-224">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="72cdd-225">Há algumas diferenças pequenas, mas importantes entre usar o SSO em um suplemento do Outlook e em um suplemento do Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="72cdd-225">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="72cdd-226">Não deixe de ler [Autenticar o usuário com um token de logon único em um suplemento do Outlook](../outlook/authenticate-a-user-with-an-sso-token.md) e [Cenário: implementar o logon único ao serviço em um suplemento do Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="72cdd-226">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](../outlook/authenticate-a-user-with-an-sso-token.md) and [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="72cdd-227">Referência da API do SSO</span><span class="sxs-lookup"><span data-stu-id="72cdd-227">SSO API reference</span></span>

### <a name="getaccesstoken"></a><span data-ttu-id="72cdd-228">getAccessToken</span><span class="sxs-lookup"><span data-stu-id="72cdd-228">getAccessToken</span></span>

<span data-ttu-id="72cdd-229">O namespace [Auth](/javascript/api/office-runtime/officeruntime.auth) do OfficeRuntime, `OfficeRuntime.Auth`, fornece um método, `getAccessToken` que permite com que o aplicativo do Office obtenha um token de acesso para o aplicativo web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-229">The OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) namespace, `OfficeRuntime.Auth`, provides a method, `getAccessToken` that enables the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="72cdd-230">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="72cdd-230">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessToken(options?: AuthOptions: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="72cdd-231">O método chama o ponto de extremidade do Azure Active Directory V 2.0 para obter um token de acesso para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="72cdd-231">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="72cdd-232">Isso permite que os suplementos identifiquem usuários.</span><span class="sxs-lookup"><span data-stu-id="72cdd-232">This enables add-ins to identify users.</span></span> <span data-ttu-id="72cdd-233">O código do lado do servidor pode usar esse token para acessar o Microsoft Graph do aplicativo Web do suplemento usando o [fluxo OAuth "em nome de"](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="72cdd-233">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="72cdd-234">No Outlook, não há suporte para esse API se o suplemento for carregado em uma caixa de correio do Gmail ou do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="72cdd-234">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="72cdd-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="72cdd-235">Hosts</span></span>|<span data-ttu-id="72cdd-236">Excel, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="72cdd-236">Excel, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="72cdd-237">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="72cdd-237">Requirement sets</span></span>](specify-office-hosts-and-api-requirements.md)|[<span data-ttu-id="72cdd-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="72cdd-238">IdentityAPI</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)|

#### <a name="parameters"></a><span data-ttu-id="72cdd-239">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="72cdd-239">Parameters</span></span>

<span data-ttu-id="72cdd-240">`options` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="72cdd-240">`options` - Optional.</span></span> <span data-ttu-id="72cdd-241">Aceite um objeto [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) (veja abaixo) para definir comportamentos de logon.</span><span class="sxs-lookup"><span data-stu-id="72cdd-241">Accepts an [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="72cdd-242">`callback` – Opcional.</span><span class="sxs-lookup"><span data-stu-id="72cdd-242">`callback` - Optional.</span></span> <span data-ttu-id="72cdd-243">Aceita um método de retorno que possa analisar o token de ID de usuário ou usar o token fluxo "em nome de" para obter acesso ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="72cdd-243">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="72cdd-244">Se [AsyncResult](/javascript/api/office/office.asyncresult) `.status` é "bem-sucedido", em seguida, `AsyncResult.value` é o v AAD bruto.</span><span class="sxs-lookup"><span data-stu-id="72cdd-244">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="72cdd-245">token de acesso 2.0 formatado.</span><span class="sxs-lookup"><span data-stu-id="72cdd-245">2.0-formatted access token.</span></span>

<span data-ttu-id="72cdd-246">A interface [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) fornece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v.</span><span class="sxs-lookup"><span data-stu-id="72cdd-246">The [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="72cdd-247">2.0 com o método`getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="72cdd-247">2.0 with the `getAccessToken` method.</span></span>
