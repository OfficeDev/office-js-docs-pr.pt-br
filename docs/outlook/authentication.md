---
title: Opções de autenticação em suplementos do Outlook
description: Os suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.
ms.date: 11/05/2019
localization_priority: Priority
ms.openlocfilehash: 2913f770b1f0335aae4634d95b8492b204d1e577
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165759"
---
# <a name="authentication-options-in-outlook-add-ins"></a><span data-ttu-id="984e5-103">Opções de autenticação em suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="984e5-103">Authentication options in Outlook add-ins</span></span>

<span data-ttu-id="984e5-104">O suplemento do Outlook pode acessar informações de qualquer lugar na Internet, seja do servidor que hospeda o suplemento, da sua rede interna ou de outro lugar na nuvem.</span><span class="sxs-lookup"><span data-stu-id="984e5-104">Your Outlook add-in can access information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud.</span></span> <span data-ttu-id="984e5-105">Se essas informações estiverem protegidas, o suplemento precisará de uma forma de autenticar o usuário.</span><span class="sxs-lookup"><span data-stu-id="984e5-105">If that information is protected, your add-in needs a way to authenticate your user.</span></span> <span data-ttu-id="984e5-106">Suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.</span><span class="sxs-lookup"><span data-stu-id="984e5-106">Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.</span></span>

## <a name="single-sign-on-access-token"></a><span data-ttu-id="984e5-107">Token de acesso de logon único</span><span class="sxs-lookup"><span data-stu-id="984e5-107">Single sign-on access token</span></span>

<span data-ttu-id="984e5-108">Os tokens de acesso de logon único oferecem uma maneira simples de o suplemento autenticar e obter tokens de acesso para fazer uma chamada para a [API do Microsoft Graph](/graph/overview).</span><span class="sxs-lookup"><span data-stu-id="984e5-108">Single sign-on access tokens provide a seamless way for your add-in to authenticate and obtain access tokens to call the [Microsoft Graph API](/graph/overview).</span></span> <span data-ttu-id="984e5-109">Esse recurso reduz conflitos porque o usuário não precisa inserir credenciais.</span><span class="sxs-lookup"><span data-stu-id="984e5-109">This capability reduces friction since the user is not required to enter their credentials.</span></span>

> [!NOTE]
> <span data-ttu-id="984e5-110">Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="984e5-110">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="984e5-111">Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="984e5-111">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span>
> <span data-ttu-id="984e5-112">Para usar o SSO, você deve carregar a versão beta da biblioteca de JavaScript do Office de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página de inicialização HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="984e5-112">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="984e5-113">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365.</span><span class="sxs-lookup"><span data-stu-id="984e5-113">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="984e5-114">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="984e5-114">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="984e5-115">Considere usar tokens de acesso SSO se o suplemento:</span><span class="sxs-lookup"><span data-stu-id="984e5-115">Consider using SSO access tokens if your add-in:</span></span>

- <span data-ttu-id="984e5-116">For usado principalmente por usuários do Office 365</span><span class="sxs-lookup"><span data-stu-id="984e5-116">Is used primarily by Office 365 users</span></span>
- <span data-ttu-id="984e5-117">Precisa de acesso para:</span><span class="sxs-lookup"><span data-stu-id="984e5-117">Needs access to:</span></span>
    - <span data-ttu-id="984e5-118">Os serviços Microsoft que são expostos como parte do Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="984e5-118">Microsoft services that are exposed as part of Microsoft Graph</span></span>
    - <span data-ttu-id="984e5-119">Um serviço que não seja da Microsoft que você controle</span><span class="sxs-lookup"><span data-stu-id="984e5-119">A non-Microsoft service that you control</span></span>

<span data-ttu-id="984e5-120">O método de autenticação SSO usa o [Fluxo Em Nome De do OAuth2 fornecido pelo Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="984e5-120">The SSO authentication method uses the [OAuth2 On-Behalf-Of flow provided by Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span> <span data-ttu-id="984e5-121">Ele exige o registro do suplemento no [Portal de Registro do Aplicativo](https://apps.dev.microsoft.com/) e a especificação dos escopos necessários do Microsoft Graph no manifesto.</span><span class="sxs-lookup"><span data-stu-id="984e5-121">It requires that the add-in register in the [Application Registration Portal](https://apps.dev.microsoft.com/) and specify any required Microsoft Graph scopes in its manifest.</span></span>

<span data-ttu-id="984e5-122">Usando este método, o suplemento pode obter um token de acesso com escopo para a API de back-end do servidor.</span><span class="sxs-lookup"><span data-stu-id="984e5-122">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="984e5-123">O suplemento usa isso como um token de portador no cabeçalho `Authorization` para autenticar um retorno de chamada para sua API.</span><span class="sxs-lookup"><span data-stu-id="984e5-123">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="984e5-124">Nesse ponto, o servidor pode:</span><span class="sxs-lookup"><span data-stu-id="984e5-124">At that point your server can:</span></span>

- <span data-ttu-id="984e5-125">concluir o fluxo Em Nome De para obter um token de acesso com escopo para a API do Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="984e5-125">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="984e5-126">Usar as informações de identidade no token para estabelecer a identidade do usuário e autenticar seus serviços de back-end</span><span class="sxs-lookup"><span data-stu-id="984e5-126">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="984e5-127">Para obter uma visão geral mais detalhada, confira a [visão geral completa do método de autenticação SSO](../develop/sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="984e5-127">For a more detailed overview, see the [full overview of the SSO authentication method](../develop/sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="984e5-128">Para obter detalhes sobre como usar o token SSO em um suplemento do Outlook, confira [Autenticar o usuário com um token de logon único em um suplemento do Outlook](authenticate-a-user-with-an-sso-token.md).</span><span class="sxs-lookup"><span data-stu-id="984e5-128">For details on using the SSO token in an Outlook add-in, see [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md).</span></span>

<span data-ttu-id="984e5-129">Confira um exemplo de suplemento que usa o token SSO em [Suplemento de Exemplo AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo).</span><span class="sxs-lookup"><span data-stu-id="984e5-129">For a sample add-in that uses the SSO token, see [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).</span></span>

## <a name="exchange-user-identity-token"></a><span data-ttu-id="984e5-130">Token de identidade do usuário do Exchange</span><span class="sxs-lookup"><span data-stu-id="984e5-130">Exchange user identity token</span></span>

<span data-ttu-id="984e5-131">Os tokens de identidade do usuário do Exchange fornecem uma maneira de o suplemento estabelecer a identidade do usuário.</span><span class="sxs-lookup"><span data-stu-id="984e5-131">Exchange user identity tokens provide a way for your add-in to establish the identity of the user.</span></span> <span data-ttu-id="984e5-132">Ao verificar a identidade do usuário, em seguida, você pode executar uma única autenticação no seu sistema de back-end e aceitar o token de identidade de usuário como uma autorização solicitações futuras.</span><span class="sxs-lookup"><span data-stu-id="984e5-132">By verifying the user's identity, you can then perform a one-time authentication into your back-end system, then accept the user identity token as an authorization for future requests.</span></span> <span data-ttu-id="984e5-133">Use o token de identidade do usuário do Exchange:</span><span class="sxs-lookup"><span data-stu-id="984e5-133">Use the Exchange user identity token:</span></span>

- <span data-ttu-id="984e5-134">Quando o suplemento for usado principalmente por usuários locais do Exchange.</span><span class="sxs-lookup"><span data-stu-id="984e5-134">When the add-in is used primarily by Exchange on-premises users.</span></span>
- <span data-ttu-id="984e5-135">Quando o suplemento precisar acessar um serviço que não seja da Microsoft que você controle.</span><span class="sxs-lookup"><span data-stu-id="984e5-135">When the add-in needs access to a non-Microsoft service that you control.</span></span>
- <span data-ttu-id="984e5-136">Como uma autenticação de fallback (e autorização para o Microsoft Graph) quando o suplemento estiver sendo executado em uma versão do Office que não é compatível com o SSO.</span><span class="sxs-lookup"><span data-stu-id="984e5-136">As a fallback authentication (and authorization to Microsoft Graph) when the add-in is running on a version of Office that doesn't support SSO.</span></span>

<span data-ttu-id="984e5-137">Seu suplemento pode chamar [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getuseridentitytokenasync-callback--usercontext-) para obter tokens de identidade do usuário do Exchange.</span><span class="sxs-lookup"><span data-stu-id="984e5-137">Your add-in can call [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getuseridentitytokenasync-callback--usercontext-) to get Exchange user identity tokens.</span></span> <span data-ttu-id="984e5-138">Para obter detalhes sobre o uso desses tokens, confira [Autenticar um usuário com um token de identidade para o Exchange](authenticate-a-user-with-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="984e5-138">For details on using these tokens, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="access-tokens-obtained-via-oauth2-flows"></a><span data-ttu-id="984e5-139">Tokens de acesso obtidos por meio de fluxos do OAuth2</span><span class="sxs-lookup"><span data-stu-id="984e5-139">Access tokens obtained via OAuth2 flows</span></span>

<span data-ttu-id="984e5-140">Os suplementos também podem acessar serviços de terceiros que oferecem suporte ao OAuth2 para autorização.</span><span class="sxs-lookup"><span data-stu-id="984e5-140">Add-ins can also access third-party services that support OAuth2 for authorization.</span></span> <span data-ttu-id="984e5-141">Considere usar tokens OAuth2 se o suplemento:</span><span class="sxs-lookup"><span data-stu-id="984e5-141">Consider using OAuth2 tokens if your add-in:</span></span>

- <span data-ttu-id="984e5-142">Precisar acessar um serviço de terceiros fora do seu controle</span><span class="sxs-lookup"><span data-stu-id="984e5-142">Needs access to a third-party service outside of your control</span></span>

<span data-ttu-id="984e5-143">Com esse método, o suplemento solicita que o usuário entre no serviço usando o método [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) para inicializar o fluxo do OAuth2 ou usando a [biblioteca office-js-helpers](https://github.com/OfficeDev/office-js-helpers) para o fluxo do OAuth2 Implícito.</span><span class="sxs-lookup"><span data-stu-id="984e5-143">Using this method, your add-in prompts the user to sign-in to the service either by using the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method to initialize the OAuth2 flow, or by using the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers) to the OAuth2 Implicit flow.</span></span>

## <a name="callback-tokens"></a><span data-ttu-id="984e5-144">Tokens de retorno de chamada</span><span class="sxs-lookup"><span data-stu-id="984e5-144">Callback tokens</span></span>

<span data-ttu-id="984e5-145">Os tokens de retorno de chamada fornecem acesso à caixa de correio do usuário a partir do back-end do servidor usando o [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) ou a [API REST do Outlook](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api).</span><span class="sxs-lookup"><span data-stu-id="984e5-145">Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api).</span></span> <span data-ttu-id="984e5-146">Considere usar tokens de retorno de chamada se o suplemento:</span><span class="sxs-lookup"><span data-stu-id="984e5-146">Consider using callback tokens if your add-in:</span></span>

- <span data-ttu-id="984e5-147">Precisar acessar a caixa de correio do usuário a partir do back-end do servidor.</span><span class="sxs-lookup"><span data-stu-id="984e5-147">Needs access to the user's mailbox from your server back-end.</span></span>

<span data-ttu-id="984e5-148">Os suplementos obtêm tokens de retorno de chamada usando um dos métodos [getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="984e5-148">Add-ins obtain callback tokens using one of the [getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) methods.</span></span> <span data-ttu-id="984e5-149">O nível de acesso é controlado pelas permissões especificadas no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="984e5-149">The level of access is controlled by the permissions specified in the add-in manifest.</span></span>