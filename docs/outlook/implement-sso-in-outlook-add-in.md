---
title: 'Cenário: implementar o logon único no seu serviço'
description: Saiba como usar o token de logon único e o token de identidade do Exchange fornecidos por um suplemento do Outlook para implementar o SSO com o serviço.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 933387e941d4fb3f4d749b319abb01f1c931c4e2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165734"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a><span data-ttu-id="4e425-103">Cenário: implementar o logon único no serviço em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="4e425-103">Scenario: Implement single sign-on to your service in an Outlook add-in</span></span>

<span data-ttu-id="4e425-104">Neste artigo exploraremos um método recomendado de usar o [token de acesso de logon único](authenticate-a-user-with-an-sso-token.md) e o [token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) juntos para fornecer um logon único na implementação do seu próprio serviço de back-end.</span><span class="sxs-lookup"><span data-stu-id="4e425-104">In this article we'll explore a recommended method of using the [single sign-on access token](authenticate-a-user-with-an-sso-token.md) and the [Exchange identity token](authenticate-a-user-with-an-identity-token.md) together to provide a single-sign on implementation to your own backend service.</span></span> <span data-ttu-id="4e425-105">Usando dois tokens em conjunto, será possível aproveitar os benefícios do token de acesso SSO quando ele estiver disponível, garantindo que o suplemento funcionará quando ele não estiver disponível, como quando o usuário alterna para um cliente não compatível ou quando a caixa de correio do usuário está em um servidor do Exchange local.</span><span class="sxs-lookup"><span data-stu-id="4e425-105">By using both tokens together, you can take advantage of the benefits of the SSO access token when it is available, while ensuring that your add-in will work when it is not, such as when the user switches to a client that does not support them, or if the user's mailbox is on an on-premises Exchange server.</span></span>


> [!NOTE]
> <span data-ttu-id="4e425-106">Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4e425-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="4e425-107">Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="4e425-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span>
> <span data-ttu-id="4e425-108">Para usar o SSO, você deve carregar a versão beta da biblioteca de JavaScript do Office de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página de inicialização HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4e425-108">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="4e425-109">Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365.</span><span class="sxs-lookup"><span data-stu-id="4e425-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="4e425-110">Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="4e425-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>


## <a name="why-use-the-sso-access-token"></a><span data-ttu-id="4e425-111">Por que usar o token de acesso SSO?</span><span class="sxs-lookup"><span data-stu-id="4e425-111">Why use the SSO access token?</span></span>

<span data-ttu-id="4e425-112">O token de identidade do Exchange está disponível em todos os conjuntos de requisitos de APIs do suplemento, portanto, pode parecer tentador depender simplesmente desse token e ignorar o token SSO completamente.</span><span class="sxs-lookup"><span data-stu-id="4e425-112">The Exchange identity token is available in all requirement sets of the add-in APIs, so it may be tempting to just rely on this token and ignore the SSO token altogether.</span></span> <span data-ttu-id="4e425-113">No entanto, o token SSO oferece algumas vantagens em relação ao token de identidade do Exchange, portanto, quando disponível, torna-se o método recomendado.</span><span class="sxs-lookup"><span data-stu-id="4e425-113">However, the SSO token offers some advantages over the Exchange identity token which make it the recommended method to use when it is available.</span></span>

- <span data-ttu-id="4e425-114">O token SSO usa um formato padrão OpenID e é emitido pelo Azure.</span><span class="sxs-lookup"><span data-stu-id="4e425-114">The SSO token uses a standard OpenID format and is issued by Azure.</span></span> <span data-ttu-id="4e425-115">Isso simplifica bastante o processo de validação desses tokens.</span><span class="sxs-lookup"><span data-stu-id="4e425-115">This greatly simplifies the process of validating these tokens.</span></span> <span data-ttu-id="4e425-116">Em comparação, os tokens de identidade do Exchange usam um formato personalizado com base no Token Web JSON padrão, exigindo trabalho personalizado para validar o token.</span><span class="sxs-lookup"><span data-stu-id="4e425-116">In comparison, Exchange identity tokens use a custom format based on the JSON Web Token standard, requiring custom work to validate the token.</span></span>
- <span data-ttu-id="4e425-117">O token SSO pode ser usado pelo back-end para recuperar um token de acesso do Microsoft Graph sem que o usuário tenha que fazer qualquer outra ação de entrada.</span><span class="sxs-lookup"><span data-stu-id="4e425-117">The SSO token can be used by your backend to retrieve an access token for Microsoft Graph without the user having to do any additional sign in action.</span></span>
- <span data-ttu-id="4e425-118">O token SSO fornece informações avançadas de identidade, como o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="4e425-118">The SSO token provides richer identity information, such as the user's display name.</span></span>

## <a name="add-in-scenario"></a><span data-ttu-id="4e425-119">Cenário de suplemento</span><span class="sxs-lookup"><span data-stu-id="4e425-119">Add-in scenario</span></span>

<span data-ttu-id="4e425-120">Para este exemplo, considere um suplemento formado pela interface do usuário e scripts (HTML + JavaScript) do suplemento e uma API Web de back-end chamada pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="4e425-120">For the purposes of this example, consider an add-in that consists of both the add-in UI and scripts (HTML + JavaScript) and a backend Web API that is called by the add-in.</span></span> <span data-ttu-id="4e425-121">A API Web de back-end faz chamadas para a [API do Microsoft Graph](/graph/overview) e a API de Dados da Contoso, uma API fictícia criada por terceiros.</span><span class="sxs-lookup"><span data-stu-id="4e425-121">The backend Web API makes calls both to the [Microsoft Graph API](/graph/overview) and the Contoso Data API, a fictional API created by a third party.</span></span> <span data-ttu-id="4e425-122">Como a API do Microsoft Graph, a API de Dados da Contoso requer autenticação OAuth.</span><span class="sxs-lookup"><span data-stu-id="4e425-122">Like the Microsoft Graph API, the Contoso Data API requires OAuth authentication.</span></span> <span data-ttu-id="4e425-123">O requisito é que a API Web de back-end seja capaz de chamar as duas APIs sem ter que solicitar ao usuário que forneça credenciais sempre que um token de acesso expirar.</span><span class="sxs-lookup"><span data-stu-id="4e425-123">The requirement is that the backend Web API should be able to call both APIs without having to prompt the user for their credentials every time an access token expires.</span></span>

<span data-ttu-id="4e425-124">Para fazer isso, a API de backend cria um banco de dados de usuários seguro.</span><span class="sxs-lookup"><span data-stu-id="4e425-124">To do this, the backend API creates a secure database of users.</span></span> <span data-ttu-id="4e425-125">Cada usuário receberá uma entrada no banco de dados onde o back-end armazenará tokens de atualização de vida longa da API do Microsoft Graph API e da API de Dados da Contoso.</span><span class="sxs-lookup"><span data-stu-id="4e425-125">Each user will get an entry in the database where the backend can store long-lived refresh tokens for both the Microsoft Graph API and the Contoso Data API.</span></span> <span data-ttu-id="4e425-126">A marcação JSON a seguir representa uma entrada do usuário no banco de dados.</span><span class="sxs-lookup"><span data-stu-id="4e425-126">The following JSON markup represents a user's entry in the database.</span></span>

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

<span data-ttu-id="4e425-127">O suplemento inclui o token de acesso SSO (se estiver disponível) ou o token de identidade do Exchange (se o token SSO não estiver disponível) com todas as chamadas feitas para a API Web de back-end.</span><span class="sxs-lookup"><span data-stu-id="4e425-127">The add-in includes either the SSO access token (if it is available) or the Exchange identity token (if the SSO token is not available) with every call it makes to the backend Web API.</span></span>

### <a name="add-in-startup"></a><span data-ttu-id="4e425-128">Inicialização do suplemento</span><span class="sxs-lookup"><span data-stu-id="4e425-128">Add-in startup</span></span>

1. <span data-ttu-id="4e425-129">Quando o suplemento iniciar, ele enviará uma solicitação à API Web de back-end para determinar se o usuário está registrado (por exemplo, se tem um registro associado no banco de dados do usuário) e se a API tem tokens de atualização para o Graph e para a Contoso.</span><span class="sxs-lookup"><span data-stu-id="4e425-129">When the add-in starts, it sends a request to the backend Web API to determine if the user is registered (i.e. has an associated record in the user database) and that the API has refresh tokens for both Graph and Contoso.</span></span> <span data-ttu-id="4e425-130">Nessa chamada, o suplemento inclui o token SSO (se disponível) e o token de identidade.</span><span class="sxs-lookup"><span data-stu-id="4e425-130">In this call, the add-in includes both the SSO token (if available) and the identity token.</span></span>

1. <span data-ttu-id="4e425-131">A API Web utiliza os métodos em [Autenticar um usuário com um token de logon único em um suplemento do Outlook](authenticate-a-user-with-an-sso-token.md) e [Autenticar um usuário com um token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) para validar e gerar um identificador exclusivo a partir dos dois tokens.</span><span class="sxs-lookup"><span data-stu-id="4e425-131">The Web API uses the methods in [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md) and [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md) to validate and generate a unique identifier from both tokens.</span></span>

1. <span data-ttu-id="4e425-132">Se um token SSO tiver sido fornecido, a API Web consultará o banco de dados do usuário em busca de uma entrada que tenha um valor `ssoId` que corresponda ao identificador exclusivo gerado pelo token SSO.</span><span class="sxs-lookup"><span data-stu-id="4e425-132">If an SSO token was provided, the Web API then queries the user database for an entry that has an `ssoId` value that matches the unique identifier generated from the SSO token.</span></span>
   - <span data-ttu-id="4e425-133">Se não houver uma entrada, vá para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="4e425-133">If an entry did not exist, continue to the next step.</span></span>
   - <span data-ttu-id="4e425-134">Se houver uma entrada, vá para a etapa 5.</span><span class="sxs-lookup"><span data-stu-id="4e425-134">If an entry exists, proceed to step 5.</span></span>

1. <span data-ttu-id="4e425-135">A API Web consultará o banco de dados em busca de uma entrada que tenha um valor `exchangeId` que corresponda ao identificador exclusivo gerado pelo token de identidade do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e425-135">The Web API queries the database for an entry that has an `exchangeId` value that matches the unique identifier generated from the Exchange identity token.</span></span>
   - <span data-ttu-id="4e425-136">Se houver uma entrada e um token SSO tiver sido fornecido, atualize o registro do usuário no banco de dados para definir o valor `ssoId` para o identificador exclusivo gerado a partir do token SSO e prossiga para a etapa 5.</span><span class="sxs-lookup"><span data-stu-id="4e425-136">If an entry exists and an SSO token was provided, update the user's record in the database to set the `ssoId` value to the unique identifier generated from the SSO token and proceed to step 5.</span></span>
   - <span data-ttu-id="4e425-137">Se houver uma entrada e nenhum token SSO tiver sido fornecido, prossiga para a etapa 5.</span><span class="sxs-lookup"><span data-stu-id="4e425-137">If an entry exists and no SSO token was provided, proceed to step 5.</span></span>
   - <span data-ttu-id="4e425-138">Se não houver entradas, crie uma nova entrada.</span><span class="sxs-lookup"><span data-stu-id="4e425-138">If no entry exists, create a new entry.</span></span> <span data-ttu-id="4e425-139">Defina `ssoId` como o identificador exclusivo gerado por meio do token SSO (se disponível) e defina `exchangeId` como o identificador exclusivo gerado por meio do token de identidade do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e425-139">Set `ssoId` to the unique identifier generated from the SSO token (if available), and set `exchangeId` to the unique identifier generated from the Exchange identity token.</span></span>

1. <span data-ttu-id="4e425-140">Verifique se há um token de atualização válido no valor `graphRefreshToken` do usuário.</span><span class="sxs-lookup"><span data-stu-id="4e425-140">Check for a valid refresh token in the user's `graphRefreshToken` value.</span></span>
   - <span data-ttu-id="4e425-141">Se o valor for inválido ou estiver ausente no token SSO fornecido, use o [Fluxo Em Nome De do OAuth2](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) para obter um token de acesso e atualizar o token do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4e425-141">If the value is missing or invalid and an SSO token was provided, use the [OAuth2 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) to obtain an access token and refresh token for Graph.</span></span> <span data-ttu-id="4e425-142">Salve o token de atualização válido no valor `graphRefreshToken` para o usuário.</span><span class="sxs-lookup"><span data-stu-id="4e425-142">Save the refresh token in the `graphRefreshToken` value for the user.</span></span>

1. <span data-ttu-id="4e425-143">Procure tokens de atualização válidos em `graphRefreshToken` e `contosoRefreshToken`.</span><span class="sxs-lookup"><span data-stu-id="4e425-143">Check for valid refresh tokens in both `graphRefreshToken` and `contosoRefreshToken`.</span></span>
   - <span data-ttu-id="4e425-144">Se ambos valores forem válidos, responda o suplemento para indicar que o usuário já está registrado e configurado.</span><span class="sxs-lookup"><span data-stu-id="4e425-144">If both values are valid, respond to the add-in to indicate that the user is already registered and configured.</span></span>
   - <span data-ttu-id="4e425-145">Se o valor for inválido, responda o suplemento para indicar que a configuração do usuário é obrigatória, além de quais serviços (Contoso ou Graph) precisam ser configurados.</span><span class="sxs-lookup"><span data-stu-id="4e425-145">If either value is invalid, respond to the add-in to indicate that user setup is required, along with which services (Graph or Contoso) need to be configured.</span></span>

1. <span data-ttu-id="4e425-146">O suplemento verifica a resposta.</span><span class="sxs-lookup"><span data-stu-id="4e425-146">The add-in checks the response.</span></span>
   - <span data-ttu-id="4e425-147">Se o usuário já estiver registrado e configurado, o suplemento prosseguirá com a operação normal.</span><span class="sxs-lookup"><span data-stu-id="4e425-147">If the user is already registered and configured, the add-in continues with normal operation.</span></span>
   - <span data-ttu-id="4e425-148">Se a configuração do usuário for exigida, o suplemento entrará em modo "configuração" e solicitará que o usuário autorize o suplemento.</span><span class="sxs-lookup"><span data-stu-id="4e425-148">If user setup is required, the add-in enters "setup" mode and prompts the user to authorize the add-in.</span></span>

### <a name="authorize-the-backend-web-api"></a><span data-ttu-id="4e425-149">Autorizar a API Web de back-end</span><span class="sxs-lookup"><span data-stu-id="4e425-149">Authorize the backend Web API</span></span>

<span data-ttu-id="4e425-150">Para minimizar a necessidade de ter que informar o usuário de fazer login, o ideal é que o procedimento para autorizar a API Web de back-end a chamar a API do Microsoft Graph e a API de Dados da Contoso ocorra apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="4e425-150">The procedure for authorizing the backend Web API to call the Microsoft Graph API and Contoso Data API should ideally only have to happen once, to minimize having to prompt the user for sign-in.</span></span>

<span data-ttu-id="4e425-151">Com base na resposta da API Web de back-end, talvez o suplemento precise da autorização do usuário da API do Microsoft Graph, da API de Dados da Contoso ou de ambas APIs.</span><span class="sxs-lookup"><span data-stu-id="4e425-151">Based on the response from the backend Web API, the add-in may need to authorize the user for the Microsoft Graph API, the Contoso Data API, or both.</span></span> <span data-ttu-id="4e425-152">Como as duas APIs usam a autenticação OAuth2, o método é semelhante para ambas.</span><span class="sxs-lookup"><span data-stu-id="4e425-152">Since both APIs use OAuth2 authentication, the method is similar for both.</span></span>

1. <span data-ttu-id="4e425-153">O suplemento informa o usuário que precisa que ele autorize o uso da API e pede a ele para clicar em um link ou em um botão para iniciar o processo.</span><span class="sxs-lookup"><span data-stu-id="4e425-153">The add-in notifies the user that it needs them to authorize their use of the API and asks them to click a link or button to start the process.</span></span>

1. <span data-ttu-id="4e425-154">O suplemento usa a [API da Caixa de Diálogo](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) ou a [biblioteca office-js-helpers](https://github.com/OfficeDev/office-js-helpers) para iniciar o [Fluxo do Código de Autorização de OAuth2](/azure/active-directory/develop/active-directory-protocols-oauth-code) para a API.</span><span class="sxs-lookup"><span data-stu-id="4e425-154">The add-in uses the [Dialog API](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) or the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers) to start the [OAuth2 Authorization Code flow](/azure/active-directory/develop/active-directory-protocols-oauth-code) for the API.</span></span>

1. <span data-ttu-id="4e425-155">Após a conclusão do fluxo, o suplemento envia o token de atualização à API Web de back-end e inclui o token SSO (se disponível) ou o token de identidade do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e425-155">Once the flow completes, the add-in sends the refresh token to the backend Web API and includes the SSO token (if available) or the Exchange identity token.</span></span>

1. <span data-ttu-id="4e425-156">A API Web de back-end localiza o usuário no banco de dados e atualiza o token de atualização apropriado.</span><span class="sxs-lookup"><span data-stu-id="4e425-156">The backend Web API locates the user in the database and updates the appropriate refresh token.</span></span>

1. <span data-ttu-id="4e425-157">O suplemento prossegue com a operação normal.</span><span class="sxs-lookup"><span data-stu-id="4e425-157">The add-in continues with normal operation.</span></span>

### <a name="normal-operation"></a><span data-ttu-id="4e425-158">Operação normal</span><span class="sxs-lookup"><span data-stu-id="4e425-158">Normal operation</span></span>

<span data-ttu-id="4e425-159">Sempre que o suplemento chamar a API Web de back-end, incluirá o token SSO ou o token de identidade do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4e425-159">Whenever the add-in calls the backend Web API, it includes either the SSO token or the Exchange identity token.</span></span> <span data-ttu-id="4e425-160">A API Web de back-end localiza o usuário pelo token e usa os tokens de atualização armazenados para obter tokens de acesso da API do Microsoft Graph e da API de Dados da Contoso.</span><span class="sxs-lookup"><span data-stu-id="4e425-160">The backend Web API locates the user by this token, then uses the stored refresh tokens to obtain access tokens for the Microsoft Graph API and the Contoso Data API.</span></span> <span data-ttu-id="4e425-161">Enquanto os tokens de atualização forem válidos, o usuário não terá que entrar novamente.</span><span class="sxs-lookup"><span data-stu-id="4e425-161">As long as the refresh tokens are valid, the user will not have to sign in again.</span></span>