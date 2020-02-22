---
title: Validar um token de identidade de suplementos do Outlook
description: O suplemento do Outlook pode enviar um token de identidade do usuário do Exchange, mas, antes de você confiar na solicitação, deve validar o token para garantir que tenha sido enviado pelo servidor Exchange solicitado.
ms.date: 11/07/2019
localization_priority: Normal
ms.openlocfilehash: b412756a980d54a20a1c8deab43cd7634c0188cb
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165797"
---
# <a name="validate-an-exchange-identity-token"></a><span data-ttu-id="3feda-103">Validar um token de identidade do Exchange</span><span class="sxs-lookup"><span data-stu-id="3feda-103">Validate an Exchange identity token</span></span>

<span data-ttu-id="3feda-104">O suplemento do Outlook pode enviar um token de identidade do usuário do Exchange, mas, antes de você confiar na solicitação, deve validar o token para garantir que tenha sido enviado pelo servidor Exchange solicitado.</span><span class="sxs-lookup"><span data-stu-id="3feda-104">Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.</span></span> <span data-ttu-id="3feda-105">Os tokens de identidade do usuário do Exchange são JSON Web Tokens (JWT).</span><span class="sxs-lookup"><span data-stu-id="3feda-105">Exchange user identity tokens are JSON Web Tokens (JWT).</span></span> <span data-ttu-id="3feda-106">As etapas necessárias para validar um JWT estão descritas em [RFC 7519 Token Web JSON (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span><span class="sxs-lookup"><span data-stu-id="3feda-106">The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

<span data-ttu-id="3feda-107">Sugerimos que você use um processo de quatro etapas para validar o token de identidade e obter o identificador exclusivo do usuário.</span><span class="sxs-lookup"><span data-stu-id="3feda-107">We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier.</span></span> <span data-ttu-id="3feda-108">Em primeiro lugar, extraia o Token Web JSON (JWT) de uma cadeia de caracteres codificada como URL em Base64.</span><span class="sxs-lookup"><span data-stu-id="3feda-108">First, extract the JSON Web Token (JWT) from a base64 URL-encoded string.</span></span> <span data-ttu-id="3feda-109">Em segundo lugar, verifique se o token foi bem elaborado, se foi criado para um suplemento do Outlook e se não expirou. Além disso, verifique se é possível extrair uma URL válida para o documento dos metadados de autenticação.</span><span class="sxs-lookup"><span data-stu-id="3feda-109">Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document.</span></span> <span data-ttu-id="3feda-110">Em seguida, recupere o documento dos metadados de autenticação do servidor Exchange e valide a assinatura anexada ao token de identidade.</span><span class="sxs-lookup"><span data-stu-id="3feda-110">Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token.</span></span> <span data-ttu-id="3feda-111">Por fim, calcule um identificador exclusivo para o usuário concatenando a ID do Exchange do usuário com a URL do documento de metadados de autenticação.</span><span class="sxs-lookup"><span data-stu-id="3feda-111">Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.</span></span>

## <a name="extract-the-json-web-token"></a><span data-ttu-id="3feda-112">Extrair o Token Web JSON</span><span class="sxs-lookup"><span data-stu-id="3feda-112">Extract the JSON Web Token</span></span>

<span data-ttu-id="3feda-113">O token retornado de [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) é uma representação da cadeia de caracteres codificada do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-113">The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token.</span></span> <span data-ttu-id="3feda-114">Neste formulário, de acordo com o 7519 RFC, todos os JWTs têm três partes separadas por um ponto.</span><span class="sxs-lookup"><span data-stu-id="3feda-114">In this form, per RFC 7519, all JWTs have three parts, separated by a period.</span></span> <span data-ttu-id="3feda-115">O formato deve ser o seguinte.</span><span class="sxs-lookup"><span data-stu-id="3feda-115">The format is as follows.</span></span>

```json
{header}.{payload}.{signature}
```

<span data-ttu-id="3feda-116">O cabeçalho e a carga devem estar decodificados em Base64 para obter uma representação JSON de todas as partes.</span><span class="sxs-lookup"><span data-stu-id="3feda-116">The header and payload should be base64-decoded to obtain a JSON representation of each part.</span></span> <span data-ttu-id="3feda-117">A assinatura deverá estar codificada em Base64 para obter uma matriz de bytes contendo a assinatura binária.</span><span class="sxs-lookup"><span data-stu-id="3feda-117">The signature should be base64-decoded to obtain a byte array containing the binary signature.</span></span>

<span data-ttu-id="3feda-118">Para saber mais sobre o conteúdo do token, confira [Dentro do token de identidade do Exchange](inside-the-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="3feda-118">For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).</span></span>

<span data-ttu-id="3feda-119">Depois que tiver os três componentes decodificados, prossiga com a validação do conteúdo do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-119">After you have the three decoded components, you can proceed with validating the content of the token.</span></span>

## <a name="validate-token-contents"></a><span data-ttu-id="3feda-120">Validar o conteúdo do token</span><span class="sxs-lookup"><span data-stu-id="3feda-120">Validate token contents</span></span>

<span data-ttu-id="3feda-121">Para validar o conteúdo do token, verifique o seguinte.</span><span class="sxs-lookup"><span data-stu-id="3feda-121">To validate the token contents, you should check the following.</span></span>

- <span data-ttu-id="3feda-122">Verifique o cabeçalho e verifique se:</span><span class="sxs-lookup"><span data-stu-id="3feda-122">Check the header and verify that the:</span></span>
    - <span data-ttu-id="3feda-123">`typ`a declaração está definida `JWT`como.</span><span class="sxs-lookup"><span data-stu-id="3feda-123">`typ` claim is set to `JWT`.</span></span>
    - <span data-ttu-id="3feda-124">`alg`a declaração está definida `RS256`como.</span><span class="sxs-lookup"><span data-stu-id="3feda-124">`alg` claim is set to `RS256`.</span></span>
    - <span data-ttu-id="3feda-125">`x5t`a declaração está presente.</span><span class="sxs-lookup"><span data-stu-id="3feda-125">`x5t` claim is present.</span></span>

- <span data-ttu-id="3feda-126">Verifique a carga e verifique se:</span><span class="sxs-lookup"><span data-stu-id="3feda-126">Check the payload and verify that the:</span></span>
    - <span data-ttu-id="3feda-127">`amurl`a declaração dentro `appctx` do é definida como o local de um arquivo de manifesto de chave de assinatura de token autorizado.</span><span class="sxs-lookup"><span data-stu-id="3feda-127">`amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file.</span></span> <span data-ttu-id="3feda-128">Por exemplo, o valor `amurl` esperado para o Office 365 https://outlook.office365.com:443/autodiscover/metadata/json/1é.</span><span class="sxs-lookup"><span data-stu-id="3feda-128">For example, the expected `amurl` value for Office 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1.</span></span> <span data-ttu-id="3feda-129">Consulte a próxima seção [Verifique o domínio](#verify-the-domain) para obter informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="3feda-129">See the next section [Verify the domain](#verify-the-domain) for additional information.</span></span>
    - <span data-ttu-id="3feda-130">A hora atual está entre as horas especificadas nas `nbf` declarações `exp` e.</span><span class="sxs-lookup"><span data-stu-id="3feda-130">Current time is between the times specified in the `nbf` and `exp` claims.</span></span> <span data-ttu-id="3feda-131">A declaração `nbf` especifica a primeira hora que o token é considerado válido e a declaração `exp` especifica a hora de expiração do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-131">The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token.</span></span> <span data-ttu-id="3feda-132">Isso é recomendável para permitir algumas variações nas configurações do relógio entre servidores.</span><span class="sxs-lookup"><span data-stu-id="3feda-132">It is recommended to allow for some variation in clock settings between servers.</span></span>
    - <span data-ttu-id="3feda-133">`aud`Claim é a URL esperada para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3feda-133">`aud` claim is the expected URL for your add-in.</span></span>
    - <span data-ttu-id="3feda-134">`version`a declaração dentro `appctx` da declaração é definida `ExIdTok.V1`como.</span><span class="sxs-lookup"><span data-stu-id="3feda-134">`version` claim inside the `appctx` claim is set to `ExIdTok.V1`.</span></span>

### <a name="verify-the-domain"></a><span data-ttu-id="3feda-135">Verificar o domínio</span><span class="sxs-lookup"><span data-stu-id="3feda-135">Verify the domain</span></span>

<span data-ttu-id="3feda-136">Ao implementar a lógica de verificação descrita anteriormente nesta seção, você também deve exigir que o domínio da `amurl` declaração corresponda ao domínio de descoberta automática do usuário.</span><span class="sxs-lookup"><span data-stu-id="3feda-136">When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user.</span></span> <span data-ttu-id="3feda-137">Para fazer isso, você precisará usar ou implementar a descoberta automática.</span><span class="sxs-lookup"><span data-stu-id="3feda-137">To do so, you'll need to use or implement Autodiscover.</span></span> <span data-ttu-id="3feda-138">Para saber mais, você pode começar com a [descoberta automática do Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span><span class="sxs-lookup"><span data-stu-id="3feda-138">To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span></span>

## <a name="validate-the-identity-token-signature"></a><span data-ttu-id="3feda-139">Validar a assinatura do token de identidade</span><span class="sxs-lookup"><span data-stu-id="3feda-139">Validate the identity token signature</span></span>

<span data-ttu-id="3feda-140">Após saber que o JWT contém as declarações necessárias, prossiga com a validação da assinatura do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-140">After you know that the JWT contains the required claims, you can proceed with validating the token signature.</span></span>

### <a name="retrieve-the-public-signing-key"></a><span data-ttu-id="3feda-141">Recuperar a chave de assinatura pública</span><span class="sxs-lookup"><span data-stu-id="3feda-141">Retrieve the public signing key</span></span>

<span data-ttu-id="3feda-142">A primeira etapa é recuperar a chave pública que corresponde ao certificado que o servidor do Exchange usou para assinar o token.</span><span class="sxs-lookup"><span data-stu-id="3feda-142">The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token.</span></span> <span data-ttu-id="3feda-143">A chave está localizada no documento dos metadados de autenticação.</span><span class="sxs-lookup"><span data-stu-id="3feda-143">The key is found in the authentication metadata document.</span></span> <span data-ttu-id="3feda-144">Este documento é um arquivo JSON hospedado na URL especificada na declaração `amurl`.</span><span class="sxs-lookup"><span data-stu-id="3feda-144">This document is a JSON file hosted at the URL specified in the `amurl` claim.</span></span>

<span data-ttu-id="3feda-145">O documento dos metadados de autenticação utiliza o seguinte formato.</span><span class="sxs-lookup"><span data-stu-id="3feda-145">The authentication metadata document uses the following format.</span></span>

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

<span data-ttu-id="3feda-146">As teclas de assinatura disponíveis estão na matriz `keys`.</span><span class="sxs-lookup"><span data-stu-id="3feda-146">The available signing keys are in the `keys` array.</span></span> <span data-ttu-id="3feda-147">Escolha a chave correta, garantindo que o valor `x5t` na propriedade `keyinfo` corresponda ao valor `x5t` no cabeçalho do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-147">Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token.</span></span> <span data-ttu-id="3feda-148">A chave pública está dentro da propriedade `value` na propriedade `keyvalue` armazenada como uma matriz de bytes codificada em Base64.</span><span class="sxs-lookup"><span data-stu-id="3feda-148">The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.</span></span>

<span data-ttu-id="3feda-149">Quando você tiver a chave pública correta, verifique a assinatura.</span><span class="sxs-lookup"><span data-stu-id="3feda-149">After you have the correct public key, verify the signature.</span></span> <span data-ttu-id="3feda-150">Os dados assinados são as duas primeiras partes do token codificado, separados por um ponto:</span><span class="sxs-lookup"><span data-stu-id="3feda-150">The signed data is the first two parts of the encoded token, separated by a period:</span></span>

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a><span data-ttu-id="3feda-151">Calcular a ID exclusiva para uma conta do Exchange</span><span class="sxs-lookup"><span data-stu-id="3feda-151">Compute the unique ID for an Exchange account</span></span>

<span data-ttu-id="3feda-152">Você pode criar um identificador exclusivo para uma conta do Exchange, concatenando a URL do documento de metadados de autenticação com o identificador do Exchange para a conta.</span><span class="sxs-lookup"><span data-stu-id="3feda-152">You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account.</span></span> <span data-ttu-id="3feda-153">Com esse identificador exclusivo em mãos, é possível usá-lo para criar um sistema de logon único (SSO) para o serviço da Web de suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="3feda-153">When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service.</span></span> <span data-ttu-id="3feda-154">Para obter detalhes sobre como usar o identificador exclusivo para SSO, confira [Autenticar um usuário com um token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="3feda-154">For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="use-a-library-to-validate-the-token"></a><span data-ttu-id="3feda-155">Usar uma biblioteca para validar o token</span><span class="sxs-lookup"><span data-stu-id="3feda-155">Use a library to validate the token</span></span>

<span data-ttu-id="3feda-156">Há diversas bibliotecas que podem fazer a análise e validação de um JWT geral.</span><span class="sxs-lookup"><span data-stu-id="3feda-156">There are a number of libraries that can do general JWT parsing and validation.</span></span> <span data-ttu-id="3feda-157">A Microsoft fornece duas bibliotecas que podem ser usadas para validar tokens de identidade do usuário do Exchange.</span><span class="sxs-lookup"><span data-stu-id="3feda-157">Microsoft provides two libraries that can be used to validate Exchange user identity tokens.</span></span>

### <a name="systemidentitymodeltokensjwt"></a><span data-ttu-id="3feda-158">System.IdentityModel.Tokens.Jwt</span><span class="sxs-lookup"><span data-stu-id="3feda-158">System.IdentityModel.Tokens.Jwt</span></span>

<span data-ttu-id="3feda-159">A biblioteca [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) pode analisar o token e também fazer a validação necessária para analisar a declaração `appctx` por conta própria e recuperar a chave de assinatura pública.</span><span class="sxs-lookup"><span data-stu-id="3feda-159">The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.</span></span>

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

<span data-ttu-id="3feda-160">A classe `ExchangeAppContext` é definida da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="3feda-160">The `ExchangeAppContext` class is defined as follows:</span></span>

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

<span data-ttu-id="3feda-161">Para obter um exemplo que usa essa biblioteca para validar tokens do Exchange e tem uma implementação de `GetSigningKeys`, confira [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span><span class="sxs-lookup"><span data-stu-id="3feda-161">For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>

### <a name="microsoftexchangewebservices"></a><span data-ttu-id="3feda-162">Microsoft.Exchange.WebServices</span><span class="sxs-lookup"><span data-stu-id="3feda-162">Microsoft.Exchange.WebServices</span></span>

<span data-ttu-id="3feda-163">A [API Gerenciada dos Serviços Web do Exchange](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) também valida tokens de identidade do usuário do Exchange.</span><span class="sxs-lookup"><span data-stu-id="3feda-163">The [Exchange Web Services Managed API](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) can also validate Exchange user identity tokens.</span></span> <span data-ttu-id="3feda-164">Como é específica do Exchange, implementa toda a lógica necessária para analisar a declaração `appctx` e verificar a versão do token.</span><span class="sxs-lookup"><span data-stu-id="3feda-164">Because it is Exchange-specific, it implements all of the necessary logic to parse the `appctx` claim and verify the token version.</span></span>

```cs
using Microsoft.Exchange.WebServices.Auth.Validation;

AppIdentityToken ValidateIdentityToken(string rawToken, string expectedAudience)
{
    try
    {
        AppIdentityToken appIdToken = AuthToken.Parse(rawToken) as AppIdentityToken;
        appIdToken.Validate(new Uri(expectedAudience));

        // No exception, validation succeeded
        return appIdToken;
    }
    catch (TokenValidationException ex)
    {
        throw new Exception(string.Format("Token validation failed: {0}", ex.Message));
    }
}
```

## <a name="see-also"></a><span data-ttu-id="3feda-165">Confira também</span><span class="sxs-lookup"><span data-stu-id="3feda-165">See also</span></span>

- [<span data-ttu-id="3feda-166">Outlook-Add-In-Token-Viewer</span><span class="sxs-lookup"><span data-stu-id="3feda-166">Outlook-Add-In-Token-Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="3feda-167">Outlook-Add-in-JavaScript-ValidateIdentityToken</span><span class="sxs-lookup"><span data-stu-id="3feda-167">Outlook-Add-in-JavaScript-ValidateIdentityToken</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)