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
# <a name="validate-an-exchange-identity-token"></a>Validar um token de identidade do Exchange

O suplemento do Outlook pode enviar um token de identidade do usuário do Exchange, mas, antes de você confiar na solicitação, deve validar o token para garantir que tenha sido enviado pelo servidor Exchange solicitado. Os tokens de identidade do usuário do Exchange são JSON Web Tokens (JWT). As etapas necessárias para validar um JWT estão descritas em [RFC 7519 Token Web JSON (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

Sugerimos que você use um processo de quatro etapas para validar o token de identidade e obter o identificador exclusivo do usuário. Em primeiro lugar, extraia o Token Web JSON (JWT) de uma cadeia de caracteres codificada como URL em Base64. Em segundo lugar, verifique se o token foi bem elaborado, se foi criado para um suplemento do Outlook e se não expirou. Além disso, verifique se é possível extrair uma URL válida para o documento dos metadados de autenticação. Em seguida, recupere o documento dos metadados de autenticação do servidor Exchange e valide a assinatura anexada ao token de identidade. Por fim, calcule um identificador exclusivo para o usuário concatenando a ID do Exchange do usuário com a URL do documento de metadados de autenticação.

## <a name="extract-the-json-web-token"></a>Extrair o Token Web JSON

O token retornado de [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) é uma representação da cadeia de caracteres codificada do token. Neste formulário, de acordo com o 7519 RFC, todos os JWTs têm três partes separadas por um ponto. O formato deve ser o seguinte.

```json
{header}.{payload}.{signature}
```

O cabeçalho e a carga devem estar decodificados em Base64 para obter uma representação JSON de todas as partes. A assinatura deverá estar codificada em Base64 para obter uma matriz de bytes contendo a assinatura binária.

Para saber mais sobre o conteúdo do token, confira [Dentro do token de identidade do Exchange](inside-the-identity-token.md).

Depois que tiver os três componentes decodificados, prossiga com a validação do conteúdo do token.

## <a name="validate-token-contents"></a>Validar o conteúdo do token

Para validar o conteúdo do token, verifique o seguinte.

- Verifique o cabeçalho e verifique se:
    - `typ`a declaração está definida `JWT`como.
    - `alg`a declaração está definida `RS256`como.
    - `x5t`a declaração está presente.

- Verifique a carga e verifique se:
    - `amurl`a declaração dentro `appctx` do é definida como o local de um arquivo de manifesto de chave de assinatura de token autorizado. Por exemplo, o valor `amurl` esperado para o Office 365 https://outlook.office365.com:443/autodiscover/metadata/json/1é. Consulte a próxima seção [Verifique o domínio](#verify-the-domain) para obter informações adicionais.
    - A hora atual está entre as horas especificadas nas `nbf` declarações `exp` e. A declaração `nbf` especifica a primeira hora que o token é considerado válido e a declaração `exp` especifica a hora de expiração do token. Isso é recomendável para permitir algumas variações nas configurações do relógio entre servidores.
    - `aud`Claim é a URL esperada para seu suplemento.
    - `version`a declaração dentro `appctx` da declaração é definida `ExIdTok.V1`como.

### <a name="verify-the-domain"></a>Verificar o domínio

Ao implementar a lógica de verificação descrita anteriormente nesta seção, você também deve exigir que o domínio da `amurl` declaração corresponda ao domínio de descoberta automática do usuário. Para fazer isso, você precisará usar ou implementar a descoberta automática. Para saber mais, você pode começar com a [descoberta automática do Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).

## <a name="validate-the-identity-token-signature"></a>Validar a assinatura do token de identidade

Após saber que o JWT contém as declarações necessárias, prossiga com a validação da assinatura do token.

### <a name="retrieve-the-public-signing-key"></a>Recuperar a chave de assinatura pública

A primeira etapa é recuperar a chave pública que corresponde ao certificado que o servidor do Exchange usou para assinar o token. A chave está localizada no documento dos metadados de autenticação. Este documento é um arquivo JSON hospedado na URL especificada na declaração `amurl`.

O documento dos metadados de autenticação utiliza o seguinte formato.

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

As teclas de assinatura disponíveis estão na matriz `keys`. Escolha a chave correta, garantindo que o valor `x5t` na propriedade `keyinfo` corresponda ao valor `x5t` no cabeçalho do token. A chave pública está dentro da propriedade `value` na propriedade `keyvalue` armazenada como uma matriz de bytes codificada em Base64.

Quando você tiver a chave pública correta, verifique a assinatura. Os dados assinados são as duas primeiras partes do token codificado, separados por um ponto:

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a>Calcular a ID exclusiva para uma conta do Exchange

Você pode criar um identificador exclusivo para uma conta do Exchange, concatenando a URL do documento de metadados de autenticação com o identificador do Exchange para a conta. Com esse identificador exclusivo em mãos, é possível usá-lo para criar um sistema de logon único (SSO) para o serviço da Web de suplementos do Outlook. Para obter detalhes sobre como usar o identificador exclusivo para SSO, confira [Autenticar um usuário com um token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md).

## <a name="use-a-library-to-validate-the-token"></a>Usar uma biblioteca para validar o token

Há diversas bibliotecas que podem fazer a análise e validação de um JWT geral. A Microsoft fornece duas bibliotecas que podem ser usadas para validar tokens de identidade do usuário do Exchange.

### <a name="systemidentitymodeltokensjwt"></a>System.IdentityModel.Tokens.Jwt

A biblioteca [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) pode analisar o token e também fazer a validação necessária para analisar a declaração `appctx` por conta própria e recuperar a chave de assinatura pública.

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

A classe `ExchangeAppContext` é definida da seguinte maneira:

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

Para obter um exemplo que usa essa biblioteca para validar tokens do Exchange e tem uma implementação de `GetSigningKeys`, confira [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).

### <a name="microsoftexchangewebservices"></a>Microsoft.Exchange.WebServices

A [API Gerenciada dos Serviços Web do Exchange](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) também valida tokens de identidade do usuário do Exchange. Como é específica do Exchange, implementa toda a lógica necessária para analisar a declaração `appctx` e verificar a versão do token.

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

## <a name="see-also"></a>Confira também

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
