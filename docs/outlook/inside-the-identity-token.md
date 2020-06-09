---
title: Dentro do token de identidade do Exchange em um suplemento do Outlook
description: Saiba mais sobre o conteúdo de um token de identidade do usuário do Exchange gerado a partir de um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: dee8416660386c25a55caa42b6e5ee8685ee8852
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609087"
---
# <a name="inside-the-exchange-identity-token"></a><span data-ttu-id="77dff-103">Dentro do token de identidade do Exchange</span><span class="sxs-lookup"><span data-stu-id="77dff-103">Inside the Exchange identity token</span></span>

<span data-ttu-id="77dff-104">O token de identidade do usuário do Exchange retornado pelo método [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) oferece uma maneira do código do suplemento incluir a identidade do usuário com chamadas para o serviço de back-end.</span><span class="sxs-lookup"><span data-stu-id="77dff-104">The Exchange user identity token returned by the [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service.</span></span> <span data-ttu-id="77dff-105">Este artigo discutirá o formato e o conteúdo do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-105">This article will discuss the format and contents of the token.</span></span>

<span data-ttu-id="77dff-106">Um token de identidade do usuário do Exchange é uma cadeia de caracteres codificada como URL em Base64 assinada pelo Exchange Server que a enviou.</span><span class="sxs-lookup"><span data-stu-id="77dff-106">An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it.</span></span> <span data-ttu-id="77dff-107">O token não é criptografado, e a chave pública que você usa para validar a assinatura é armazenada no Exchange Server que emitiu o token.</span><span class="sxs-lookup"><span data-stu-id="77dff-107">The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token.</span></span> <span data-ttu-id="77dff-108">O token tem três partes: um cabeçalho, uma carga e uma assinatura.</span><span class="sxs-lookup"><span data-stu-id="77dff-108">The token has three parts: a header, a payload, and a signature.</span></span> <span data-ttu-id="77dff-109">Na cadeia de caracteres do token, as partes são separadas por um caractere de ponto (`.`) para facilitar a divisão do token para você</span><span class="sxs-lookup"><span data-stu-id="77dff-109">In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.</span></span>

<span data-ttu-id="77dff-110">O Exchange usa um formato JWT (Token Web JSON) como token de identidade.</span><span class="sxs-lookup"><span data-stu-id="77dff-110">Exchange uses a the JSON Web Token (JWT) format for the identity token.</span></span> <span data-ttu-id="77dff-111">Para saber mais sobre tokens JWT, confira [RFC 7519 Token Web JSON (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span><span class="sxs-lookup"><span data-stu-id="77dff-111">For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

## <a name="identity-token-header"></a><span data-ttu-id="77dff-112">Cabeçalho de token de identidade</span><span class="sxs-lookup"><span data-stu-id="77dff-112">Identity token header</span></span>

<span data-ttu-id="77dff-113">O cabeçalho fornece informações sobre o formato e informações de assinatura do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-113">The header provides information about the format and signature information of the token.</span></span> <span data-ttu-id="77dff-114">O exemplo a seguir mostra a aparência do cabeçalho do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-114">The following example shows what the header of the token looks like.</span></span>

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
<span data-ttu-id="77dff-115">A tabela a seguir descreve as partes do cabeçalho do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-115">The following table describes the parts of the token header.</span></span>

| <span data-ttu-id="77dff-116">Declaração</span><span class="sxs-lookup"><span data-stu-id="77dff-116">Claim</span></span> | <span data-ttu-id="77dff-117">Valor</span><span class="sxs-lookup"><span data-stu-id="77dff-117">Value</span></span> | <span data-ttu-id="77dff-118">Descrição</span><span class="sxs-lookup"><span data-stu-id="77dff-118">Description</span></span> |
|:-----|:-----|:-----|
| `typ` | `JWT` | <span data-ttu-id="77dff-119">Identifica o token como um Token Web JSON.</span><span class="sxs-lookup"><span data-stu-id="77dff-119">Identifies the token as a JSON Web Token.</span></span> <span data-ttu-id="77dff-120">Todos os tokens de identidade fornecidos pelo Exchange Server são tokens JWT.</span><span class="sxs-lookup"><span data-stu-id="77dff-120">All identity tokens provided by Exchange server are JWT tokens.</span></span> |
| `alg` | `RS256` | <span data-ttu-id="77dff-121">O algoritmo de hash que é usado para criar a assinatura.</span><span class="sxs-lookup"><span data-stu-id="77dff-121">The hashing algorithm that is used to create the signature.</span></span> <span data-ttu-id="77dff-122">Todos os tokens fornecidos pelo Exchange Server usam o algoritmo de hash RSASSA-PKCS1-v1_5 com SHA-256.</span><span class="sxs-lookup"><span data-stu-id="77dff-122">All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm.</span></span> |
| `x5t` | <span data-ttu-id="77dff-123">Impressão digital de certificado</span><span class="sxs-lookup"><span data-stu-id="77dff-123">Certificate thumbprint</span></span> | <span data-ttu-id="77dff-124">A impressão digital X. 509 do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-124">The X.509 thumbprint of the token.</span></span> |

## <a name="identity-token-payload"></a><span data-ttu-id="77dff-125">Carga de token de identidade</span><span class="sxs-lookup"><span data-stu-id="77dff-125">Identity token payload</span></span>

<span data-ttu-id="77dff-p107">A carga contém as declarações de autenticação que identificam a conta de email e identificam o Exchange Server que enviou o token. O exemplo a seguir mostra a aparência de seção de carga.</span><span class="sxs-lookup"><span data-stu-id="77dff-p107">The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.</span></span>

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
<span data-ttu-id="77dff-128">A tabela a seguir lista as partes da carga do token de identidade.</span><span class="sxs-lookup"><span data-stu-id="77dff-128">The following table lists the parts of the identity token payload.</span></span>

| <span data-ttu-id="77dff-129">Declaração</span><span class="sxs-lookup"><span data-stu-id="77dff-129">Claim</span></span> | <span data-ttu-id="77dff-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="77dff-130">Description</span></span> |
|:-----|:-----|
| `aud` | <span data-ttu-id="77dff-131">A URL do suplemento que solicitou o token.</span><span class="sxs-lookup"><span data-stu-id="77dff-131">The URL of the add-in that requested the token.</span></span> <span data-ttu-id="77dff-132">Um token só será válido se for enviado do suplemento está sendo executado no navegador do cliente.</span><span class="sxs-lookup"><span data-stu-id="77dff-132">A token is only valid if it is sent from the add-in that is running in the client's browser.</span></span> <span data-ttu-id="77dff-133">Se o suplemento usa o esquema de manifestos v1.1 de Suplementos do Office, essa URL é a URL especificada no primeiro elemento `SourceLocation`, no tipo de formulário `ItemRead` ou `ItemEdit`, o que ocorrer primeiro como parte do elemento [FormSettings](../reference/manifest/formsettings.md) no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="77dff-133">If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `SourceLocation` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](../reference/manifest/formsettings.md) element in the add-in manifest.</span></span> |
| `iss` | <span data-ttu-id="77dff-p109">Um identificador exclusivo para o Exchange Server que emitiu o token. Todos os tokens emitidos por esse Exchange Server terão o mesmo identificador.</span><span class="sxs-lookup"><span data-stu-id="77dff-p109">A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.</span></span> |
| `nbf` | <span data-ttu-id="77dff-p110">A data e a hora do início da validade do token. O valor é o número de segundos desde 1º de janeiro de 1970.</span><span class="sxs-lookup"><span data-stu-id="77dff-p110">The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970.</span></span> |
| `exp` | <span data-ttu-id="77dff-p111">A data e a hora de validade do token. O valor é o número de segundos desde 1º de janeiro de 1970.</span><span class="sxs-lookup"><span data-stu-id="77dff-p111">The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.</span></span> |
| `appctxsender` | <span data-ttu-id="77dff-140">Um identificador exclusivo para o Exchange Server que enviou o contexto do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="77dff-140">A unique identifier for the Exchange server that sent the application context.</span></span> |
| `isbrowserhostedapp` | <span data-ttu-id="77dff-141">Indica se o suplemento está hospedado em um navegador.</span><span class="sxs-lookup"><span data-stu-id="77dff-141">Indicates whether the add-in is hosted in a browser.</span></span> |
| `appctx` | <span data-ttu-id="77dff-142">O contexto do aplicativo para o token.</span><span class="sxs-lookup"><span data-stu-id="77dff-142">The application context for the token.</span></span> |

<span data-ttu-id="77dff-143">As informações na declaração appctx fornecem o identificador exclusivo da conta e a localização da chave pública usada para assinar o token.</span><span class="sxs-lookup"><span data-stu-id="77dff-143">The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token.</span></span> <span data-ttu-id="77dff-144">A tabela a seguir lista as partes da declaração `appctx`.</span><span class="sxs-lookup"><span data-stu-id="77dff-144">The following table lists the parts of the `appctx` claim.</span></span>

| <span data-ttu-id="77dff-145">Propriedade de contexto Application</span><span class="sxs-lookup"><span data-stu-id="77dff-145">Application context property</span></span> | <span data-ttu-id="77dff-146">Descrição</span><span class="sxs-lookup"><span data-stu-id="77dff-146">Description</span></span> |
|:-----|:-----|
| `msexchuid` | <span data-ttu-id="77dff-147">Um identificador exclusivo associado à conta de email e ao Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="77dff-147">A unique identifier associated with the email account and the Exchange server.</span></span> |
| `version` | <span data-ttu-id="77dff-148">O número da versão do token.</span><span class="sxs-lookup"><span data-stu-id="77dff-148">The version number of the token.</span></span> <span data-ttu-id="77dff-149">Para todos os tokens fornecidos pelo Exchange, o valor é `ExIdTok.V1`.</span><span class="sxs-lookup"><span data-stu-id="77dff-149">For all tokens provided by Exchange, the value is `ExIdTok.V1`.</span></span> |
| `amurl` | <span data-ttu-id="77dff-150">A URL do documento de metadados de autenticação que contém a chave pública do certificado x. 509 usado para assinar o token.</span><span class="sxs-lookup"><span data-stu-id="77dff-150">The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.</span></span><br/><br/><span data-ttu-id="77dff-151">Para saber mais sobre como usar o documento de metadados de autenticação, confira [Validar um token de identidade do Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="77dff-151">For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span> |

## <a name="identity-token-signature"></a><span data-ttu-id="77dff-152">Assinatura de token de identidade</span><span class="sxs-lookup"><span data-stu-id="77dff-152">Identity token signature</span></span>

<span data-ttu-id="77dff-p114">A assinatura é criada pelo hash das seções de cabeçalho e carga com o algoritmo especificado no cabeçalho e usando o certificado X509 autoassinado localizado no servidor no local especificado na carga. Seu serviço Web pode validar essa assinatura para ajudar a garantir que o token de identidade é proveniente do servidor esperado.</span><span class="sxs-lookup"><span data-stu-id="77dff-p114">The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.</span></span>

## <a name="see-also"></a><span data-ttu-id="77dff-155">Confira também</span><span class="sxs-lookup"><span data-stu-id="77dff-155">See also</span></span>

<span data-ttu-id="77dff-156">Para obter um exemplo que analisa o token de identidade do usuário do Exchange, confira [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span><span class="sxs-lookup"><span data-stu-id="77dff-156">For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>
