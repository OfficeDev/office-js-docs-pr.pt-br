---
title: Dentro do token de identidade do Exchange em um suplemento do Outlook
description: Saiba mais sobre o conteúdo de um token de identidade do usuário do Exchange gerado a partir de um suplemento do Outlook.
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7d586203395521deb14e18a3ae52b01459224b75
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546428"
---
# <a name="inside-the-exchange-identity-token"></a>Dentro do token de identidade do Exchange

O token de identidade do usuário do Exchange retornado pelo método [getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) oferece uma maneira do código do suplemento incluir a identidade do usuário com chamadas para o serviço de back-end. Este artigo discutirá o formato e o conteúdo do token.

Um token de identidade do usuário do Exchange é uma cadeia de caracteres codificada como URL em Base64 assinada pelo Exchange Server que a enviou. O token não é criptografado, e a chave pública que você usa para validar a assinatura é armazenada no Exchange Server que emitiu o token. O token tem três partes: um cabeçalho, uma carga e uma assinatura. Na cadeia de caracteres do token, as partes são separadas por um caractere de ponto (`.`) para facilitar a divisão do token para você

O Exchange usa um formato JWT (Token Web JSON) como token de identidade. Para saber mais sobre tokens JWT, confira [RFC 7519 Token Web JSON (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

## <a name="identity-token-header"></a>Cabeçalho de token de identidade

O cabeçalho fornece informações sobre o formato e informações de assinatura do token. O exemplo a seguir mostra a aparência do cabeçalho do token.

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
A tabela a seguir descreve as partes do cabeçalho do token.

| Declaração | Valor | Descrição |
|:-----|:-----|:-----|
| `typ` | `JWT` | Identifica o token como um Token Web JSON. Todos os tokens de identidade fornecidos pelo Exchange Server são tokens JWT. |
| `alg` | `RS256` | O algoritmo de hash que é usado para criar a assinatura. Todos os tokens fornecidos pelo Exchange Server usam o algoritmo de hash RSASSA-PKCS1-v1_5 com SHA-256. |
| `x5t` | Impressão digital de certificado | A impressão digital X. 509 do token. |

## <a name="identity-token-payload"></a>Carga de token de identidade

The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.

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
 
A tabela a seguir lista as partes da carga do token de identidade.

| Declaração | Descrição |
|:-----|:-----|
| `aud` | A URL do suplemento que solicitou o token. Um token só será válido se for enviado do suplemento está sendo executado no navegador do cliente. A URL do suplemento é especificada no manifesto. A marcação depende do tipo de manifesto.</br></br>**Manifesto XML:** Se o suplemento usar o esquema de manifestos de Suplementos do Office v1.1, essa URL será a URL **\<SourceLocation\>** especificada no primeiro elemento, `ItemRead` `ItemEdit`sob o tipo de formulário ou, o que ocorrer primeiro como parte do [elemento FormSettings](/javascript/api/manifest/formsettings) no manifesto do suplemento.</br></br>**Manifesto do Teams (versão prévia):** A URL é especificada na propriedade "extensions.audienceClaimUrl". |
| `iss` | Um identificador exclusivo para o Exchange Server que emitiu o token. Todos os tokens emitidos por esse Exchange Server terão o mesmo identificador. |
| `nbf` | The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970. |
| `exp` | The date and time that the token is valid until. The value is the number of seconds since January 1, 1970. |
| `appctxsender` | Um identificador exclusivo para o Exchange Server que enviou o contexto do aplicativo. |
| `isbrowserhostedapp` | Indica se o suplemento está hospedado em um navegador. |
| `appctx` | O contexto do aplicativo para o token. |

As informações na declaração appctx fornecem o identificador exclusivo da conta e a localização da chave pública usada para assinar o token. A tabela a seguir lista as partes da declaração `appctx`.

| Propriedade de contexto Application | Descrição |
|:-----|:-----|
| `msexchuid` | Um identificador exclusivo associado à conta de email e ao Exchange Server. |
| `version` | O número da versão do token. Para todos os tokens fornecidos pelo Exchange, o valor é `ExIdTok.V1`. |
| `amurl` | A URL do documento de metadados de autenticação que contém a chave pública do certificado x. 509 usado para assinar o token.<br/><br/>Para saber mais sobre como usar o documento de metadados de autenticação, confira [Validar um token de identidade do Exchange](validate-an-identity-token.md). |

## <a name="identity-token-signature"></a>Assinatura de token de identidade

The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.

## <a name="see-also"></a>Confira também

Para obter um exemplo que analisa o token de identidade do usuário do Exchange, confira [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).
