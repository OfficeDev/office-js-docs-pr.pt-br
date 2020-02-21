---
title: Dentro do token de identidade do Exchange em um suplemento do Outlook
description: Saiba mais sobre o conteúdo de um token de identidade do usuário do Exchange gerado a partir de um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4cbbcdc587495a9b490f300414235cba1c5c570a
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165812"
---
# <a name="inside-the-exchange-identity-token"></a>Dentro do token de identidade do Exchange

O token de identidade do usuário do Exchange retornado pelo método [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) oferece uma maneira do código do suplemento incluir a identidade do usuário com chamadas para o serviço de back-end. Este artigo discutirá o formato e o conteúdo do token.

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

A carga contém as declarações de autenticação que identificam a conta de email e identificam o Exchange Server que enviou o token. O exemplo a seguir mostra a aparência de seção de carga.

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
| `aud` | A URL do suplemento que solicitou o token. Um token só será válido se for enviado do suplemento está sendo executado no navegador do cliente. Se o suplemento usa o esquema de manifestos v1.1 de Suplementos do Office, essa URL é a URL especificada no primeiro elemento `SourceLocation`, no tipo de formulário `ItemRead` ou `ItemEdit`, o que ocorrer primeiro como parte do elemento [FormSettings](../reference/manifest/formsettings.md) no manifesto do suplemento. |
| `iss` | Um identificador exclusivo para o Exchange Server que emitiu o token. Todos os tokens emitidos por esse Exchange Server terão o mesmo identificador. |
| `nbf` | A data e a hora do início da validade do token. O valor é o número de segundos desde 1º de janeiro de 1970. |
| `exp` | A data e a hora de validade do token. O valor é o número de segundos desde 1º de janeiro de 1970. |
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

A assinatura é criada pelo hash das seções de cabeçalho e carga com o algoritmo especificado no cabeçalho e usando o certificado X509 autoassinado localizado no servidor no local especificado na carga. Seu serviço Web pode validar essa assinatura para ajudar a garantir que o token de identidade é proveniente do servidor esperado.

## <a name="see-also"></a>Confira também

Para obter um exemplo que analisa o token de identidade do usuário do Exchange, confira [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).
