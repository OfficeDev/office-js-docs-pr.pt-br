---
title: Elemento WebApplicationInfo no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2ab06b7ec21bccf13039badcc94b9de0ce7b8600
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870272"
---
# <a name="webapplicationinfo-element"></a>Elemento WebApplicationInfo

Suporta o logon único (SSO) em Suplementos do Office. Este elemento contém informações sobre o suplemento como:

- Um *recurso* do OAuth 2.0 para o qual o aplicativo de hospedagem do Office pode precisar de permissões.
- Um *cliente* do OAuth 2.0 que pode exigir permissões para o Microsoft Graph.

> [!NOTE]
> Atualmente, a API de logon único tem suporte para Word, Excel, Outlook e PowerPoint. Para saber mais sobre os programas para os quais a API de logon único tem suporte no momento, consulte [Conjuntos de requisitos da IdentityAPI](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação moderna para o locatário do Office 365. Para saber como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** é um elemento filho do elemento [VersionOverrides](versionoverrides.md) no manifesto.  

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Id**    |  Sim   |  A **Id do Aplicativo** do serviço associado do suplemento conforme registrado no ponto de extremidade do Azure Active Directory (Azure AD) v 2.0.|
|  **Recurso**  |  Sim   |  Especifica o **URI da ID do Aplicativo** do suplemento, conforme registrado no ponto de extremidade do Azure Active Directory v 2.0.|
|  [Escopos](scopes.md)                |  Não  |  Especifica as permissões que seu suplemento precisa para o Microsoft Graph.  |

> [!NOTE] 
> Atualmente, é necessário que o recurso do seu suplemento corresponda ao seu host. O Office não solicitará um token para um suplemento, a menos que possa provar a propriedade, e hoje isso é feito hospedando o suplemento sob o nome de domínio totalmente qualificado do recurso.

## <a name="webapplicationinfo-example"></a>Exemplo de WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
