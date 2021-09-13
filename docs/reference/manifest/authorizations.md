---
title: Elemento Authorizations no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do complemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 4b13e26f13fae6fefd579868df8b67dd94cb35c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151673"
---
# <a name="authorizations-element"></a>Elemento Authorizations

Especifica os recursos externos que o aplicativo Web do complemento precisa de autorização e as permissões necessárias.

**Autorizações** é um elemento filho do [elemento WebApplicationInfo](webapplicationinfo.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Autorização](authorization.md)                |  Sim     |   Identifica um recurso externo de que o aplicativo Web do complemento precisa de autorização e os escopos (permissões) necessários. |

## <a name="example"></a>Exemplo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
