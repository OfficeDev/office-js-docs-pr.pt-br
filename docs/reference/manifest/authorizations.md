---
title: Elemento Authorizations no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do complemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 068e6753e2e8e947e5e6e3c0885e7cd006165660862a37346eea114abb81a9b8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092496"
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
