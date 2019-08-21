---
title: Elemento Authorization no arquivo de manifesto
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477940"
---
# <a name="authorization-element"></a>Elemento Authorization

Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.

**Authorization** é um elemento filho do elemento [Authorizations](authorizations.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Recurso**  |  Sim   |  Especifica a URL do recurso externo.|
|  [Escopos](scopes.md)                |  Sim  |  Especifica as permissões que o suplemento precisa para o recurso.  |

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
