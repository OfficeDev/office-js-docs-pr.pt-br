---
title: Elemento Authorization no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cece0934eb9db3175b173e97d7ab478827b7cda2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718437"
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
