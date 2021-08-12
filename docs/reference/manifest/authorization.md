---
title: Elemento Authorization no arquivo de manifesto
description: Especifica um recurso externo que o aplicativo Web do complemento precisa de autorização para e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: af40a47c4ae30b6d18d3457704487027ff18ac92da2a3ae23cf1afe5c1e9b46a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087706"
---
# <a name="authorization-element"></a>Elemento Authorization

Especifica os recursos externos que o aplicativo Web do complemento precisa de autorização e as permissões necessárias.

**A** autorização é um elemento filho [do elemento Authorizations](authorizations.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Recurso**  |  Sim   |  Especifica a URL do recurso externo.|
|  [Escopos](scopes.md)                |  Sim  |  Especifica as permissões que o complemento precisa para o recurso.  |

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
