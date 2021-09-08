---
title: Elemento Authorization no arquivo de manifesto
description: Especifica um recurso externo que o aplicativo Web do complemento precisa de autorização para e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937471"
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
