---
title: Elemento Scopes no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432638"
---
# <a name="scopes-element"></a>Elemento Scopes

Contém permissões para o Microsoft Graph de que o suplemento precisa. Este elemento é usado pela Loja do Office para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Tipo  |  Descrição  |
|:-----|:-----|:-----|
|  **Escopo**                |  cadeia de caracteres     |   O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All. |

## <a name="example"></a>Exemplo

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
