---
title: Elemento Scopes no arquivo de manifesto
description: O elemento Scopes contém permissões que o add-in precisa para se conectar a um recurso externo.
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 16e8a19a7aa73efa6aac00c915fde8d2b8647bad
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681532"
---
# <a name="scopes-element"></a>Elemento Scopes

Contém permissões que o complemento precisa para um recurso externo, como o Microsoft Graph. Quando o microsoft Graph é o recurso, o AppSource usa o elemento Scopes para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.

**Escopos** é um elemento filho do [elemento WebApplicationInfo](webapplicationinfo.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Escopo**                |  Sim     |   O nome de uma permissão; por exemplo, Files.Read.All ou profile. |

## <a name="example"></a>Exemplo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
