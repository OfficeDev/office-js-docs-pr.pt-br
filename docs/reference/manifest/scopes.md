---
title: Elemento Scopes no arquivo de manifesto
description: O elemento Scopes contém permissões que o add-in precisa para se conectar a um recurso externo.
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 346a143fdba35a153229b00052a463f726fd9056
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152127"
---
# <a name="scopes-element"></a>Elemento Scopes

Contém permissões que o complemento precisa para um recurso externo, como o Microsoft Graph. Quando o microsoft Graph é o recurso, o AppSource usa o elemento Scopes para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.

**Escopos** é um elemento filho dos elementos [WebApplicationInfo](webapplicationinfo.md) e [Authorization](authorization.md) no manifesto.

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
