---
title: Elemento Scopes no arquivo de manifesto
description: O elemento Scopes contém permissões que o add-in precisa para se conectar a um recurso externo.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 05582ae05c13fae8e2272de3fe6111c5ff639f938a817fd0b50ad22e4234d033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087251"
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
