---
title: Elemento Scopes no arquivo de manifesto
description: O elemento de escopos contém permissões que o suplemento precisa para se conectar a um recurso externo.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612238"
---
# <a name="scopes-element"></a>Elemento Scopes

Contém permissões que o suplemento precisa para um recurso externo, como o Microsoft Graph. Quando o Microsoft Graph é o recurso, AppSource usa o elemento de escopos para criar uma caixa de diálogo de consentimento. Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.

**Escopos** é um elemento filho dos elementos [WebApplicationInfo](webapplicationinfo.md) e [Authorization](authorization.md) no manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Escopo**                |  Sim     |   O nome de uma permissão; por exemplo, files. Read. All ou Profile. |

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
