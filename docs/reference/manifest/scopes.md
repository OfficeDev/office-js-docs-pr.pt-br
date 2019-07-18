---
title: Elemento Scopes no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771783"
---
# <a name="scopes-element"></a><span data-ttu-id="4df7a-102">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="4df7a-102">Scopes element</span></span>

<span data-ttu-id="4df7a-103">Contém permissões para o Microsoft Graph de que o suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="4df7a-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="4df7a-104">AppSource usa o elemento escopos para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="4df7a-104">AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="4df7a-105">Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="4df7a-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4df7a-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4df7a-106">Child elements</span></span>

|  <span data-ttu-id="4df7a-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="4df7a-107">Element</span></span> |  <span data-ttu-id="4df7a-108">Tipo</span><span class="sxs-lookup"><span data-stu-id="4df7a-108">Type</span></span>  |  <span data-ttu-id="4df7a-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="4df7a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4df7a-110">**Escopo**</span><span class="sxs-lookup"><span data-stu-id="4df7a-110">**Scope**</span></span>                |  <span data-ttu-id="4df7a-111">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4df7a-111">string</span></span>     |   <span data-ttu-id="4df7a-112">O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="4df7a-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="4df7a-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4df7a-113">Example</span></span>

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
