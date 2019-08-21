---
title: Elemento Scopes no arquivo de manifesto
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1e36bdcd0cdcaa8c842e924c2543d56bdc4e26a7
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477730"
---
# <a name="scopes-element"></a><span data-ttu-id="24845-102">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="24845-102">Scopes element</span></span>

<span data-ttu-id="24845-103">Contém permissões que o suplemento precisa para um recurso externo, como o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="24845-103">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="24845-104">Quando o Microsoft Graph é o recurso, AppSource usa o elemento de escopos para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="24845-104">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="24845-105">Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="24845-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="24845-106">\*\*\*\* Escopos é um elemento filho dos elementos [WebApplicationInfo](webapplicationinfo.md) e [Authorization](authorization.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="24845-106">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="24845-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="24845-107">Child elements</span></span>

|  <span data-ttu-id="24845-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="24845-108">Element</span></span> |  <span data-ttu-id="24845-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="24845-109">Required</span></span>  |  <span data-ttu-id="24845-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="24845-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="24845-111">**Escopo**</span><span class="sxs-lookup"><span data-stu-id="24845-111">**Scope**</span></span>                |  <span data-ttu-id="24845-112">Sim</span><span class="sxs-lookup"><span data-stu-id="24845-112">Yes</span></span>     |   <span data-ttu-id="24845-113">O nome de uma permissão; por exemplo, files. Read. All ou Profile.</span><span class="sxs-lookup"><span data-stu-id="24845-113">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="24845-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24845-114">Example</span></span>

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
