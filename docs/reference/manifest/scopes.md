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
# <a name="scopes-element"></a><span data-ttu-id="3dd65-103">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="3dd65-103">Scopes element</span></span>

<span data-ttu-id="3dd65-104">Contém permissões que o suplemento precisa para um recurso externo, como o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3dd65-104">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="3dd65-105">Quando o Microsoft Graph é o recurso, AppSource usa o elemento de escopos para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="3dd65-105">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="3dd65-106">Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="3dd65-106">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="3dd65-107">**Escopos** é um elemento filho dos elementos [WebApplicationInfo](webapplicationinfo.md) e [Authorization](authorization.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="3dd65-107">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3dd65-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3dd65-108">Child elements</span></span>

|  <span data-ttu-id="3dd65-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="3dd65-109">Element</span></span> |  <span data-ttu-id="3dd65-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3dd65-110">Required</span></span>  |  <span data-ttu-id="3dd65-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="3dd65-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3dd65-112">**Escopo**</span><span class="sxs-lookup"><span data-stu-id="3dd65-112">**Scope**</span></span>                |  <span data-ttu-id="3dd65-113">Sim</span><span class="sxs-lookup"><span data-stu-id="3dd65-113">Yes</span></span>     |   <span data-ttu-id="3dd65-114">O nome de uma permissão; por exemplo, files. Read. All ou Profile.</span><span class="sxs-lookup"><span data-stu-id="3dd65-114">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="3dd65-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dd65-115">Example</span></span>

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
