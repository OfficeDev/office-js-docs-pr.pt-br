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
# <a name="scopes-element"></a><span data-ttu-id="c536b-102">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="c536b-102">Scopes element</span></span>

<span data-ttu-id="c536b-103">Contém permissões para o Microsoft Graph de que o suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="c536b-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="c536b-104">Este elemento é usado pela Loja do Office para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="c536b-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="c536b-105">Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="c536b-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c536b-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="c536b-106">Child elements</span></span>

|  <span data-ttu-id="c536b-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="c536b-107">Element</span></span> |  <span data-ttu-id="c536b-108">Tipo</span><span class="sxs-lookup"><span data-stu-id="c536b-108">Type</span></span>  |  <span data-ttu-id="c536b-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="c536b-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c536b-110">**Escopo**</span><span class="sxs-lookup"><span data-stu-id="c536b-110">**Scope**</span></span>                |  <span data-ttu-id="c536b-111">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c536b-111">string</span></span>     |   <span data-ttu-id="c536b-112">O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="c536b-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="c536b-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c536b-113">Example</span></span>

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
