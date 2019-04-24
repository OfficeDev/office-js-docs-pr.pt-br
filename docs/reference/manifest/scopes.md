---
title: Elemento Scopes no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 903f7ff68313549234c07926cc63dc7e783ae400
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451938"
---
# <a name="scopes-element"></a><span data-ttu-id="6065d-102">Elemento Scopes</span><span class="sxs-lookup"><span data-stu-id="6065d-102">Scopes element</span></span>

<span data-ttu-id="6065d-103">Contém permissões para o Microsoft Graph de que o suplemento precisa.</span><span class="sxs-lookup"><span data-stu-id="6065d-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="6065d-104">Este elemento é usado pela Loja do Office para criar uma caixa de diálogo de consentimento.</span><span class="sxs-lookup"><span data-stu-id="6065d-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="6065d-105">Quando os usuários instalam o suplemento da Office Store, eles são solicitados a conceder ao suplemento permissões especificas para os dados do Microsoft Graph do usuário.</span><span class="sxs-lookup"><span data-stu-id="6065d-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6065d-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="6065d-106">Child elements</span></span>

|  <span data-ttu-id="6065d-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="6065d-107">Element</span></span> |  <span data-ttu-id="6065d-108">Tipo</span><span class="sxs-lookup"><span data-stu-id="6065d-108">Type</span></span>  |  <span data-ttu-id="6065d-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="6065d-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6065d-110">**Escopo**</span><span class="sxs-lookup"><span data-stu-id="6065d-110">**Scope**</span></span>                |  <span data-ttu-id="6065d-111">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6065d-111">string</span></span>     |   <span data-ttu-id="6065d-112">O nome de uma permissão para o Microsoft Graph; por exemplo, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="6065d-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="6065d-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6065d-113">Example</span></span>

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
