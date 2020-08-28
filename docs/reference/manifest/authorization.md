---
title: Elemento Authorization no arquivo de manifesto
description: Especifica um recurso externo para o qual o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294259"
---
# <a name="authorization-element"></a><span data-ttu-id="b07f3-103">Elemento Authorization</span><span class="sxs-lookup"><span data-stu-id="b07f3-103">Authorization element</span></span>

<span data-ttu-id="b07f3-104">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="b07f3-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="b07f3-105">**Authorization** é um elemento filho do elemento [Authorizations](authorizations.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="b07f3-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b07f3-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="b07f3-106">Child elements</span></span>

|  <span data-ttu-id="b07f3-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="b07f3-107">Element</span></span> |  <span data-ttu-id="b07f3-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b07f3-108">Required</span></span>  |  <span data-ttu-id="b07f3-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="b07f3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b07f3-110">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="b07f3-110">**Resource**</span></span>  |  <span data-ttu-id="b07f3-111">Sim</span><span class="sxs-lookup"><span data-stu-id="b07f3-111">Yes</span></span>   |  <span data-ttu-id="b07f3-112">Especifica a URL do recurso externo.</span><span class="sxs-lookup"><span data-stu-id="b07f3-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="b07f3-113">Escopos</span><span class="sxs-lookup"><span data-stu-id="b07f3-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="b07f3-114">Sim</span><span class="sxs-lookup"><span data-stu-id="b07f3-114">Yes</span></span>  |  <span data-ttu-id="b07f3-115">Especifica as permissões que o suplemento precisa para o recurso.</span><span class="sxs-lookup"><span data-stu-id="b07f3-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="b07f3-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b07f3-116">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
