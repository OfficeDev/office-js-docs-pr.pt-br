---
title: Elemento Authorization no arquivo de manifesto
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477940"
---
# <a name="authorization-element"></a><span data-ttu-id="f48fb-102">Elemento Authorization</span><span class="sxs-lookup"><span data-stu-id="f48fb-102">Authorization element</span></span>

<span data-ttu-id="f48fb-103">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="f48fb-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="f48fb-104">**Authorization** é um elemento filho do elemento [Authorizations](authorizations.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="f48fb-104">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f48fb-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f48fb-105">Child elements</span></span>

|  <span data-ttu-id="f48fb-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="f48fb-106">Element</span></span> |  <span data-ttu-id="f48fb-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f48fb-107">Required</span></span>  |  <span data-ttu-id="f48fb-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="f48fb-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f48fb-109">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="f48fb-109">**Resource**</span></span>  |  <span data-ttu-id="f48fb-110">Sim</span><span class="sxs-lookup"><span data-stu-id="f48fb-110">Yes</span></span>   |  <span data-ttu-id="f48fb-111">Especifica a URL do recurso externo.</span><span class="sxs-lookup"><span data-stu-id="f48fb-111">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="f48fb-112">Escopos</span><span class="sxs-lookup"><span data-stu-id="f48fb-112">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="f48fb-113">Sim</span><span class="sxs-lookup"><span data-stu-id="f48fb-113">Yes</span></span>  |  <span data-ttu-id="f48fb-114">Especifica as permissões que o suplemento precisa para o recurso.</span><span class="sxs-lookup"><span data-stu-id="f48fb-114">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="f48fb-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f48fb-115">Example</span></span>

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
