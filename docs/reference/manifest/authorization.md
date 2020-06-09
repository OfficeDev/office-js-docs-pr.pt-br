---
title: Elemento Authorization no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8c6249706b8eef11f579378fe5c9dc83016d17c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608758"
---
# <a name="authorization-element"></a><span data-ttu-id="72ef4-103">Elemento Authorization</span><span class="sxs-lookup"><span data-stu-id="72ef4-103">Authorization element</span></span>

<span data-ttu-id="72ef4-104">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="72ef4-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="72ef4-105">**Authorization** é um elemento filho do elemento [Authorizations](authorizations.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="72ef4-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="72ef4-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="72ef4-106">Child elements</span></span>

|  <span data-ttu-id="72ef4-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="72ef4-107">Element</span></span> |  <span data-ttu-id="72ef4-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="72ef4-108">Required</span></span>  |  <span data-ttu-id="72ef4-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="72ef4-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="72ef4-110">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="72ef4-110">**Resource**</span></span>  |  <span data-ttu-id="72ef4-111">Sim</span><span class="sxs-lookup"><span data-stu-id="72ef4-111">Yes</span></span>   |  <span data-ttu-id="72ef4-112">Especifica a URL do recurso externo.</span><span class="sxs-lookup"><span data-stu-id="72ef4-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="72ef4-113">Escopos</span><span class="sxs-lookup"><span data-stu-id="72ef4-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="72ef4-114">Sim</span><span class="sxs-lookup"><span data-stu-id="72ef4-114">Yes</span></span>  |  <span data-ttu-id="72ef4-115">Especifica as permissões que o suplemento precisa para o recurso.</span><span class="sxs-lookup"><span data-stu-id="72ef4-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="72ef4-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="72ef4-116">Example</span></span>

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
