---
title: Elemento Authorizations no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 675585f99fc6261a2145219d553f02b9f9abded3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608751"
---
# <a name="authorizations-element"></a><span data-ttu-id="bf8f3-103">Elemento Authorizations</span><span class="sxs-lookup"><span data-stu-id="bf8f3-103">Authorizations element</span></span>

<span data-ttu-id="bf8f3-104">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="bf8f3-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="bf8f3-105">**Autorizações** é um elemento filho do elemento [WebApplicationInfo](webapplicationinfo.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="bf8f3-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bf8f3-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="bf8f3-106">Child elements</span></span>

|  <span data-ttu-id="bf8f3-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="bf8f3-107">Element</span></span> |  <span data-ttu-id="bf8f3-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="bf8f3-108">Required</span></span>  |  <span data-ttu-id="bf8f3-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="bf8f3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bf8f3-110">Autorização</span><span class="sxs-lookup"><span data-stu-id="bf8f3-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="bf8f3-111">Sim</span><span class="sxs-lookup"><span data-stu-id="bf8f3-111">Yes</span></span>     |   <span data-ttu-id="bf8f3-112">Identifica um recurso externo para o qual o aplicativo Web do suplemento precisa de autorização e os escopos (permissões) necessários.</span><span class="sxs-lookup"><span data-stu-id="bf8f3-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="bf8f3-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bf8f3-113">Example</span></span>

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
