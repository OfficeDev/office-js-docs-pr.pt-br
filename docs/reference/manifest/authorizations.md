---
title: Elemento Authorizations no arquivo de manifesto
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6a271423ddd549431c2f580e2793faab3c49090e
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477954"
---
# <a name="authorizations-element"></a><span data-ttu-id="19989-102">Elemento Authorizations</span><span class="sxs-lookup"><span data-stu-id="19989-102">Authorizations element</span></span>

<span data-ttu-id="19989-103">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="19989-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="19989-104">**Autorizações** é um elemento filho do elemento [WebApplicationInfo](webapplicationinfo.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="19989-104">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="19989-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="19989-105">Child elements</span></span>

|  <span data-ttu-id="19989-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="19989-106">Element</span></span> |  <span data-ttu-id="19989-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="19989-107">Required</span></span>  |  <span data-ttu-id="19989-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="19989-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="19989-109">Autorização</span><span class="sxs-lookup"><span data-stu-id="19989-109">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="19989-110">Sim</span><span class="sxs-lookup"><span data-stu-id="19989-110">Yes</span></span>     |   <span data-ttu-id="19989-111">Identifica um recurso externo para o qual o aplicativo Web do suplemento precisa de autorização e os escopos (permissões) necessários.</span><span class="sxs-lookup"><span data-stu-id="19989-111">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="19989-112">Exemplo</span><span class="sxs-lookup"><span data-stu-id="19989-112">Example</span></span>

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
