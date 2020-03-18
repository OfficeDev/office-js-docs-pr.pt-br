---
title: Elemento Authorizations no arquivo de manifesto
description: Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 7ae0b9d0ec32a20846142a9fc89c48fe9cdf8053
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720656"
---
# <a name="authorizations-element"></a><span data-ttu-id="1580b-103">Elemento Authorizations</span><span class="sxs-lookup"><span data-stu-id="1580b-103">Authorizations element</span></span>

<span data-ttu-id="1580b-104">Especifica os recursos externos que o aplicativo Web do suplemento precisa de autorização e as permissões necessárias.</span><span class="sxs-lookup"><span data-stu-id="1580b-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="1580b-105">**Autorizações** é um elemento filho do elemento [WebApplicationInfo](webapplicationinfo.md) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="1580b-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1580b-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1580b-106">Child elements</span></span>

|  <span data-ttu-id="1580b-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="1580b-107">Element</span></span> |  <span data-ttu-id="1580b-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1580b-108">Required</span></span>  |  <span data-ttu-id="1580b-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="1580b-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1580b-110">Autorização</span><span class="sxs-lookup"><span data-stu-id="1580b-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="1580b-111">Sim</span><span class="sxs-lookup"><span data-stu-id="1580b-111">Yes</span></span>     |   <span data-ttu-id="1580b-112">Identifica um recurso externo para o qual o aplicativo Web do suplemento precisa de autorização e os escopos (permissões) necessários.</span><span class="sxs-lookup"><span data-stu-id="1580b-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="1580b-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1580b-113">Example</span></span>

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
