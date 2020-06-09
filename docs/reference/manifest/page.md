---
title: Elemento Page no arquivo de manifesto
description: O elemento de página define as configurações de página HTML que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611495"
---
# <a name="page-element"></a><span data-ttu-id="4c17a-103">Elemento Page</span><span class="sxs-lookup"><span data-stu-id="4c17a-103">Page element</span></span>

<span data-ttu-id="4c17a-104">Define as configurações de página HTML usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="4c17a-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="4c17a-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="4c17a-105">Attributes</span></span>

<span data-ttu-id="4c17a-106">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="4c17a-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="4c17a-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4c17a-107">Child elements</span></span>

|  <span data-ttu-id="4c17a-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="4c17a-108">Element</span></span>  |  <span data-ttu-id="4c17a-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4c17a-109">Required</span></span>  |  <span data-ttu-id="4c17a-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="4c17a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4c17a-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4c17a-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="4c17a-112">Sim</span><span class="sxs-lookup"><span data-stu-id="4c17a-112">Yes</span></span>  | <span data-ttu-id="4c17a-113">Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4c17a-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="4c17a-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4c17a-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
