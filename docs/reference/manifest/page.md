---
title: Elemento Page no arquivo de manifesto
description: O elemento de página define as configurações de página HTML que uma função personalizada usa no Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720481"
---
# <a name="page-element"></a><span data-ttu-id="d6622-103">Elemento Page</span><span class="sxs-lookup"><span data-stu-id="d6622-103">Page element</span></span>

<span data-ttu-id="d6622-104">Define as configurações de página HTML usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="d6622-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d6622-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="d6622-105">Attributes</span></span>

<span data-ttu-id="d6622-106">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="d6622-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="d6622-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="d6622-107">Child elements</span></span>

|  <span data-ttu-id="d6622-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="d6622-108">Element</span></span>  |  <span data-ttu-id="d6622-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d6622-109">Required</span></span>  |  <span data-ttu-id="d6622-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="d6622-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d6622-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d6622-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="d6622-112">Sim</span><span class="sxs-lookup"><span data-stu-id="d6622-112">Yes</span></span>  | <span data-ttu-id="d6622-113">Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d6622-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="d6622-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d6622-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
