---
title: Elemento Page no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452071"
---
# <a name="page-element"></a><span data-ttu-id="d7ee8-102">Elemento Page</span><span class="sxs-lookup"><span data-stu-id="d7ee8-102">Page element</span></span>

<span data-ttu-id="d7ee8-103">Define as configurações de página HTML usadas por uma função personalizada no Excel.</span><span class="sxs-lookup"><span data-stu-id="d7ee8-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d7ee8-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7ee8-104">Attributes</span></span>

<span data-ttu-id="d7ee8-105">Nenhuma</span><span class="sxs-lookup"><span data-stu-id="d7ee8-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="d7ee8-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="d7ee8-106">Child elements</span></span>

|  <span data-ttu-id="d7ee8-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="d7ee8-107">Element</span></span>  |  <span data-ttu-id="d7ee8-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d7ee8-108">Required</span></span>  |  <span data-ttu-id="d7ee8-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7ee8-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d7ee8-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d7ee8-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="d7ee8-111">Sim</span><span class="sxs-lookup"><span data-stu-id="d7ee8-111">Yes</span></span>  | <span data-ttu-id="d7ee8-112">Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d7ee8-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="d7ee8-113">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7ee8-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
