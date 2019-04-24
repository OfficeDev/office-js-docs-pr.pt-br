---
title: Elemento Supertip no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdbba342fa591ddff3faf94ecd63a4740fb904da
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450538"
---
# <a name="supertip"></a><span data-ttu-id="18ff4-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="18ff4-102">Supertip</span></span>

<span data-ttu-id="18ff4-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="18ff4-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="18ff4-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="18ff4-105">Child elements</span></span>

|  <span data-ttu-id="18ff4-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="18ff4-106">Element</span></span> |  <span data-ttu-id="18ff4-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="18ff4-107">Required</span></span>  |  <span data-ttu-id="18ff4-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="18ff4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="18ff4-109">Title</span><span class="sxs-lookup"><span data-stu-id="18ff4-109">Title</span></span>](#title)        | <span data-ttu-id="18ff4-110">Sim</span><span class="sxs-lookup"><span data-stu-id="18ff4-110">Yes</span></span> |   <span data-ttu-id="18ff4-111">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="18ff4-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="18ff4-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="18ff4-112">Description</span></span>](#description)  | <span data-ttu-id="18ff4-113">Sim</span><span class="sxs-lookup"><span data-stu-id="18ff4-113">Yes</span></span> |  <span data-ttu-id="18ff4-114">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="18ff4-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="18ff4-115">Title</span><span class="sxs-lookup"><span data-stu-id="18ff4-115">Title</span></span>

<span data-ttu-id="18ff4-p102">Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="18ff4-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="18ff4-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="18ff4-119">Description</span></span>

<span data-ttu-id="18ff4-p103">Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **LongStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="18ff4-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="18ff4-123">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18ff4-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
