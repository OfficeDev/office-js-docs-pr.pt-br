---
title: Elemento Supertip no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433667"
---
# <a name="supertip"></a><span data-ttu-id="7c171-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="7c171-102">Supertip</span></span>

<span data-ttu-id="7c171-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="7c171-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7c171-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="7c171-105">Child elements</span></span>

|  <span data-ttu-id="7c171-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="7c171-106">Element</span></span> |  <span data-ttu-id="7c171-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7c171-107">Required</span></span>  |  <span data-ttu-id="7c171-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c171-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7c171-109">Título</span><span class="sxs-lookup"><span data-stu-id="7c171-109">Title</span></span>](#title)        | <span data-ttu-id="7c171-110">Sim</span><span class="sxs-lookup"><span data-stu-id="7c171-110">Yes</span></span> |   <span data-ttu-id="7c171-111">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="7c171-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="7c171-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c171-112">Description</span></span>](#description)  | <span data-ttu-id="7c171-113">Sim</span><span class="sxs-lookup"><span data-stu-id="7c171-113">Yes</span></span> |  <span data-ttu-id="7c171-114">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="7c171-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="7c171-115">Title</span><span class="sxs-lookup"><span data-stu-id="7c171-115">Title</span></span>

<span data-ttu-id="7c171-p102">Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="7c171-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="7c171-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c171-119">Description</span></span>

<span data-ttu-id="7c171-p103">Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **LongStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="7c171-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="7c171-123">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c171-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
