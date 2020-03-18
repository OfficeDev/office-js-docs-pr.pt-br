---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Superdica define uma dica de ferramenta rica (título e descrição).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720348"
---
# <a name="supertip"></a><span data-ttu-id="c317c-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="c317c-103">Supertip</span></span>

<span data-ttu-id="c317c-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c317c-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c317c-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="c317c-106">Child elements</span></span>

|  <span data-ttu-id="c317c-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="c317c-107">Element</span></span> |  <span data-ttu-id="c317c-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c317c-108">Required</span></span>  |  <span data-ttu-id="c317c-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="c317c-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c317c-110">Title</span><span class="sxs-lookup"><span data-stu-id="c317c-110">Title</span></span>](#title) | <span data-ttu-id="c317c-111">Sim</span><span class="sxs-lookup"><span data-stu-id="c317c-111">Yes</span></span> | <span data-ttu-id="c317c-112">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="c317c-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="c317c-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="c317c-113">Description</span></span>](#description) | <span data-ttu-id="c317c-114">Sim</span><span class="sxs-lookup"><span data-stu-id="c317c-114">Yes</span></span> | <span data-ttu-id="c317c-115">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="c317c-115">The description for the supertip.</span></span><br><span data-ttu-id="c317c-116">**Observação**: (Outlook) só há suporte para clientes Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="c317c-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="c317c-117">Cargo</span><span class="sxs-lookup"><span data-stu-id="c317c-117">Title</span></span>

<span data-ttu-id="c317c-118">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="c317c-118">Required.</span></span> <span data-ttu-id="c317c-119">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="c317c-119">The text for the supertip.</span></span> <span data-ttu-id="c317c-120">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c317c-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="c317c-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="c317c-121">Description</span></span>

<span data-ttu-id="c317c-122">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="c317c-122">Required.</span></span> <span data-ttu-id="c317c-123">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="c317c-123">The description for the supertip.</span></span> <span data-ttu-id="c317c-124">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **LongStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="c317c-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="c317c-125">Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .</span><span class="sxs-lookup"><span data-stu-id="c317c-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="c317c-126">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c317c-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
