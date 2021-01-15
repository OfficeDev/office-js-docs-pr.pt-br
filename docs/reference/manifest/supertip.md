---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Superdica define uma dica de ferramenta rica (título e descrição).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771295"
---
# <a name="supertip"></a><span data-ttu-id="800ca-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="800ca-103">Supertip</span></span>

<span data-ttu-id="800ca-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="800ca-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="800ca-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="800ca-106">Child elements</span></span>

|  <span data-ttu-id="800ca-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="800ca-107">Element</span></span> |  <span data-ttu-id="800ca-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="800ca-108">Required</span></span>  |  <span data-ttu-id="800ca-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="800ca-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="800ca-110">Title</span><span class="sxs-lookup"><span data-stu-id="800ca-110">Title</span></span>](#title) | <span data-ttu-id="800ca-111">Sim</span><span class="sxs-lookup"><span data-stu-id="800ca-111">Yes</span></span> | <span data-ttu-id="800ca-112">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="800ca-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="800ca-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="800ca-113">Description</span></span>](#description) | <span data-ttu-id="800ca-114">Sim</span><span class="sxs-lookup"><span data-stu-id="800ca-114">Yes</span></span> | <span data-ttu-id="800ca-115">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="800ca-115">The description for the supertip.</span></span><br><span data-ttu-id="800ca-116">**Observação**: (Outlook) só há suporte para clientes Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="800ca-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="800ca-117">Título</span><span class="sxs-lookup"><span data-stu-id="800ca-117">Title</span></span>

<span data-ttu-id="800ca-118">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="800ca-118">Required.</span></span> <span data-ttu-id="800ca-119">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="800ca-119">The text for the supertip.</span></span> <span data-ttu-id="800ca-120">O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="800ca-120">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="800ca-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="800ca-121">Description</span></span>

<span data-ttu-id="800ca-122">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="800ca-122">Required.</span></span> <span data-ttu-id="800ca-123">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="800ca-123">The description for the supertip.</span></span> <span data-ttu-id="800ca-124">O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **LongStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="800ca-124">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="800ca-125">Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .</span><span class="sxs-lookup"><span data-stu-id="800ca-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="800ca-126">Exemplo</span><span class="sxs-lookup"><span data-stu-id="800ca-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
