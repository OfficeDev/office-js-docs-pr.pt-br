---
title: Elemento Supertip no arquivo de manifesto
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659648"
---
# <a name="supertip"></a><span data-ttu-id="0c209-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="0c209-102">Supertip</span></span>

<span data-ttu-id="0c209-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="0c209-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0c209-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="0c209-105">Child elements</span></span>

|  <span data-ttu-id="0c209-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="0c209-106">Element</span></span> |  <span data-ttu-id="0c209-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0c209-107">Required</span></span>  |  <span data-ttu-id="0c209-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="0c209-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0c209-109">Title</span><span class="sxs-lookup"><span data-stu-id="0c209-109">Title</span></span>](#title) | <span data-ttu-id="0c209-110">Sim</span><span class="sxs-lookup"><span data-stu-id="0c209-110">Yes</span></span> | <span data-ttu-id="0c209-111">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="0c209-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="0c209-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="0c209-112">Description</span></span>](#description) | <span data-ttu-id="0c209-113">Sim</span><span class="sxs-lookup"><span data-stu-id="0c209-113">Yes</span></span> | <span data-ttu-id="0c209-114">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="0c209-114">The description for the supertip.</span></span><br><span data-ttu-id="0c209-115">**Observação**: (Outlook) só há suporte para clientes Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="0c209-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="0c209-116">Título</span><span class="sxs-lookup"><span data-stu-id="0c209-116">Title</span></span>

<span data-ttu-id="0c209-p102">Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0c209-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="0c209-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="0c209-120">Description</span></span>

<span data-ttu-id="0c209-p103">Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **LongStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0c209-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="0c209-124">Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .</span><span class="sxs-lookup"><span data-stu-id="0c209-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="0c209-125">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0c209-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
