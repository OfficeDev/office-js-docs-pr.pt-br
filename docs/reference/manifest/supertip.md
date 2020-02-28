---
title: Elemento Supertip no arquivo de manifesto
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325231"
---
# <a name="supertip"></a><span data-ttu-id="ab011-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="ab011-102">Supertip</span></span>

<span data-ttu-id="ab011-p101">Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="ab011-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ab011-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ab011-105">Child elements</span></span>

|  <span data-ttu-id="ab011-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="ab011-106">Element</span></span> |  <span data-ttu-id="ab011-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ab011-107">Required</span></span>  |  <span data-ttu-id="ab011-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab011-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="ab011-109">Title</span><span class="sxs-lookup"><span data-stu-id="ab011-109">Title</span></span>](#title) | <span data-ttu-id="ab011-110">Sim</span><span class="sxs-lookup"><span data-stu-id="ab011-110">Yes</span></span> | <span data-ttu-id="ab011-111">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="ab011-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="ab011-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab011-112">Description</span></span>](#description) | <span data-ttu-id="ab011-113">Sim</span><span class="sxs-lookup"><span data-stu-id="ab011-113">Yes</span></span> | <span data-ttu-id="ab011-114">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="ab011-114">The description for the supertip.</span></span><br><span data-ttu-id="ab011-115">**Observação**: (Outlook) só há suporte para clientes Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="ab011-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="ab011-116">Cargo</span><span class="sxs-lookup"><span data-stu-id="ab011-116">Title</span></span>

<span data-ttu-id="ab011-117">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="ab011-117">Required.</span></span> <span data-ttu-id="ab011-118">O texto da superdica.</span><span class="sxs-lookup"><span data-stu-id="ab011-118">The text for the supertip.</span></span> <span data-ttu-id="ab011-119">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="ab011-119">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="ab011-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab011-120">Description</span></span>

<span data-ttu-id="ab011-121">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="ab011-121">Required.</span></span> <span data-ttu-id="ab011-122">A descrição da superdica.</span><span class="sxs-lookup"><span data-stu-id="ab011-122">The description for the supertip.</span></span> <span data-ttu-id="ab011-123">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **LongStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="ab011-123">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="ab011-124">Para o Outlook, apenas clientes Windows e Mac dão suporte ao elemento **Description** .</span><span class="sxs-lookup"><span data-stu-id="ab011-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="ab011-125">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab011-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
