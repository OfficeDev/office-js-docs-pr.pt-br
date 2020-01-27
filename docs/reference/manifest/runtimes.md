---
title: Tempos de execução no arquivo de manifesto
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554003"
---
# <a name="runtimes-element"></a><span data-ttu-id="e2324-102">Elemento de runtimes</span><span class="sxs-lookup"><span data-stu-id="e2324-102">Runtimes element</span></span>

<span data-ttu-id="e2324-103">Este recurso está em visualização.</span><span class="sxs-lookup"><span data-stu-id="e2324-103">This feature is in preview.</span></span> <span data-ttu-id="e2324-104">Especifica o tempo de execução do suplemento e permite que as funções personalizadas e o painel de tarefas compartilhem dados globais e façam chamadas de função entre si.</span><span class="sxs-lookup"><span data-stu-id="e2324-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="e2324-105">Deve seguir o `<Host>` elemento no seu arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="e2324-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="e2324-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e2324-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="e2324-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e2324-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="e2324-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e2324-108">Child elements</span></span>

|  <span data-ttu-id="e2324-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="e2324-109">Element</span></span> |  <span data-ttu-id="e2324-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e2324-110">Required</span></span>  |  <span data-ttu-id="e2324-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="e2324-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e2324-112">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="e2324-112">**Runtime**</span></span>     | <span data-ttu-id="e2324-113">Sim</span><span class="sxs-lookup"><span data-stu-id="e2324-113">Yes</span></span> |  <span data-ttu-id="e2324-114">O tempo de execução do suplemento, geralmente usado com funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="e2324-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="e2324-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="e2324-115">See also</span></span>

- [<span data-ttu-id="e2324-116">Runtime</span><span class="sxs-lookup"><span data-stu-id="e2324-116">Runtime</span></span>](runtime.md)
