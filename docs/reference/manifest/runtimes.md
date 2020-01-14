---
title: Tempos de execução no arquivo de manifesto
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111174"
---
# <a name="runtimes-element"></a><span data-ttu-id="86b89-102">Elemento de runtimes</span><span class="sxs-lookup"><span data-stu-id="86b89-102">Runtimes element</span></span>

<span data-ttu-id="86b89-103">Este recurso está em visualização.</span><span class="sxs-lookup"><span data-stu-id="86b89-103">This feature is in preview.</span></span> <span data-ttu-id="86b89-104">Especifica o tempo de execução do suplemento e permite que as funções personalizadas e o painel de tarefas compartilhem dados globais e façam chamadas de função entre si.</span><span class="sxs-lookup"><span data-stu-id="86b89-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="86b89-105">Deve seguir o `<Host>` elemento no seu arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="86b89-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="86b89-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="86b89-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="86b89-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="86b89-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="86b89-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="86b89-108">Child elements</span></span>

|  <span data-ttu-id="86b89-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="86b89-109">Element</span></span> |  <span data-ttu-id="86b89-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86b89-110">Required</span></span>  |  <span data-ttu-id="86b89-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="86b89-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="86b89-112">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="86b89-112">**Runtime**</span></span>     | <span data-ttu-id="86b89-113">Sim</span><span class="sxs-lookup"><span data-stu-id="86b89-113">Yes</span></span> |  <span data-ttu-id="86b89-114">O tempo de execução do suplemento, geralmente usado com funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="86b89-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="86b89-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="86b89-115">See also</span></span>

<span data-ttu-id="86b89-116">-[Tempos](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="86b89-116">-[Runtimes](runtimes.md)</span></span>
