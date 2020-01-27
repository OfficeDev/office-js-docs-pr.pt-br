---
title: Tempo de execução no arquivo de manifesto
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41553996"
---
# <a name="runtime-element"></a><span data-ttu-id="07a9b-102">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="07a9b-102">Runtime element</span></span>

<span data-ttu-id="07a9b-103">Este recurso está em visualização.</span><span class="sxs-lookup"><span data-stu-id="07a9b-103">This feature is in preview.</span></span> <span data-ttu-id="07a9b-104">Elemento filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="07a9b-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="07a9b-105">Este elemento facilita o compartilhamento de dados globais e chamadas de função entre as funções personalizadas do Excel e o painel de tarefas do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="07a9b-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="07a9b-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="07a9b-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="07a9b-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="07a9b-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="07a9b-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="07a9b-108">Contained in</span></span>

- [<span data-ttu-id="07a9b-109">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="07a9b-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="07a9b-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="07a9b-110">Attributes</span></span>

|  <span data-ttu-id="07a9b-111">Atributo</span><span class="sxs-lookup"><span data-stu-id="07a9b-111">Attribute</span></span>  |  <span data-ttu-id="07a9b-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="07a9b-112">Required</span></span>  |  <span data-ttu-id="07a9b-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="07a9b-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="07a9b-114">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="07a9b-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="07a9b-115">Sim</span><span class="sxs-lookup"><span data-stu-id="07a9b-115">Yes</span></span>  | <span data-ttu-id="07a9b-116">Deve sempre ser listado como longo se você quiser que as funções personalizadas do Excel funcionem enquanto o painel de tarefas do seu suplemento estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="07a9b-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="07a9b-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="07a9b-117">**resid**</span></span>  |  <span data-ttu-id="07a9b-118">Sim</span><span class="sxs-lookup"><span data-stu-id="07a9b-118">Yes</span></span>  | <span data-ttu-id="07a9b-119">Se usado para funções personalizadas do Excel, `resid` o deve apontar `TaskPaneAndCustomFunction.Url`para.</span><span class="sxs-lookup"><span data-stu-id="07a9b-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="07a9b-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="07a9b-120">See also</span></span>

- [<span data-ttu-id="07a9b-121">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="07a9b-121">Runtimes</span></span>](runtimes.md)
