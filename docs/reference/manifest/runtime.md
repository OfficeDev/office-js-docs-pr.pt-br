---
title: Tempo de execução no arquivo de manifesto
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120632"
---
# <a name="runtime-element"></a><span data-ttu-id="57538-102">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="57538-102">Runtime element</span></span>

<span data-ttu-id="57538-103">Este recurso está em visualização.</span><span class="sxs-lookup"><span data-stu-id="57538-103">This feature is in preview.</span></span> <span data-ttu-id="57538-104">Elemento filho do [`<Runtimes>`](runtime.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="57538-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="57538-105">Este elemento facilita o compartilhamento de dados globais e chamadas de função entre as funções personalizadas do Excel e o painel de tarefas do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="57538-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="57538-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="57538-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="57538-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="57538-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="57538-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="57538-108">Contained in</span></span>

<span data-ttu-id="57538-109">-[Tempos](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="57538-109">-[Runtimes](runtimes.md)</span></span>

## <a name="attributes"></a><span data-ttu-id="57538-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="57538-110">Attributes</span></span>

|  <span data-ttu-id="57538-111">Atributo</span><span class="sxs-lookup"><span data-stu-id="57538-111">Attribute</span></span>  |  <span data-ttu-id="57538-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57538-112">Required</span></span>  |  <span data-ttu-id="57538-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="57538-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="57538-114">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="57538-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="57538-115">Sim</span><span class="sxs-lookup"><span data-stu-id="57538-115">Yes</span></span>  | <span data-ttu-id="57538-116">Deve sempre ser listado como longo se você quiser que as funções personalizadas do Excel funcionem enquanto o painel de tarefas do seu suplemento estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="57538-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="57538-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="57538-117">**resid**</span></span>  |  <span data-ttu-id="57538-118">Sim</span><span class="sxs-lookup"><span data-stu-id="57538-118">Yes</span></span>  | <span data-ttu-id="57538-119">Se usado para funções personalizadas do Excel, `resid` o deve apontar `TaskPaneAndCustomFunction.Url`para.</span><span class="sxs-lookup"><span data-stu-id="57538-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="57538-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="57538-120">See also</span></span>

<span data-ttu-id="57538-121">-[Tempo](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="57538-121">-[Runtime](runtime.md)</span></span>
