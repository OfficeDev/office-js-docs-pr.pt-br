---
title: Tempo de execução no arquivo de manifesto
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111167"
---
# <a name="runtime-element"></a><span data-ttu-id="b7c24-102">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="b7c24-102">Runtime element</span></span>

<span data-ttu-id="b7c24-103">Este recurso está em visualização.</span><span class="sxs-lookup"><span data-stu-id="b7c24-103">This feature is in preview.</span></span> <span data-ttu-id="b7c24-104">Elemento filho do [`<Runtimes>`](runtime.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="b7c24-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="b7c24-105">Este elemento facilita o compartilhamento de dados globais e chamadas de função entre as funções personalizadas do Excel e o painel de tarefas do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b7c24-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span> 

## <a name="contained-in"></a><span data-ttu-id="b7c24-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="b7c24-106">Contained in</span></span>

<span data-ttu-id="b7c24-107">-[Tempos](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="b7c24-107">-[Runtimes](runtimes.md)</span></span>

<span data-ttu-id="b7c24-108">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b7c24-108">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="b7c24-109">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b7c24-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a><span data-ttu-id="b7c24-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="b7c24-110">Attributes</span></span>

|  <span data-ttu-id="b7c24-111">Atributo</span><span class="sxs-lookup"><span data-stu-id="b7c24-111">Attribute</span></span>  |  <span data-ttu-id="b7c24-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b7c24-112">Required</span></span>  |  <span data-ttu-id="b7c24-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="b7c24-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b7c24-114">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="b7c24-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="b7c24-115">Sim</span><span class="sxs-lookup"><span data-stu-id="b7c24-115">Yes</span></span>  | <span data-ttu-id="b7c24-116">Deve sempre ser listado como longo se você quiser que as funções personalizadas do Excel funcionem enquanto o painel de tarefas do seu suplemento estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="b7c24-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="b7c24-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="b7c24-117">**resid**</span></span>  |  <span data-ttu-id="b7c24-118">Sim</span><span class="sxs-lookup"><span data-stu-id="b7c24-118">Yes</span></span>  | <span data-ttu-id="b7c24-119">Se usado para funções personalizadas do Excel, `resid` o deve apontar `TaskPaneAndCustomFunction.Url`para.</span><span class="sxs-lookup"><span data-stu-id="b7c24-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b7c24-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="b7c24-120">See also</span></span>

<span data-ttu-id="b7c24-121">-[Tempo](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="b7c24-121">-[Runtime](runtime.md)</span></span>
