---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para sua faixa de opções, painel de tarefas e funções personalizadas.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217757"
---
# <a name="runtime-element"></a><span data-ttu-id="ae129-103">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="ae129-103">Runtime element</span></span>

<span data-ttu-id="ae129-104">Elemento filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="ae129-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="ae129-105">Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="ae129-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="ae129-106">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ae129-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="ae129-107">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ae129-107">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ae129-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ae129-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ae129-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="ae129-109">Contained in</span></span>

- [<span data-ttu-id="ae129-110">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="ae129-110">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="ae129-111">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae129-111">Attributes</span></span>

|  <span data-ttu-id="ae129-112">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae129-112">Attribute</span></span>  |  <span data-ttu-id="ae129-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ae129-113">Required</span></span>  |  <span data-ttu-id="ae129-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae129-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ae129-115">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="ae129-115">**lifetime="long"**</span></span>  |  <span data-ttu-id="ae129-116">Sim</span><span class="sxs-lookup"><span data-stu-id="ae129-116">Yes</span></span>  | <span data-ttu-id="ae129-117">Deve ser sempre `long` se você quiser usar um tempo de execução compartilhado para o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="ae129-117">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="ae129-118">**resid**</span><span class="sxs-lookup"><span data-stu-id="ae129-118">**resid**</span></span>  |  <span data-ttu-id="ae129-119">Sim</span><span class="sxs-lookup"><span data-stu-id="ae129-119">Yes</span></span>  | <span data-ttu-id="ae129-120">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ae129-120">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="ae129-121">O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="ae129-121">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ae129-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="ae129-122">See also</span></span>

- [<span data-ttu-id="ae129-123">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="ae129-123">Runtimes</span></span>](runtimes.md)
