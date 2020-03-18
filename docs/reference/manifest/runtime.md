---
title: Tempo de execução no arquivo de manifesto (versão prévia)
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para sua faixa de opções, painel de tarefas e funções personalizadas.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717926"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="56467-103">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="56467-103">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="56467-104">Elemento filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="56467-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="56467-105">Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="56467-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="56467-106">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="56467-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="56467-107">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="56467-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="56467-108">O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="56467-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="56467-109">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="56467-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="56467-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="56467-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="56467-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="56467-111">Contained in</span></span>

- [<span data-ttu-id="56467-112">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="56467-112">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="56467-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="56467-113">Attributes</span></span>

|  <span data-ttu-id="56467-114">Atributo</span><span class="sxs-lookup"><span data-stu-id="56467-114">Attribute</span></span>  |  <span data-ttu-id="56467-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="56467-115">Required</span></span>  |  <span data-ttu-id="56467-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="56467-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="56467-117">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="56467-117">**lifetime="long"**</span></span>  |  <span data-ttu-id="56467-118">Sim</span><span class="sxs-lookup"><span data-stu-id="56467-118">Yes</span></span>  | <span data-ttu-id="56467-119">Deve ser `long` sempre se você quiser usar um tempo de execução compartilhado para o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="56467-119">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="56467-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="56467-120">**resid**</span></span>  |  <span data-ttu-id="56467-121">Sim</span><span class="sxs-lookup"><span data-stu-id="56467-121">Yes</span></span>  | <span data-ttu-id="56467-122">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="56467-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="56467-123">O `resid` deve corresponder a `id` um atributo de `Url` um elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="56467-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="56467-124">Também confira</span><span class="sxs-lookup"><span data-stu-id="56467-124">See also</span></span>

- [<span data-ttu-id="56467-125">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="56467-125">Runtimes</span></span>](runtimes.md)
