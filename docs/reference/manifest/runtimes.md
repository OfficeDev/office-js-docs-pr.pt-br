---
title: Tempos de execução no arquivo de manifesto (versão prévia)
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720418"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="3bb92-103">Elemento de runtimes (visualização)</span><span class="sxs-lookup"><span data-stu-id="3bb92-103">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="3bb92-104">Especifica o tempo de execução do suplemento e permite funções personalizadas, botões da faixa de opções e o painel de tarefas para usar o mesmo tempo de execução do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3bb92-104">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="3bb92-105">Filho do `<Host>` elemento no seu arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="3bb92-105">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="3bb92-106">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="3bb92-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="3bb92-107">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3bb92-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3bb92-108">O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="3bb92-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="3bb92-109">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="3bb92-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="3bb92-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="3bb92-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="3bb92-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="3bb92-111">Contained in</span></span> 
[<span data-ttu-id="3bb92-112">Host</span><span class="sxs-lookup"><span data-stu-id="3bb92-112">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="3bb92-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3bb92-113">Child elements</span></span>

|  <span data-ttu-id="3bb92-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="3bb92-114">Element</span></span> |  <span data-ttu-id="3bb92-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3bb92-115">Required</span></span>  |  <span data-ttu-id="3bb92-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="3bb92-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3bb92-117">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="3bb92-117">**Runtime**</span></span>     | <span data-ttu-id="3bb92-118">Sim</span><span class="sxs-lookup"><span data-stu-id="3bb92-118">Yes</span></span> |  <span data-ttu-id="3bb92-119">O tempo de execução do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3bb92-119">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="3bb92-120">Também confira</span><span class="sxs-lookup"><span data-stu-id="3bb92-120">See also</span></span>

- [<span data-ttu-id="3bb92-121">Runtime</span><span class="sxs-lookup"><span data-stu-id="3bb92-121">Runtime</span></span>](runtime.md)
