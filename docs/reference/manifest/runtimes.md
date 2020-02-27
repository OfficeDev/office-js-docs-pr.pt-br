---
title: Tempos de execução no arquivo de manifesto (versão prévia)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283868"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="66f1d-102">Elemento de runtimes (visualização)</span><span class="sxs-lookup"><span data-stu-id="66f1d-102">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="66f1d-103">Especifica o tempo de execução do suplemento e permite funções personalizadas, botões da faixa de opções e o painel de tarefas para usar o mesmo tempo de execução do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="66f1d-103">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="66f1d-104">Filho do `<Host>` elemento no seu arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="66f1d-104">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="66f1d-105">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="66f1d-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="66f1d-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="66f1d-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="66f1d-107">O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="66f1d-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="66f1d-108">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="66f1d-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="66f1d-109">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="66f1d-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="66f1d-110">Contido em</span><span class="sxs-lookup"><span data-stu-id="66f1d-110">Contained in</span></span> 
[<span data-ttu-id="66f1d-111">Host</span><span class="sxs-lookup"><span data-stu-id="66f1d-111">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="66f1d-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="66f1d-112">Child elements</span></span>

|  <span data-ttu-id="66f1d-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="66f1d-113">Element</span></span> |  <span data-ttu-id="66f1d-114">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="66f1d-114">Required</span></span>  |  <span data-ttu-id="66f1d-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="66f1d-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="66f1d-116">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="66f1d-116">**Runtime**</span></span>     | <span data-ttu-id="66f1d-117">Sim</span><span class="sxs-lookup"><span data-stu-id="66f1d-117">Yes</span></span> |  <span data-ttu-id="66f1d-118">O tempo de execução do suplemento.</span><span class="sxs-lookup"><span data-stu-id="66f1d-118">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="66f1d-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="66f1d-119">See also</span></span>

- [<span data-ttu-id="66f1d-120">Runtime</span><span class="sxs-lookup"><span data-stu-id="66f1d-120">Runtime</span></span>](runtime.md)
