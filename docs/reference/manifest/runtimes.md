---
title: Tempos de execução no arquivo de manifesto
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 082491befc6b9dbdc474b0e40f9defd90a4ef75f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159357"
---
# <a name="runtimes-element"></a><span data-ttu-id="0bb37-103">Elemento de runtimes</span><span class="sxs-lookup"><span data-stu-id="0bb37-103">Runtimes element</span></span>

<span data-ttu-id="0bb37-104">Especifica o tempo de execução do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="0bb37-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="0bb37-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="0bb37-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="0bb37-106">Ao executar no Office no Windows, seu suplemento usa o navegador Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="0bb37-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="0bb37-107">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="0bb37-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="0bb37-108">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="0bb37-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="0bb37-109">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="0bb37-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="0bb37-110">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="0bb37-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="0bb37-111">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="0bb37-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0bb37-112">**Outlook**: o recurso de ativação baseado em eventos está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="0bb37-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="0bb37-113">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="0bb37-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="0bb37-114">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="0bb37-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="0bb37-115">Contido em</span><span class="sxs-lookup"><span data-stu-id="0bb37-115">Contained in</span></span>

[<span data-ttu-id="0bb37-116">Host</span><span class="sxs-lookup"><span data-stu-id="0bb37-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="0bb37-117">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="0bb37-117">Child elements</span></span>

|  <span data-ttu-id="0bb37-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="0bb37-118">Element</span></span> |  <span data-ttu-id="0bb37-119">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0bb37-119">Required</span></span>  |  <span data-ttu-id="0bb37-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="0bb37-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0bb37-121">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="0bb37-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="0bb37-122">Sim</span><span class="sxs-lookup"><span data-stu-id="0bb37-122">Yes</span></span> |  <span data-ttu-id="0bb37-123">O tempo de execução do suplemento.</span><span class="sxs-lookup"><span data-stu-id="0bb37-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0bb37-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="0bb37-124">See also</span></span>

- [<span data-ttu-id="0bb37-125">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="0bb37-125">Runtime</span></span>](runtime.md)
