---
title: Tempos de execução no arquivo de manifesto
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: ef00bea317ae479d912b3a02f269ef97045b015d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608094"
---
# <a name="runtimes-element"></a><span data-ttu-id="28cca-103">Elemento de runtimes</span><span class="sxs-lookup"><span data-stu-id="28cca-103">Runtimes element</span></span>

<span data-ttu-id="28cca-104">Especifica o tempo de execução do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="28cca-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="28cca-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="28cca-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="28cca-106">Ao executar no Office no Windows, seu suplemento usa o navegador Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="28cca-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="28cca-107">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="28cca-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="28cca-108">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="28cca-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="28cca-109">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="28cca-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="28cca-110">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="28cca-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="28cca-111">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="28cca-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="28cca-112">**Excel**: o tempo de execução compartilhado atualmente só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="28cca-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="28cca-113">**Outlook**: o recurso de ativação baseado em eventos está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="28cca-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="28cca-114">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="28cca-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="28cca-115">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="28cca-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="28cca-116">Contido em</span><span class="sxs-lookup"><span data-stu-id="28cca-116">Contained in</span></span>

[<span data-ttu-id="28cca-117">Host</span><span class="sxs-lookup"><span data-stu-id="28cca-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="28cca-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="28cca-118">Child elements</span></span>

|  <span data-ttu-id="28cca-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="28cca-119">Element</span></span> |  <span data-ttu-id="28cca-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="28cca-120">Required</span></span>  |  <span data-ttu-id="28cca-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="28cca-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="28cca-122">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="28cca-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="28cca-123">Sim</span><span class="sxs-lookup"><span data-stu-id="28cca-123">Yes</span></span> |  <span data-ttu-id="28cca-124">O tempo de execução do suplemento.</span><span class="sxs-lookup"><span data-stu-id="28cca-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="28cca-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="28cca-125">See also</span></span>

- [<span data-ttu-id="28cca-126">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="28cca-126">Runtime</span></span>](runtime.md)
