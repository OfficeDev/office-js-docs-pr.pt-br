---
title: Tempos de execução no arquivo de manifesto
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 22156a171ca2f423024efb1b3d2a6fdae07dfef6
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278361"
---
# <a name="runtimes-element"></a><span data-ttu-id="ab51a-103">Elemento de runtimes</span><span class="sxs-lookup"><span data-stu-id="ab51a-103">Runtimes element</span></span>

<span data-ttu-id="ab51a-104">Especifica o tempo de execução do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ab51a-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="ab51a-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="ab51a-105">Child of the [`<Host>`](host.md) element.</span></span>

<span data-ttu-id="ab51a-106">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="ab51a-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="ab51a-107">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ab51a-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="ab51a-108">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="ab51a-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="ab51a-109">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="ab51a-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="ab51a-110">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="ab51a-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ab51a-111">**Excel**: o tempo de execução compartilhado está atualmente em versão prévia e disponível apenas no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="ab51a-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="ab51a-112">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="ab51a-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="ab51a-113">**Outlook**: o recurso de ativação baseado em eventos está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ab51a-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="ab51a-114">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="ab51a-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="ab51a-115">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ab51a-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ab51a-116">Contido em</span><span class="sxs-lookup"><span data-stu-id="ab51a-116">Contained in</span></span>

[<span data-ttu-id="ab51a-117">Host</span><span class="sxs-lookup"><span data-stu-id="ab51a-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="ab51a-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ab51a-118">Child elements</span></span>

|  <span data-ttu-id="ab51a-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="ab51a-119">Element</span></span> |  <span data-ttu-id="ab51a-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ab51a-120">Required</span></span>  |  <span data-ttu-id="ab51a-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab51a-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="ab51a-122">Runtime</span><span class="sxs-lookup"><span data-stu-id="ab51a-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="ab51a-123">Sim</span><span class="sxs-lookup"><span data-stu-id="ab51a-123">Yes</span></span> |  <span data-ttu-id="ab51a-124">O tempo de execução do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ab51a-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ab51a-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="ab51a-125">See also</span></span>

- [<span data-ttu-id="ab51a-126">Runtime</span><span class="sxs-lookup"><span data-stu-id="ab51a-126">Runtime</span></span>](runtime.md)
