---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 74bb2b432f46d5876601052003e20ff843e13b06
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104823"
---
# <a name="runtimes-element"></a><span data-ttu-id="3a7df-103">Elemento Runtimes</span><span class="sxs-lookup"><span data-stu-id="3a7df-103">Runtimes element</span></span>

<span data-ttu-id="3a7df-104">Especifica o tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="3a7df-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="3a7df-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="3a7df-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="3a7df-106">Ao executar no Office no Windows, seu complemento usa o navegador Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="3a7df-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="3a7df-107">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="3a7df-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="3a7df-108">Para saber mais, confira Configurar seu complemento do Excel para usar um tempo de execução [JavaScript compartilhado.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="3a7df-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="3a7df-109">No Outlook, esse elemento habilita a ativação de um complemento baseado em eventos.</span><span class="sxs-lookup"><span data-stu-id="3a7df-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="3a7df-110">Para saber mais, confira [Configurar seu complemento do Outlook para ativação baseada em eventos.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="3a7df-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="3a7df-111">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="3a7df-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a7df-112">**Outlook**: o recurso de ativação baseada em eventos está atualmente em [visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e só está disponível no Outlook na Web e no Windows.</span><span class="sxs-lookup"><span data-stu-id="3a7df-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and Windows.</span></span> <span data-ttu-id="3a7df-113">Para obter mais informações, [consulte Como visualizar o recurso de ativação baseada em eventos.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="3a7df-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="3a7df-114">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="3a7df-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="3a7df-115">Contido em</span><span class="sxs-lookup"><span data-stu-id="3a7df-115">Contained in</span></span>

[<span data-ttu-id="3a7df-116">Host</span><span class="sxs-lookup"><span data-stu-id="3a7df-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="3a7df-117">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3a7df-117">Child elements</span></span>

|  <span data-ttu-id="3a7df-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="3a7df-118">Element</span></span> |  <span data-ttu-id="3a7df-119">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3a7df-119">Required</span></span>  |  <span data-ttu-id="3a7df-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="3a7df-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="3a7df-121">Runtime</span><span class="sxs-lookup"><span data-stu-id="3a7df-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="3a7df-122">Sim</span><span class="sxs-lookup"><span data-stu-id="3a7df-122">Yes</span></span> |  <span data-ttu-id="3a7df-123">O tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="3a7df-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="3a7df-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="3a7df-124">See also</span></span>

- [<span data-ttu-id="3a7df-125">Runtime</span><span class="sxs-lookup"><span data-stu-id="3a7df-125">Runtime</span></span>](runtime.md)
