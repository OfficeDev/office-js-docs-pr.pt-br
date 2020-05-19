---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: c2c404bcaad6e24af58f5c0ed8835343abb97e5f
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278410"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="236c0-103">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="236c0-103">Runtime element (preview)</span></span>

<span data-ttu-id="236c0-104">Configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="236c0-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="236c0-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="236c0-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="236c0-106">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="236c0-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="236c0-107">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="236c0-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="236c0-108">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="236c0-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="236c0-109">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="236c0-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="236c0-110">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="236c0-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="236c0-111">**Excel**: o tempo de execução compartilhado está atualmente em versão prévia e disponível apenas no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="236c0-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="236c0-112">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="236c0-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="236c0-113">**Outlook**: a ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="236c0-113">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="236c0-114">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="236c0-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="236c0-115">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="236c0-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="236c0-116">Contido em</span><span class="sxs-lookup"><span data-stu-id="236c0-116">Contained in</span></span>

- [<span data-ttu-id="236c0-117">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="236c0-117">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="236c0-118">Atributos</span><span class="sxs-lookup"><span data-stu-id="236c0-118">Attributes</span></span>

|  <span data-ttu-id="236c0-119">Atributo</span><span class="sxs-lookup"><span data-stu-id="236c0-119">Attribute</span></span>  |  <span data-ttu-id="236c0-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="236c0-120">Required</span></span>  |  <span data-ttu-id="236c0-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="236c0-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="236c0-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="236c0-122">**resid**</span></span>  |  <span data-ttu-id="236c0-123">Sim</span><span class="sxs-lookup"><span data-stu-id="236c0-123">Yes</span></span>  | <span data-ttu-id="236c0-124">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="236c0-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="236c0-125">O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="236c0-125">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="236c0-126">**marca**</span><span class="sxs-lookup"><span data-stu-id="236c0-126">**lifetime**</span></span>  |  <span data-ttu-id="236c0-127">Não</span><span class="sxs-lookup"><span data-stu-id="236c0-127">No</span></span>  | <span data-ttu-id="236c0-128">O valor padrão para `lifetime` é `short` e não precisa ser especificado.</span><span class="sxs-lookup"><span data-stu-id="236c0-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="236c0-129">Os suplementos do Outlook usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="236c0-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="236c0-130">Se você quiser usar um tempo de execução compartilhado em um suplemento do Excel, defina explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="236c0-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="236c0-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="236c0-131">See also</span></span>

- [<span data-ttu-id="236c0-132">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="236c0-132">Runtimes</span></span>](runtimes.md)
