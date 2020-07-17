---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159364"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="fae27-103">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="fae27-103">Runtime element (preview)</span></span>

<span data-ttu-id="fae27-104">Configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="fae27-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="fae27-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="fae27-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="fae27-106">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="fae27-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="fae27-107">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="fae27-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="fae27-108">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="fae27-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="fae27-109">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="fae27-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="fae27-110">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="fae27-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fae27-111">**Outlook**: a ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="fae27-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="fae27-112">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="fae27-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="fae27-113">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="fae27-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="fae27-114">Contido em</span><span class="sxs-lookup"><span data-stu-id="fae27-114">Contained in</span></span>

- [<span data-ttu-id="fae27-115">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="fae27-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="fae27-116">Atributos</span><span class="sxs-lookup"><span data-stu-id="fae27-116">Attributes</span></span>

|  <span data-ttu-id="fae27-117">Atributo</span><span class="sxs-lookup"><span data-stu-id="fae27-117">Attribute</span></span>  |  <span data-ttu-id="fae27-118">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="fae27-118">Required</span></span>  |  <span data-ttu-id="fae27-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="fae27-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="fae27-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="fae27-120">**resid**</span></span>  |  <span data-ttu-id="fae27-121">Sim</span><span class="sxs-lookup"><span data-stu-id="fae27-121">Yes</span></span>  | <span data-ttu-id="fae27-122">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae27-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="fae27-123">O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="fae27-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="fae27-124">**marca**</span><span class="sxs-lookup"><span data-stu-id="fae27-124">**lifetime**</span></span>  |  <span data-ttu-id="fae27-125">Não</span><span class="sxs-lookup"><span data-stu-id="fae27-125">No</span></span>  | <span data-ttu-id="fae27-126">O valor padrão para `lifetime` é `short` e não precisa ser especificado.</span><span class="sxs-lookup"><span data-stu-id="fae27-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="fae27-127">Os suplementos do Outlook usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="fae27-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="fae27-128">Se você quiser usar um tempo de execução compartilhado em um suplemento do Excel, defina explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="fae27-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="fae27-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="fae27-129">See also</span></span>

- [<span data-ttu-id="fae27-130">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="fae27-130">Runtimes</span></span>](runtimes.md)
