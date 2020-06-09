---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: e81bd7222585bfa7d5f0f34fe5d9b32e4d45a71e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608101"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="63fb0-103">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="63fb0-103">Runtime element (preview)</span></span>

<span data-ttu-id="63fb0-104">Configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="63fb0-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="63fb0-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="63fb0-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="63fb0-106">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="63fb0-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="63fb0-107">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="63fb0-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="63fb0-108">No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="63fb0-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="63fb0-109">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="63fb0-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="63fb0-110">**Tipo de suplemento:** Painel de tarefas, email</span><span class="sxs-lookup"><span data-stu-id="63fb0-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="63fb0-111">**Excel**: o tempo de execução compartilhado atualmente só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="63fb0-111">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="63fb0-112">**Outlook**: a ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="63fb0-112">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="63fb0-113">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="63fb0-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="63fb0-114">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="63fb0-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="63fb0-115">Contido em</span><span class="sxs-lookup"><span data-stu-id="63fb0-115">Contained in</span></span>

- [<span data-ttu-id="63fb0-116">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="63fb0-116">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="63fb0-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="63fb0-117">Attributes</span></span>

|  <span data-ttu-id="63fb0-118">Atributo</span><span class="sxs-lookup"><span data-stu-id="63fb0-118">Attribute</span></span>  |  <span data-ttu-id="63fb0-119">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="63fb0-119">Required</span></span>  |  <span data-ttu-id="63fb0-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="63fb0-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="63fb0-121">**resid**</span><span class="sxs-lookup"><span data-stu-id="63fb0-121">**resid**</span></span>  |  <span data-ttu-id="63fb0-122">Sim</span><span class="sxs-lookup"><span data-stu-id="63fb0-122">Yes</span></span>  | <span data-ttu-id="63fb0-123">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="63fb0-123">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="63fb0-124">O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="63fb0-124">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="63fb0-125">**marca**</span><span class="sxs-lookup"><span data-stu-id="63fb0-125">**lifetime**</span></span>  |  <span data-ttu-id="63fb0-126">Não</span><span class="sxs-lookup"><span data-stu-id="63fb0-126">No</span></span>  | <span data-ttu-id="63fb0-127">O valor padrão para `lifetime` é `short` e não precisa ser especificado.</span><span class="sxs-lookup"><span data-stu-id="63fb0-127">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="63fb0-128">Os suplementos do Outlook usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="63fb0-128">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="63fb0-129">Se você quiser usar um tempo de execução compartilhado em um suplemento do Excel, defina explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="63fb0-129">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="63fb0-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="63fb0-130">See also</span></span>

- [<span data-ttu-id="63fb0-131">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="63fb0-131">Runtimes</span></span>](runtimes.md)
