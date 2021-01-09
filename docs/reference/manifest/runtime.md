---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus diversos componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789181"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="d7f45-103">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="d7f45-103">Runtime element (preview)</span></span>

<span data-ttu-id="d7f45-104">Configura o seu complemento para usar um tempo de execução JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d7f45-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="d7f45-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="d7f45-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="d7f45-106">No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d7f45-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="d7f45-107">Para saber mais, confira Configurar seu complemento do Excel para usar um tempo de execução [JavaScript compartilhado.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="d7f45-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="d7f45-108">No Outlook, esse elemento habilita a ativação de um complemento baseado em eventos.</span><span class="sxs-lookup"><span data-stu-id="d7f45-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="d7f45-109">Para saber mais, confira [Configurar seu complemento do Outlook para ativação baseada em eventos.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="d7f45-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="d7f45-110">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="d7f45-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d7f45-111">**Outlook**: a ativação baseada em eventos está [atualmente em visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e só está disponível no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="d7f45-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="d7f45-112">Para obter mais informações, [consulte Como visualizar o recurso de ativação baseada em eventos.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="d7f45-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="d7f45-113">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d7f45-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="d7f45-114">Contido em</span><span class="sxs-lookup"><span data-stu-id="d7f45-114">Contained in</span></span>

- [<span data-ttu-id="d7f45-115">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="d7f45-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="d7f45-116">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7f45-116">Attributes</span></span>

|  <span data-ttu-id="d7f45-117">Atributo</span><span class="sxs-lookup"><span data-stu-id="d7f45-117">Attribute</span></span>  |  <span data-ttu-id="d7f45-118">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d7f45-118">Required</span></span>  |  <span data-ttu-id="d7f45-119">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7f45-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d7f45-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="d7f45-120">**resid**</span></span>  |  <span data-ttu-id="d7f45-121">Sim</span><span class="sxs-lookup"><span data-stu-id="d7f45-121">Yes</span></span>  | <span data-ttu-id="d7f45-122">Especifica o local da URL da página HTML do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="d7f45-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="d7f45-123">Ele `resid` não pode ter mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="d7f45-123">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="d7f45-124">**tempo de vida**</span><span class="sxs-lookup"><span data-stu-id="d7f45-124">**lifetime**</span></span>  |  <span data-ttu-id="d7f45-125">Não</span><span class="sxs-lookup"><span data-stu-id="d7f45-125">No</span></span>  | <span data-ttu-id="d7f45-126">O valor padrão `lifetime` é e não precisa ser `short` especificado.</span><span class="sxs-lookup"><span data-stu-id="d7f45-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="d7f45-127">Os complementos do Outlook usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="d7f45-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="d7f45-128">Se você quiser usar um tempo de execução compartilhado em um complemento do Excel, de definir explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="d7f45-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="d7f45-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="d7f45-129">See also</span></span>

- [<span data-ttu-id="d7f45-130">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="d7f45-130">Runtimes</span></span>](runtimes.md)
