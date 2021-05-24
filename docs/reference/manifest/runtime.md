---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590908"
---
# <a name="runtime-element"></a><span data-ttu-id="b9717-103">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="b9717-103">Runtime element</span></span>

<span data-ttu-id="b9717-104">Configura seu complemento para usar um tempo de execução javaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="b9717-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="b9717-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="b9717-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="b9717-106">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="b9717-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="b9717-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b9717-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b9717-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="b9717-108">Contained in</span></span>

- [<span data-ttu-id="b9717-109">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="b9717-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="b9717-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="b9717-110">Child elements</span></span>

|  <span data-ttu-id="b9717-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="b9717-111">Element</span></span> |  <span data-ttu-id="b9717-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b9717-112">Required</span></span>  |  <span data-ttu-id="b9717-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9717-113">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="b9717-114">Override</span><span class="sxs-lookup"><span data-stu-id="b9717-114">Override</span></span>](override.md) | <span data-ttu-id="b9717-115">Não</span><span class="sxs-lookup"><span data-stu-id="b9717-115">No</span></span> | <span data-ttu-id="b9717-116">**Outlook**: especifica o local da URL do arquivo JavaScript que Outlook Desktop requer para manipuladores de ponto de extensão [LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="b9717-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span> <span data-ttu-id="b9717-117">**Importante:** no momento, você só pode definir um `<Override>` elemento e ele deve ser do tipo `javascript` .</span><span class="sxs-lookup"><span data-stu-id="b9717-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="b9717-118">Atributos</span><span class="sxs-lookup"><span data-stu-id="b9717-118">Attributes</span></span>

|  <span data-ttu-id="b9717-119">Atributo</span><span class="sxs-lookup"><span data-stu-id="b9717-119">Attribute</span></span>  |  <span data-ttu-id="b9717-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b9717-120">Required</span></span>  |  <span data-ttu-id="b9717-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9717-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b9717-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="b9717-122">**resid**</span></span>  |  <span data-ttu-id="b9717-123">Sim</span><span class="sxs-lookup"><span data-stu-id="b9717-123">Yes</span></span>  | <span data-ttu-id="b9717-124">Especifica o local da URL da página HTML do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="b9717-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="b9717-125">O `resid` pode ter não mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="b9717-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="b9717-126">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="b9717-126">**lifetime**</span></span>  |  <span data-ttu-id="b9717-127">Não</span><span class="sxs-lookup"><span data-stu-id="b9717-127">No</span></span>  | <span data-ttu-id="b9717-128">O valor padrão `lifetime` para é e não precisa ser `short` especificado.</span><span class="sxs-lookup"><span data-stu-id="b9717-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="b9717-129">Outlook os complementos usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="b9717-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="b9717-130">Se você quiser usar um tempo de execução compartilhado em um Excel de Excel, de definir explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="b9717-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b9717-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="b9717-131">See also</span></span>

- [<span data-ttu-id="b9717-132">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="b9717-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="b9717-133">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="b9717-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="b9717-134">Configurar seu Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="b9717-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
