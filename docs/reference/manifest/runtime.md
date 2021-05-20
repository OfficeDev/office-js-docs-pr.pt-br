---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, fita, painel de tarefas, funções personalizadas.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555301"
---
# <a name="runtime-element"></a><span data-ttu-id="ecb05-103">Elemento runtime</span><span class="sxs-lookup"><span data-stu-id="ecb05-103">Runtime element</span></span>

<span data-ttu-id="ecb05-104">Configura seu complemento para usar um tempo de execução JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="ecb05-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="ecb05-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="ecb05-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="ecb05-106">**Tipo de complemento:** Painel de tarefas, Correio</span><span class="sxs-lookup"><span data-stu-id="ecb05-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="ecb05-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ecb05-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ecb05-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="ecb05-108">Contained in</span></span>

- [<span data-ttu-id="ecb05-109">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="ecb05-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="ecb05-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ecb05-110">Child elements</span></span>

|  <span data-ttu-id="ecb05-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="ecb05-111">Element</span></span> |  <span data-ttu-id="ecb05-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ecb05-112">Required</span></span>  |  <span data-ttu-id="ecb05-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="ecb05-113">Description</span></span>  |
|:-----|:-----|:-----|
| <span data-ttu-id="ecb05-114">[Substituição](override.md) (visualização)</span><span class="sxs-lookup"><span data-stu-id="ecb05-114">[Override](override.md) (preview)</span></span> | <span data-ttu-id="ecb05-115">Não</span><span class="sxs-lookup"><span data-stu-id="ecb05-115">No</span></span> | <span data-ttu-id="ecb05-116">**Outlook**: Especifica a localização do URL do arquivo JavaScript que Outlook Desktop requer para manipuladores de [pontos de extensão LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="ecb05-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span> <span data-ttu-id="ecb05-117">**Importante**: No momento, você só pode definir um `<Override>` elemento e ele deve ser de tipo `javascript` .</span><span class="sxs-lookup"><span data-stu-id="ecb05-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="ecb05-118">Atributos</span><span class="sxs-lookup"><span data-stu-id="ecb05-118">Attributes</span></span>

|  <span data-ttu-id="ecb05-119">Atributo</span><span class="sxs-lookup"><span data-stu-id="ecb05-119">Attribute</span></span>  |  <span data-ttu-id="ecb05-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ecb05-120">Required</span></span>  |  <span data-ttu-id="ecb05-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="ecb05-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ecb05-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="ecb05-122">**resid**</span></span>  |  <span data-ttu-id="ecb05-123">Sim</span><span class="sxs-lookup"><span data-stu-id="ecb05-123">Yes</span></span>  | <span data-ttu-id="ecb05-124">Especifica a localização da URL da página HTML para o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="ecb05-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="ecb05-125">O `resid` pode não ter mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="ecb05-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="ecb05-126">**vida**</span><span class="sxs-lookup"><span data-stu-id="ecb05-126">**lifetime**</span></span>  |  <span data-ttu-id="ecb05-127">Não</span><span class="sxs-lookup"><span data-stu-id="ecb05-127">No</span></span>  | <span data-ttu-id="ecb05-128">O valor padrão para `lifetime` é e não precisa ser `short` especificado.</span><span class="sxs-lookup"><span data-stu-id="ecb05-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="ecb05-129">Outlook adicionais usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="ecb05-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="ecb05-130">Se você quiser usar um tempo de execução compartilhado em um complemento Excel, defina explicitamente o valor para `long` .</span><span class="sxs-lookup"><span data-stu-id="ecb05-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ecb05-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="ecb05-131">See also</span></span>

- [<span data-ttu-id="ecb05-132">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="ecb05-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="ecb05-133">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="ecb05-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="ecb05-134">Configure seu Outlook complemento para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="ecb05-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
