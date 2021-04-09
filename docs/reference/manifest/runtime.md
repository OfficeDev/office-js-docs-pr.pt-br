---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652241"
---
# <a name="runtime-element"></a><span data-ttu-id="5e56c-103">Elemento Runtime</span><span class="sxs-lookup"><span data-stu-id="5e56c-103">Runtime element</span></span>

<span data-ttu-id="5e56c-104">Configura seu complemento para usar um tempo de execução javaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="5e56c-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="5e56c-105">Filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="5e56c-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="5e56c-106">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="5e56c-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="5e56c-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="5e56c-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="5e56c-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="5e56c-108">Contained in</span></span>

- [<span data-ttu-id="5e56c-109">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="5e56c-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="5e56c-110">Atributos</span><span class="sxs-lookup"><span data-stu-id="5e56c-110">Attributes</span></span>

|  <span data-ttu-id="5e56c-111">Atributo</span><span class="sxs-lookup"><span data-stu-id="5e56c-111">Attribute</span></span>  |  <span data-ttu-id="5e56c-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5e56c-112">Required</span></span>  |  <span data-ttu-id="5e56c-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="5e56c-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5e56c-114">**resid**</span><span class="sxs-lookup"><span data-stu-id="5e56c-114">**resid**</span></span>  |  <span data-ttu-id="5e56c-115">Sim</span><span class="sxs-lookup"><span data-stu-id="5e56c-115">Yes</span></span>  | <span data-ttu-id="5e56c-116">Especifica o local da URL da página HTML do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="5e56c-116">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="5e56c-117">O `resid` pode ter não mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="5e56c-117">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="5e56c-118">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="5e56c-118">**lifetime**</span></span>  |  <span data-ttu-id="5e56c-119">Não</span><span class="sxs-lookup"><span data-stu-id="5e56c-119">No</span></span>  | <span data-ttu-id="5e56c-120">O valor padrão `lifetime` para é e não precisa ser `short` especificado.</span><span class="sxs-lookup"><span data-stu-id="5e56c-120">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="5e56c-121">Os complementos do Outlook usam apenas o `short` valor.</span><span class="sxs-lookup"><span data-stu-id="5e56c-121">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="5e56c-122">Se você quiser usar um tempo de execução compartilhado em um complemento do Excel, de definir explicitamente o valor como `long` .</span><span class="sxs-lookup"><span data-stu-id="5e56c-122">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="5e56c-123">Confira também</span><span class="sxs-lookup"><span data-stu-id="5e56c-123">See also</span></span>

- [<span data-ttu-id="5e56c-124">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="5e56c-124">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="5e56c-125">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="5e56c-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="5e56c-126">Configurar seu complemento do Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="5e56c-126">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
