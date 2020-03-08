---
title: Tempo de execução no arquivo de manifesto (versão prévia)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561825"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="4424a-102">Elemento Runtime (visualização)</span><span class="sxs-lookup"><span data-stu-id="4424a-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="4424a-103">Elemento filho do [`<Runtimes>`](runtimes.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="4424a-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="4424a-104">Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="4424a-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="4424a-105">Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="4424a-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="4424a-106">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4424a-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4424a-107">O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="4424a-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="4424a-108">Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="4424a-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="4424a-109">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4424a-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="4424a-110">Contido em</span><span class="sxs-lookup"><span data-stu-id="4424a-110">Contained in</span></span>

- [<span data-ttu-id="4424a-111">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="4424a-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="4424a-112">Atributos</span><span class="sxs-lookup"><span data-stu-id="4424a-112">Attributes</span></span>

|  <span data-ttu-id="4424a-113">Atributo</span><span class="sxs-lookup"><span data-stu-id="4424a-113">Attribute</span></span>  |  <span data-ttu-id="4424a-114">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4424a-114">Required</span></span>  |  <span data-ttu-id="4424a-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="4424a-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4424a-116">**Lifetime = "Long"**</span><span class="sxs-lookup"><span data-stu-id="4424a-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="4424a-117">Sim</span><span class="sxs-lookup"><span data-stu-id="4424a-117">Yes</span></span>  | <span data-ttu-id="4424a-118">Deve ser `long` sempre se você quiser usar um tempo de execução compartilhado para o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="4424a-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="4424a-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="4424a-119">**resid**</span></span>  |  <span data-ttu-id="4424a-120">Sim</span><span class="sxs-lookup"><span data-stu-id="4424a-120">Yes</span></span>  | <span data-ttu-id="4424a-121">Especifica o local da URL da página HTML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4424a-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="4424a-122">O `resid` deve corresponder a `id` um atributo de `Url` um elemento no `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="4424a-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="4424a-123">Também consulte</span><span class="sxs-lookup"><span data-stu-id="4424a-123">See also</span></span>

- [<span data-ttu-id="4424a-124">Tempos de execução</span><span class="sxs-lookup"><span data-stu-id="4424a-124">Runtimes</span></span>](runtimes.md)
