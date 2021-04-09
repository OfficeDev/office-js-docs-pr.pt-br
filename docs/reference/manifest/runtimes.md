---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652228"
---
# <a name="runtimes-element"></a><span data-ttu-id="05230-103">Elemento Runtimes</span><span class="sxs-lookup"><span data-stu-id="05230-103">Runtimes element</span></span>

<span data-ttu-id="05230-104">Especifica o tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="05230-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="05230-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="05230-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="05230-106">Ao executar no Office no Windows, o seu complemento usa o navegador do Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="05230-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="05230-107">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="05230-107">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="05230-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="05230-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="05230-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="05230-109">Contained in</span></span>

[<span data-ttu-id="05230-110">Host</span><span class="sxs-lookup"><span data-stu-id="05230-110">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="05230-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="05230-111">Child elements</span></span>

|  <span data-ttu-id="05230-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="05230-112">Element</span></span> |  <span data-ttu-id="05230-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="05230-113">Required</span></span>  |  <span data-ttu-id="05230-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="05230-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="05230-115">Runtime</span><span class="sxs-lookup"><span data-stu-id="05230-115">Runtime</span></span>](runtime.md) | <span data-ttu-id="05230-116">Sim</span><span class="sxs-lookup"><span data-stu-id="05230-116">Yes</span></span> |  <span data-ttu-id="05230-117">O tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="05230-117">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="05230-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="05230-118">See also</span></span>

- [<span data-ttu-id="05230-119">Runtime</span><span class="sxs-lookup"><span data-stu-id="05230-119">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="05230-120">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="05230-120">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="05230-121">Configurar seu complemento do Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="05230-121">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
