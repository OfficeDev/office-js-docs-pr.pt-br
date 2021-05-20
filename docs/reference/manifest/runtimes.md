---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555294"
---
# <a name="runtimes-element"></a><span data-ttu-id="07381-103">Elemento runtimes</span><span class="sxs-lookup"><span data-stu-id="07381-103">Runtimes element</span></span>

<span data-ttu-id="07381-104">Especifica o tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="07381-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="07381-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="07381-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="07381-106">Ao ser executado em Office em Windows, um complemento que tem um `<Runtimes>` elemento em seu manifesto não é necessariamente executado no mesmo controle de webview que de outra forma seria.</span><span class="sxs-lookup"><span data-stu-id="07381-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="07381-107">Para obter mais informações sobre como as versões de Windows e Office determinar qual controle do webview é normalmente usado, consulte [Navegadores usados por Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso Microsoft Edge com o WebView2 (baseado em Chromium) forem atendidas, o complemento usará esse navegador, quer ele tenha ou não um `<Runtimes>` elemento.</span><span class="sxs-lookup"><span data-stu-id="07381-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="07381-108">No entanto, quando essas condições não são atendidas, um complemento com um `<Runtimes>` elemento sempre usa o Internet Explorer 11, independentemente da Windows ou Microsoft 365 versão.</span><span class="sxs-lookup"><span data-stu-id="07381-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="07381-109">**Tipo de complemento:** Painel de tarefas, Correio</span><span class="sxs-lookup"><span data-stu-id="07381-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="07381-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="07381-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="07381-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="07381-111">Contained in</span></span>

[<span data-ttu-id="07381-112">Host</span><span class="sxs-lookup"><span data-stu-id="07381-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="07381-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="07381-113">Child elements</span></span>

|  <span data-ttu-id="07381-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="07381-114">Element</span></span> |  <span data-ttu-id="07381-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="07381-115">Required</span></span>  |  <span data-ttu-id="07381-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="07381-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="07381-117">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="07381-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="07381-118">Sim</span><span class="sxs-lookup"><span data-stu-id="07381-118">Yes</span></span> |  <span data-ttu-id="07381-119">O tempo de execução para o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="07381-119">The runtime for your add-in.</span></span> <span data-ttu-id="07381-120">**Importante**: No momento, você só pode definir um `<Runtime>` elemento.</span><span class="sxs-lookup"><span data-stu-id="07381-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="07381-121">Confira também</span><span class="sxs-lookup"><span data-stu-id="07381-121">See also</span></span>

- [<span data-ttu-id="07381-122">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="07381-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="07381-123">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="07381-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="07381-124">Configure seu Outlook complemento para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="07381-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
