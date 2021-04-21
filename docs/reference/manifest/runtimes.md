---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917083"
---
# <a name="runtimes-element"></a><span data-ttu-id="d3341-103">Elemento Runtimes</span><span class="sxs-lookup"><span data-stu-id="d3341-103">Runtimes element</span></span>

<span data-ttu-id="d3341-104">Especifica o tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="d3341-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="d3341-105">Filho do [`<Host>`](host.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="d3341-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="d3341-106">Ao executar no Office no Windows, um add-in que tenha um elemento em seu manifesto não necessariamente é executado no mesmo controle `<Runtimes>` de webview como faria.</span><span class="sxs-lookup"><span data-stu-id="d3341-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="d3341-107">Para obter mais informações sobre como as versões do Windows e do Office determinam qual controle webview normalmente é usado, consulte [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso do Microsoft Edge com WebView2 (baseado em Chromium) são atendidas, o complemento usa esse navegador se ele tem ou não um `<Runtimes>` elemento.</span><span class="sxs-lookup"><span data-stu-id="d3341-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="d3341-108">No entanto, quando essas condições não são atendidas, um complemento com um elemento sempre usa o Internet Explorer 11, independentemente da versão do Windows ou `<Runtimes>` do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="d3341-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="d3341-109">**Tipo de complemento:** Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="d3341-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="d3341-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d3341-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="d3341-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="d3341-111">Contained in</span></span>

[<span data-ttu-id="d3341-112">Host</span><span class="sxs-lookup"><span data-stu-id="d3341-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="d3341-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="d3341-113">Child elements</span></span>

|  <span data-ttu-id="d3341-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="d3341-114">Element</span></span> |  <span data-ttu-id="d3341-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d3341-115">Required</span></span>  |  <span data-ttu-id="d3341-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="d3341-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="d3341-117">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="d3341-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="d3341-118">Sim</span><span class="sxs-lookup"><span data-stu-id="d3341-118">Yes</span></span> |  <span data-ttu-id="d3341-119">O tempo de execução do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="d3341-119">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="d3341-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="d3341-120">See also</span></span>

- [<span data-ttu-id="d3341-121">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="d3341-121">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="d3341-122">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="d3341-122">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="d3341-123">Configurar seu complemento do Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="d3341-123">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
