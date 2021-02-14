---
title: Depurar suplementos do Office em um Mac
description: Saiba como usar um Mac para depurar Os Complementos do Office.
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: b2164e3ed672b2911db6841fad24441b67882204
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237942"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="67dcf-103">Depurar suplementos do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="67dcf-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="67dcf-p101">Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um Mac.</span><span class="sxs-lookup"><span data-stu-id="67dcf-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="67dcf-106">Depuração com Safari Web Inspetor em um Mac</span><span class="sxs-lookup"><span data-stu-id="67dcf-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="67dcf-107">Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, o Safari Web Inspector poderá ser usado para depurar um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="67dcf-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="67dcf-108">Para poder depurar Os Complementos do Office no Mac, você deve ter o Mac OS High Sierra E o Mac Office versão 16.9.1 (build 18012504) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="67dcf-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office version 16.9.1 (build 18012504) or later.</span></span> <span data-ttu-id="67dcf-109">Se você não tiver um build do Office Mac, poderá obter um inando no programa de desenvolvedores do [Microsoft 365.](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="67dcf-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="67dcf-110">Para iniciar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` do aplicativo relevante do Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="67dcf-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > <span data-ttu-id="67dcf-111">As builds da Mac App Store do Office não são suportadas pelo `OfficeWebAddinDeveloperExtras` sinalizador.</span><span class="sxs-lookup"><span data-stu-id="67dcf-111">Mac App Store builds of Office do not support the `OfficeWebAddinDeveloperExtras` flag.</span></span>

<span data-ttu-id="67dcf-112">Em seguida, abra o aplicativo do Office e [realize o sideload do seu suplemento](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="67dcf-112">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="67dcf-113">Clique com o botão direito do mouse no suplemento e você verá a opção **Inspecionar Elemento** no menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="67dcf-113">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="67dcf-114">Marque essa opção e ela exibirá o inspetor, onde você poderá definir os pontos de interrupção e depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="67dcf-114">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="67dcf-115">Se você estiver tentando usar o inspetor e a caixa de diálogo piscar, atualize o Office para a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="67dcf-115">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="67dcf-116">Se isso não resolver a oscilação, tente a seguinte solução alternativa:</span><span class="sxs-lookup"><span data-stu-id="67dcf-116">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="67dcf-117">Reduza o tamanho da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="67dcf-117">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="67dcf-118">Escolha **Inspecionar Elemento**, que será aberto em uma nova janela.</span><span class="sxs-lookup"><span data-stu-id="67dcf-118">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="67dcf-119">Redimensione a caixa de diálogo para seu tamanho original.</span><span class="sxs-lookup"><span data-stu-id="67dcf-119">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="67dcf-120">Use o inspetor, conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="67dcf-120">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="67dcf-121">Limpar cache do aplicativo do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="67dcf-121">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
