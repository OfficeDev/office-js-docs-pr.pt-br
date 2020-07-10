---
title: Depurar suplementos do Office em um Mac
description: Saiba como usar um Mac para depurar suplementos do Office
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 12785a195c336e0de8c619379a3839bd15079b2c
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094124"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="4abc8-103">Depurar suplementos do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="4abc8-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="4abc8-104">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML.</span><span class="sxs-lookup"><span data-stu-id="4abc8-104">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML.</span></span> <span data-ttu-id="4abc8-105">This article describes how to debug add-ins running on a Mac.</span><span class="sxs-lookup"><span data-stu-id="4abc8-105">This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="4abc8-106">Depuração com Safari Web Inspetor em um Mac</span><span class="sxs-lookup"><span data-stu-id="4abc8-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="4abc8-107">Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, o Safari Web Inspector poderá ser usado para depurar um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="4abc8-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="4abc8-108">Para poder depurar Suplementos do Office no Mac, você deverá ter o Mac OS High Sierra E o Mac Office Versão: 16.9.1 (build 18012504) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="4abc8-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="4abc8-109">Se você não tiver uma compilação Mac do Office, poderá obter uma participando do [programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="4abc8-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="4abc8-110">Para iniciar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` do aplicativo relevante do Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="4abc8-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="4abc8-111">Em seguida, abra o aplicativo do Office e [realize o sideload do seu suplemento](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="4abc8-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="4abc8-112">Clique com o botão direito do mouse no suplemento e você verá a opção **Inspecionar Elemento** no menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="4abc8-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="4abc8-113">Marque essa opção e ela exibirá o inspetor, onde você poderá definir os pontos de interrupção e depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="4abc8-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4abc8-114">Se você estiver tentando usar o inspetor e a caixa de diálogo piscar, atualize o Office para a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="4abc8-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="4abc8-115">Se isso não resolver a oscilação, tente a seguinte solução alternativa:</span><span class="sxs-lookup"><span data-stu-id="4abc8-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="4abc8-116">Reduza o tamanho da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="4abc8-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="4abc8-117">Escolha **Inspecionar Elemento**, que será aberto em uma nova janela.</span><span class="sxs-lookup"><span data-stu-id="4abc8-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="4abc8-118">Redimensione a caixa de diálogo para seu tamanho original.</span><span class="sxs-lookup"><span data-stu-id="4abc8-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="4abc8-119">Use o inspetor, conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="4abc8-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="4abc8-120">Limpar cache do aplicativo do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="4abc8-120">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
