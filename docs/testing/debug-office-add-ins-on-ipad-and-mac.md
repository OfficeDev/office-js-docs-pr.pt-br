---
title: Depurar suplementos do Office em um Mac
description: ''
ms.date: 07/29/2019
localization_priority: Priority
ms.openlocfilehash: 10b1181cab23252137df299736341c990978aa1d
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940679"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="a1f59-102">Depurar suplementos do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="a1f59-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="a1f59-p101">Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um Mac.</span><span class="sxs-lookup"><span data-stu-id="a1f59-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on a Mac. Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="a1f59-105">Depuração com Safari Web Inspetor em um Mac</span><span class="sxs-lookup"><span data-stu-id="a1f59-105">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="a1f59-106">Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, o Safari Web Inspector poderá ser usado para depurar um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="a1f59-106">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="a1f59-107">Para poder depurar Suplementos do Office no Mac, você deverá ter o Mac OS High Sierra E o Mac Office Versão: 16.9.1 (build 18012504) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="a1f59-107">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="a1f59-108">Se você não tiver um build do Office Mac, poderá obter um, ingressando no [Programa para desenvolvedores do Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="a1f59-108">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="a1f59-109">Para iniciar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` do aplicativo relevante do Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="a1f59-109">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="a1f59-110">Em seguida, abra o aplicativo do Office e [realize o sideload do seu suplemento](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="a1f59-110">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="a1f59-111">Clique com o botão direito do mouse no suplemento e você verá a opção **Inspecionar Elemento** no menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="a1f59-111">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="a1f59-112">Marque essa opção e ela exibirá o inspetor, onde você poderá definir os pontos de interrupção e depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a1f59-112">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a1f59-113">Se você estiver tentando usar o inspetor e a caixa de diálogo piscar, atualize o Office para a versão mais recente.</span><span class="sxs-lookup"><span data-stu-id="a1f59-113">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="a1f59-114">Se isso não resolver a oscilação, tente a seguinte solução alternativa:</span><span class="sxs-lookup"><span data-stu-id="a1f59-114">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="a1f59-115">Reduza o tamanho da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="a1f59-115">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="a1f59-116">Escolha **Inspecionar Elemento**, que será aberto em uma nova janela.</span><span class="sxs-lookup"><span data-stu-id="a1f59-116">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="a1f59-117">Redimensione a caixa de diálogo para seu tamanho original.</span><span class="sxs-lookup"><span data-stu-id="a1f59-117">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="a1f59-118">Use o inspetor, conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="a1f59-118">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="a1f59-119">Limpar cache do aplicativo do Office em um Mac</span><span class="sxs-lookup"><span data-stu-id="a1f59-119">Clearing the Office application's cache on a Mac or iPad</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
