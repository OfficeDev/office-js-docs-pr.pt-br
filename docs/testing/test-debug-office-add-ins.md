---
title: Testar e depurar suplementos do Office
description: ''
ms.date: 11/24/2017
localization_priority: Priority
ms.openlocfilehash: 7ffa281807ca1541f8ebcc5f722c1043db115509
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388833"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="4ac6d-102">Testar e depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4ac6d-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="4ac6d-103">Esta seção contém orientações sobre testes, depuração de bugs e solução de problemas em suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="4ac6d-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="4ac6d-104">Fazer sideload de suplemento para teste</span><span class="sxs-lookup"><span data-stu-id="4ac6d-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="4ac6d-105">É possível fazer o sideload para instalar um suplemento do Office para teste sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="4ac6d-105">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="4ac6d-106">O procedimento para realizar o sideload de um suplemento varia de acordo com a plataforma e, em alguns casos, também com o produto.</span><span class="sxs-lookup"><span data-stu-id="4ac6d-106">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="4ac6d-107">Os artigos a seguir descrevem como fazer o sideload de suplementos do Office em uma plataforma específica ou em um produto específico:</span><span class="sxs-lookup"><span data-stu-id="4ac6d-107">The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="4ac6d-108">Fazer sideload de Suplementos do Office no Windows</span><span class="sxs-lookup"><span data-stu-id="4ac6d-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="4ac6d-109">Fazer sideload de Suplementos do Office no Office Online</span><span class="sxs-lookup"><span data-stu-id="4ac6d-109">Sideload Office Add-ins in Office Online</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="4ac6d-110">Fazer sideload de Suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="4ac6d-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="4ac6d-111">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="4ac6d-111">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="4ac6d-112">Depurar um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="4ac6d-112">Debug an Office Add-in</span></span>

<span data-ttu-id="4ac6d-113">O procedimento para depurar um suplemento do Office também varia de acordo com a plataforma.</span><span class="sxs-lookup"><span data-stu-id="4ac6d-113">The procedure for debugging an Office Add-in varies by platform as well.</span></span> <span data-ttu-id="4ac6d-114">Cada um dos artigos a seguir descreve como depurar suplementos do Office em uma plataforma específica:</span><span class="sxs-lookup"><span data-stu-id="4ac6d-114">Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="4ac6d-115">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ac6d-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="4ac6d-116">Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10</span><span class="sxs-lookup"><span data-stu-id="4ac6d-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="4ac6d-117">Depurar suplementos no Office Online</span><span class="sxs-lookup"><span data-stu-id="4ac6d-117">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="4ac6d-118">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="4ac6d-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="4ac6d-119">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="4ac6d-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="4ac6d-120">Confira as informações sobre como validar o arquivo de manifesto que descreve os suplementos do Office e solucionar problemas com o arquivo de manifesto em [Validar e solucionar problemas com seu manifesto](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="4ac6d-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="4ac6d-121">Solucionar problemas de erros de usuário</span><span class="sxs-lookup"><span data-stu-id="4ac6d-121">Troubleshoot user errors</span></span>

<span data-ttu-id="4ac6d-122">Confira informações sobre como solucionar problemas comuns que os usuários podem encontrar em seu suplemento do Office em [Solucionar erros de usuários com os suplementos do Office](testing-and-troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="4ac6d-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
