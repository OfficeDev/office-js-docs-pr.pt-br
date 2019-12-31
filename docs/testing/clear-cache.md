---
title: Limpar o cache do Office
description: Saiba como limpar o cache do Office em seu computador.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915040"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="4b368-103">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="4b368-103">Clear the Office cache</span></span>

<span data-ttu-id="4b368-104">Você pode remover um suplemento em que foi feito sideload no Windows, Mac ou iOS limpando o cache do Office em seu computador.</span><span class="sxs-lookup"><span data-stu-id="4b368-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="4b368-105">Além disso, se você fizer alterações no manifesto do seu suplemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos de suplemento), você deve limpar o cache do Office e, em seguida, fazer o sideload novamente usando o manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="4b368-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="4b368-106">Isso permitirá que o Office processe o suplemento conforme descrito no manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="4b368-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="4b368-107">Limpar o cache do Office no Windows</span><span class="sxs-lookup"><span data-stu-id="4b368-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="4b368-108">Para limpar o cache do Office no Windows, exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="4b368-108">To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="4b368-109">Limpar o cache do Office no Mac</span><span class="sxs-lookup"><span data-stu-id="4b368-109">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="4b368-110">Limpar o cache do Office no iOS</span><span class="sxs-lookup"><span data-stu-id="4b368-110">Clear the Office cache on iOS</span></span>

<span data-ttu-id="4b368-111">Para limpar o cache do Office no iOS, chame `window.location.reload(true)` a partir do JavaScript no suplemento para forçar um recarregamento.</span><span class="sxs-lookup"><span data-stu-id="4b368-111">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="4b368-112">Uma outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="4b368-112">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="4b368-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="4b368-113">See also</span></span>

- [<span data-ttu-id="4b368-114">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4b368-114">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="4b368-115">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="4b368-115">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="4b368-116">Depurar seu suplemento com o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="4b368-116">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="4b368-117">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="4b368-117">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="4b368-118">Depurar suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4b368-118">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)