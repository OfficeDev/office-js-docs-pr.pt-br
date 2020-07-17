---
title: Objetos Window que não são compatíveis com suplementos do Office
description: Este artigo especifica alguns dos objetos de tempo de execução da janela que não funcionam em suplementos do Office.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160499"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a><span data-ttu-id="5e0ae-103">Objetos Window que não são compatíveis com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5e0ae-103">Window objects that are unsupported in Office Add-ins</span></span>

<span data-ttu-id="5e0ae-104">Para algumas versões do Windows e do Office, os suplementos são executados em um tempo de execução do Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-104">For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime.</span></span> <span data-ttu-id="5e0ae-105">(Para obter detalhes, consulte [navegadores usados por suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).) Algumas propriedades ou subpropriedades do `window` objeto global não são suportadas no Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-105">(For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11.</span></span> <span data-ttu-id="5e0ae-106">Essas propriedades estão desabilitadas em suplementos para garantir que o suplemento forneça uma experiência consistente para todos os usuários, independentemente do navegador que o suplemento estiver usando.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-106">These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser the add-in is using.</span></span> <span data-ttu-id="5e0ae-107">Isso também ajuda o AngularJS a carregar corretamente.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-107">This also helps AngularJS load properly.</span></span>

<span data-ttu-id="5e0ae-108">Veja a seguir uma lista das propriedades desabilitadas.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-108">The following is a list of the disabled properties.</span></span> <span data-ttu-id="5e0ae-109">A lista é um trabalho em andamento.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-109">The list is a work in progress.</span></span> <span data-ttu-id="5e0ae-110">Se você descobrir `window` Propriedades adicionais que não funcionam em suplementos, use a ferramenta de comentários abaixo para nos dizer.</span><span class="sxs-lookup"><span data-stu-id="5e0ae-110">If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.</span></span>

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a><span data-ttu-id="5e0ae-111">Confira também</span><span class="sxs-lookup"><span data-stu-id="5e0ae-111">See also</span></span>

- [<span data-ttu-id="5e0ae-112">Navegadores usados pelos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5e0ae-112">Browsers used by Office Add-ins</span></span>](../concepts/browsers-used-by-office-web-add-ins.md)