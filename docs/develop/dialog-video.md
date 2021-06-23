---
title: Use a caixa de diálogo do Office para reproduzir um vídeo
description: Saiba como abrir e reproduzir um vídeo na caixa Office caixa de diálogo
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: bc69827164f2e5a2fed03239566ff814db0397b9
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076066"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="f5ef0-103">Use a caixa Office caixa de diálogo para mostrar um vídeo</span><span class="sxs-lookup"><span data-stu-id="f5ef0-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="f5ef0-104">Este artigo explica como reproduzir um vídeo em uma caixa de Office caixa de diálogo do Add-in.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="f5ef0-105">Este artigo presume que você esteja familiarizado com as noções básicas de uso da caixa de diálogo Office conforme descrito em Usar Office API de diálogo Office em seus Office [Add-ins](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="f5ef0-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="f5ef0-106">Para reproduzir um vídeo em uma caixa de diálogo com a API de Office de diálogo, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="f5ef0-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="f5ef0-107">Crie uma página contendo um iframe e nenhum outro conteúdo.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="f5ef0-108">A página deve estar no mesmo domínio que a página host.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="f5ef0-109">Para um lembrete do que é uma página host, consulte Abrir uma caixa [de diálogo de uma página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="f5ef0-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="f5ef0-110">No atributo `src` do iframe, aponte para a URL de um vídeo online.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="f5ef0-111">O protocolo da URL do vídeo deve ser HTTPS.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="f5ef0-112">Neste artigo, chamaremos essa página de "video.dialogbox.html".</span><span class="sxs-lookup"><span data-stu-id="f5ef0-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="f5ef0-113">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="f5ef0-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="f5ef0-114">Use uma chamada de `displayDialogAsync` na página host para abrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="f5ef0-115">Se o suplemento precisar saber quando o usuário fecha a caixa de diálogo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006.</span><span class="sxs-lookup"><span data-stu-id="f5ef0-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="f5ef0-116">Para obter detalhes, consulte [Erros e eventos na caixa Office caixa de diálogo](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="f5ef0-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="f5ef0-117">Para ver um exemplo de um vídeo que está sendo gravado em uma caixa de diálogo, consulte o padrão de [design de placemat de vídeo](../design/first-run-experience-patterns.md#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="f5ef0-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Captura de tela mostrando um vídeo que está sendo exibido em uma caixa de diálogo do Excel.](../images/video-placemats-dialog-open.png)
