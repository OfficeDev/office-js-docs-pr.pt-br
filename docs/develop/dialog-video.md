---
title: Use a caixa de diálogo do Office para reproduzir um vídeo
description: Saiba como abrir e reproduzir um vídeo na caixa de diálogo do Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 9c65dfb9c0cf1adbc827be25b655e380dc39e2d2
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596526"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="54da0-103">Usar a caixa de diálogo do Office para mostrar um vídeo</span><span class="sxs-lookup"><span data-stu-id="54da0-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="54da0-104">Este artigo explica como reproduzir um vídeo em uma caixa de diálogo do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="54da0-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="54da0-105">Este artigo presume que você esteja familiarizado com as noções básicas de usar a caixa de diálogo do Office, conforme descrito em [usar a API de diálogo do Office em seus suplementos do Office](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="54da0-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="54da0-106">Para reproduzir um vídeo em uma caixa de diálogo com a API de diálogo do Office, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="54da0-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="54da0-107">Criar uma página contendo um iframe e nenhum outro conteúdo.</span><span class="sxs-lookup"><span data-stu-id="54da0-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="54da0-108">A página deve estar no mesmo domínio que a página host.</span><span class="sxs-lookup"><span data-stu-id="54da0-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="54da0-109">Para obter um lembrete sobre o que é uma página de host, consulte [abrir uma caixa de diálogo em uma página de host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="54da0-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="54da0-110">No `src` atributo do iframe, aponte para a URL de um vídeo online.</span><span class="sxs-lookup"><span data-stu-id="54da0-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="54da0-111">O protocolo da URL do vídeo deve ser HTTPS.</span><span class="sxs-lookup"><span data-stu-id="54da0-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="54da0-112">Neste artigo, chamaremos esta página "Video. DialogBox. html".</span><span class="sxs-lookup"><span data-stu-id="54da0-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="54da0-113">Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="54da0-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="54da0-114">Use uma chamada de `displayDialogAsync` na página host para abrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="54da0-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="54da0-115">Se o suplemento precisar saber quando o usuário fecha a caixa de diálogo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006.</span><span class="sxs-lookup"><span data-stu-id="54da0-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="54da0-116">Para obter detalhes, consulte [erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="54da0-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="54da0-117">Para ver um exemplo de reprodução de vídeo em uma caixa de diálogo, confira o [padrão de design do roteiro de vídeo](../design/first-run-experience-patterns.md#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="54da0-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Captura de tela de um vídeo sendo executado em uma caixa de diálogo de suplemento](../images/video-placemats-dialog-open.png)
