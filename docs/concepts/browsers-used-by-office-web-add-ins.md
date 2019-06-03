---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 05/28/2019
localization_priority: Priority
ms.openlocfilehash: 92218bb012ae9031ebfc429606885a0ec0ea85b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/30/2019
ms.locfileid: "34592126"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="56d60-103">Navegadores usados pelos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="56d60-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="56d60-104">Os suplementos do Office são aplicativos Web exibidos usando iFrames durante a execução do Office Online e no uso de controles de navegador incorporados no Office para clientes de desktops e móveis.</span><span class="sxs-lookup"><span data-stu-id="56d60-104">Office add-ins are web applications that are displayed using iFrames when running in Office Online and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="56d60-105">Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="56d60-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="56d60-106">O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.</span><span class="sxs-lookup"><span data-stu-id="56d60-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="56d60-107">Qual navegador é usado depende do:</span><span class="sxs-lookup"><span data-stu-id="56d60-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="56d60-108">Sistema operacional do computador.</span><span class="sxs-lookup"><span data-stu-id="56d60-108">The computer’s operating system.</span></span>
- <span data-ttu-id="56d60-109">Se o suplemento está em execução no Office Online, no Office 365 ou no Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="56d60-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="56d60-110">A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.</span><span class="sxs-lookup"><span data-stu-id="56d60-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="56d60-111">**SO / Plataforma**</span><span class="sxs-lookup"><span data-stu-id="56d60-111">**OS / Platform**</span></span>|<span data-ttu-id="56d60-112">**Navegador**</span><span class="sxs-lookup"><span data-stu-id="56d60-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="56d60-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="56d60-113">Office Online</span></span>|<span data-ttu-id="56d60-114">O navegador no qual o Office Online está aberto.</span><span class="sxs-lookup"><span data-stu-id="56d60-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="56d60-115">Mac</span><span class="sxs-lookup"><span data-stu-id="56d60-115">Mac</span></span>|<span data-ttu-id="56d60-116">Safari</span><span class="sxs-lookup"><span data-stu-id="56d60-116">Safari</span></span>|
|<span data-ttu-id="56d60-117">iOS</span><span class="sxs-lookup"><span data-stu-id="56d60-117">iOS</span></span>|<span data-ttu-id="56d60-118">Safari</span><span class="sxs-lookup"><span data-stu-id="56d60-118">Safari</span></span>|
|<span data-ttu-id="56d60-119">Android</span><span class="sxs-lookup"><span data-stu-id="56d60-119">Android</span></span>|<span data-ttu-id="56d60-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="56d60-120">Chrome</span></span>|
|<span data-ttu-id="56d60-121">Windows / Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="56d60-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="56d60-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="56d60-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="56d60-123">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="56d60-123">Windows 10 ver.</span></span> <span data-ttu-id="56d60-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="56d60-124">< 1903 / Office 365</span></span>|<span data-ttu-id="56d60-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="56d60-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="56d60-126">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="56d60-126">Windows 10 ver.</span></span> <span data-ttu-id="56d60-127">>= 1903 / versão do Office 365 < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="56d60-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="56d60-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="56d60-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="56d60-129">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="56d60-129">Windows 10 ver.</span></span> <span data-ttu-id="56d60-130">>= 1903 / versão do Office 365 >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="56d60-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="56d60-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="56d60-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="56d60-132">\*Quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela") lê a marcação `<title>` na página que é aberta no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="56d60-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="56d60-133">Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="56d60-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="56d60-134">O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5.</span><span class="sxs-lookup"><span data-stu-id="56d60-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="56d60-135">Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill.</span><span class="sxs-lookup"><span data-stu-id="56d60-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="56d60-136">Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.</span><span class="sxs-lookup"><span data-stu-id="56d60-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="56d60-137">Até que eles estejam disponíveis, você precisará ser um Windows Insider para obter a versão 1903 do Windows ou superior, e ser um Office Insider para obter a versão 16.0.11629 do Office ou superior.</span><span class="sxs-lookup"><span data-stu-id="56d60-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="56d60-138">Para participar do programa Windows Insider:</span><span class="sxs-lookup"><span data-stu-id="56d60-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="56d60-139">Vá até [Windows Insider](https://insider.windows.com) e clique no link para participar do Windows Insider.</span><span class="sxs-lookup"><span data-stu-id="56d60-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="56d60-140">Você será direcionado para uma página com instruções sobre como usar as Configurações do Windows para habilitar as compilações de visualização do Windows.</span><span class="sxs-lookup"><span data-stu-id="56d60-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="56d60-141">Siga as instruções.</span><span class="sxs-lookup"><span data-stu-id="56d60-141">Follow the instructions.</span></span> <span data-ttu-id="56d60-142">Quando for selecionar a velocidade das atualizações, escolha a opção mais rápida.</span><span class="sxs-lookup"><span data-stu-id="56d60-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="56d60-143">Para participar do programa Office Insider:</span><span class="sxs-lookup"><span data-stu-id="56d60-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="56d60-144">Vá até [Introdução ao Programa Office Insider](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="56d60-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="56d60-145">Siga as instruções na página para participar.</span><span class="sxs-lookup"><span data-stu-id="56d60-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="56d60-146">Quando solicitado a especificar um canal, selecione Insider.</span><span class="sxs-lookup"><span data-stu-id="56d60-146">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="56d60-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="56d60-147">See also</span></span>

- [<span data-ttu-id="56d60-148">Requisitos para a Execução de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="56d60-148">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
