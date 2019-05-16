---
title: Visualizadores Web usados por Suplementos do Office
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 6cb0d6e97dd559727b6a1e140d8417e1146e479a
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992123"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="b4112-102">Visualizadores Web usados por Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b4112-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="b4112-103">Visto que os Suplementos do Office são aplicativos da Web, eles precisam de um visualizador de páginas da Web para exibir as páginas de HTML do aplicativo da Web e um mecanismo JavaScript para executar o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b4112-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="b4112-104">Ambos são fornecidos por um navegador instalado no computador do usuário.</span><span class="sxs-lookup"><span data-stu-id="b4112-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="b4112-105">Qual navegador é usado depende do:</span><span class="sxs-lookup"><span data-stu-id="b4112-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="b4112-106">Sistema operacional do computador.</span><span class="sxs-lookup"><span data-stu-id="b4112-106">The computer’s operating system.</span></span>
- <span data-ttu-id="b4112-107">Se o suplemento está em execução no Office Online, no Office 365 ou no Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="b4112-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="b4112-108">A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.</span><span class="sxs-lookup"><span data-stu-id="b4112-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="b4112-109">**SO / Plataforma**</span><span class="sxs-lookup"><span data-stu-id="b4112-109">**OS / Platform**</span></span>|<span data-ttu-id="b4112-110">**Navegador**</span><span class="sxs-lookup"><span data-stu-id="b4112-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b4112-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4112-111">Office Online</span></span>|<span data-ttu-id="b4112-112">O navegador no qual o Office Online está aberto.</span><span class="sxs-lookup"><span data-stu-id="b4112-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="b4112-113">Mac</span><span class="sxs-lookup"><span data-stu-id="b4112-113">Mac</span></span>|<span data-ttu-id="b4112-114">Safari</span><span class="sxs-lookup"><span data-stu-id="b4112-114">Safari</span></span>|
|<span data-ttu-id="b4112-115">iOS</span><span class="sxs-lookup"><span data-stu-id="b4112-115">iOS</span></span>|<span data-ttu-id="b4112-116">Safari</span><span class="sxs-lookup"><span data-stu-id="b4112-116">Safari</span></span>|
|<span data-ttu-id="b4112-117">Android</span><span class="sxs-lookup"><span data-stu-id="b4112-117">Android</span></span>|<span data-ttu-id="b4112-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="b4112-118">Chrome</span></span>|
|<span data-ttu-id="b4112-119">Windows / Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="b4112-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="b4112-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="b4112-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="b4112-121">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="b4112-121">Windows 10 ver.</span></span> <span data-ttu-id="b4112-122">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="b4112-122">< 1903 / Office 365</span></span>|<span data-ttu-id="b4112-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="b4112-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="b4112-124">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="b4112-124">Windows 10 ver.</span></span> <span data-ttu-id="b4112-125">>= 1903 / versão do Office 365 < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="b4112-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="b4112-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="b4112-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="b4112-127">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="b4112-127">Windows 10 ver.</span></span> <span data-ttu-id="b4112-128">>= 1903 / versão do Office 365 >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="b4112-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="b4112-129">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="b4112-129">Microsoft Edge\*</span></span>|

<span data-ttu-id="b4112-130">\*Quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela") lê a marcação `<title>` na página que é aberta no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b4112-130">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="b4112-131">Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="b4112-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b4112-132">O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5.</span><span class="sxs-lookup"><span data-stu-id="b4112-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="b4112-133">Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill.</span><span class="sxs-lookup"><span data-stu-id="b4112-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="b4112-134">Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.</span><span class="sxs-lookup"><span data-stu-id="b4112-134">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="b4112-135">Até que eles estejam disponíveis, você precisará ser um Windows Insider para obter a versão 1903 do Windows ou superior, e ser um Office Insider para obter a versão 16.0.11629 do Office ou superior.</span><span class="sxs-lookup"><span data-stu-id="b4112-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="b4112-136">Para participar do programa Windows Insider:</span><span class="sxs-lookup"><span data-stu-id="b4112-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="b4112-137">Vá até [Windows Insider](https://insider.windows.com) e clique no link para participar do Windows Insider.</span><span class="sxs-lookup"><span data-stu-id="b4112-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="b4112-138">Você será direcionado para uma página com instruções sobre como usar as Configurações do Windows para habilitar as compilações de visualização do Windows.</span><span class="sxs-lookup"><span data-stu-id="b4112-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="b4112-139">Siga as instruções.</span><span class="sxs-lookup"><span data-stu-id="b4112-139">Follow the instructions.</span></span> <span data-ttu-id="b4112-140">Quando for selecionar a velocidade das atualizações, escolha a opção mais rápida.</span><span class="sxs-lookup"><span data-stu-id="b4112-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="b4112-141">Para participar do programa Office Insider:</span><span class="sxs-lookup"><span data-stu-id="b4112-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="b4112-142">Vá até [Introdução ao Programa Office Insider](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="b4112-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="b4112-143">Siga as instruções na página para participar.</span><span class="sxs-lookup"><span data-stu-id="b4112-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="b4112-144">Quando solicitado a especificar um canal, selecione Insider.</span><span class="sxs-lookup"><span data-stu-id="b4112-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="b4112-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="b4112-145">See also</span></span>

- [<span data-ttu-id="b4112-146">Requisitos para a Execução de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b4112-146">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
