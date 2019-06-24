---
title: Navegadores usados pelos Suplementos do Office
description: Especifica como o sistema operacional e a versão do Office determinam o navegador que é usado pelos suplementos do Office.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 56b74c0e43c8e9709ecd03a8c60a89d3869e44f8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128105"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="97d7a-103">Navegadores usados pelos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="97d7a-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="97d7a-104">Os suplementos do Office são aplicativos Web exibidos usando iFrames durante a execução do Office na Web e no uso de controles de navegador incorporados no Office para clientes desktops e móveis.</span><span class="sxs-lookup"><span data-stu-id="97d7a-104">Office add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="97d7a-105">Os suplementos também precisam de um mecanismo JavaScript para executar o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="97d7a-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="97d7a-106">O navegador incorporado e o mecanismo são fornecidos por um navegador instalado no computador do usuário.</span><span class="sxs-lookup"><span data-stu-id="97d7a-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="97d7a-107">Qual navegador é usado depende do:</span><span class="sxs-lookup"><span data-stu-id="97d7a-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="97d7a-108">Sistema operacional do computador.</span><span class="sxs-lookup"><span data-stu-id="97d7a-108">The computer’s operating system.</span></span>
- <span data-ttu-id="97d7a-109">Se o suplemento está em execução no Office na Web, no Office 365 ou no Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="97d7a-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="97d7a-110">A tabela a seguir mostra qual navegador é usado pelas várias plataformas e sistemas operacionais.</span><span class="sxs-lookup"><span data-stu-id="97d7a-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="97d7a-111">**SO / Plataforma**</span><span class="sxs-lookup"><span data-stu-id="97d7a-111">**OS / Platform**</span></span>|<span data-ttu-id="97d7a-112">**Navegador**</span><span class="sxs-lookup"><span data-stu-id="97d7a-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="97d7a-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="97d7a-113">Office on the web</span></span>|<span data-ttu-id="97d7a-114">O navegador no qual o Office está aberto.</span><span class="sxs-lookup"><span data-stu-id="97d7a-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="97d7a-115">Mac</span><span class="sxs-lookup"><span data-stu-id="97d7a-115">Mac</span></span>|<span data-ttu-id="97d7a-116">Safari</span><span class="sxs-lookup"><span data-stu-id="97d7a-116">Safari</span></span>|
|<span data-ttu-id="97d7a-117">iOS</span><span class="sxs-lookup"><span data-stu-id="97d7a-117">iOS</span></span>|<span data-ttu-id="97d7a-118">Safari</span><span class="sxs-lookup"><span data-stu-id="97d7a-118">Safari</span></span>|
|<span data-ttu-id="97d7a-119">Android</span><span class="sxs-lookup"><span data-stu-id="97d7a-119">Android</span></span>|<span data-ttu-id="97d7a-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="97d7a-120">Chrome</span></span>|
|<span data-ttu-id="97d7a-121">Windows / Office 2013 sem assinatura ou posterior.</span><span class="sxs-lookup"><span data-stu-id="97d7a-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="97d7a-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="97d7a-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="97d7a-123">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="97d7a-123">Windows 10 ver.</span></span> <span data-ttu-id="97d7a-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="97d7a-124">< 1903 / Office 365</span></span>|<span data-ttu-id="97d7a-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="97d7a-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="97d7a-126">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="97d7a-126">Windows 10 ver.</span></span> <span data-ttu-id="97d7a-127">>= 1903 / versão do Office 365 < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="97d7a-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="97d7a-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="97d7a-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="97d7a-129">Versão do Windows 10</span><span class="sxs-lookup"><span data-stu-id="97d7a-129">Windows 10 ver.</span></span> <span data-ttu-id="97d7a-130">>= 1903 / versão do Office 365 >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="97d7a-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="97d7a-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="97d7a-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="97d7a-132">\*Quando o Microsoft Edge está sendo usado, o Windows 10 Narrator (às vezes chamado de "leitor de tela") lê a marcação `<title>` na página que é aberta no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="97d7a-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="97d7a-133">Quando o Internet Explorer 11 está sendo usado, o Narrador lê a barra de título do painel de tarefas, que vem do valor `<DisplayName>` no manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="97d7a-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="97d7a-134">O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5.</span><span class="sxs-lookup"><span data-stu-id="97d7a-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="97d7a-135">Se qualquer um dos usuários de suplemento tiverem plataformas com Internet Explorer 11, para que seja possível usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você precisará fazer o transpile do seu JavaScript para o ES5 ou usar um polyfill.</span><span class="sxs-lookup"><span data-stu-id="97d7a-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="97d7a-136">Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.</span><span class="sxs-lookup"><span data-stu-id="97d7a-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="97d7a-137">Até que eles estejam disponíveis, você precisará ser um Windows Insider para obter a versão 1903 do Windows ou superior, e ser um Office Insider para obter a versão 16.0.11629 do Office ou superior.</span><span class="sxs-lookup"><span data-stu-id="97d7a-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="97d7a-138">Para participar do programa Windows Insider:</span><span class="sxs-lookup"><span data-stu-id="97d7a-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="97d7a-139">Vá até [Windows Insider](https://insider.windows.com) e clique no link para participar do Windows Insider.</span><span class="sxs-lookup"><span data-stu-id="97d7a-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="97d7a-140">Você será direcionado para uma página com instruções sobre como usar as Configurações do Windows para habilitar as compilações de visualização do Windows.</span><span class="sxs-lookup"><span data-stu-id="97d7a-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="97d7a-141">Siga as instruções.</span><span class="sxs-lookup"><span data-stu-id="97d7a-141">Follow the instructions.</span></span> <span data-ttu-id="97d7a-142">Quando for selecionar a velocidade das atualizações, escolha a opção mais rápida.</span><span class="sxs-lookup"><span data-stu-id="97d7a-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="97d7a-143">Para participar do programa Office Insider:</span><span class="sxs-lookup"><span data-stu-id="97d7a-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="97d7a-144">Vá até [Introdução ao Programa Office Insider](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="97d7a-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="97d7a-145">Siga as instruções na página para participar.</span><span class="sxs-lookup"><span data-stu-id="97d7a-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="97d7a-146">Quando solicitado a especificar um canal, selecione Insider.</span><span class="sxs-lookup"><span data-stu-id="97d7a-146">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="97d7a-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="97d7a-147">See also</span></span>

- [<span data-ttu-id="97d7a-148">Requisitos para a Execução de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="97d7a-148">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
