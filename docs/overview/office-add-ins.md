---
title: Visão geral da plataforma Suplementos do Office | Microsoft Docs
description: Use tecnologias da Web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com os aplicativos Word, Excel, PowerPoint, OneNote, Project e Outlook.
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6b162a166bda0c988f5fbbaade3b0bef4b650984
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094068"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="de590-103">Visão geral da plataforma de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="de590-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span><span class="sxs-lookup"><span data-stu-id="de590-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="de590-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span><span class="sxs-lookup"><span data-stu-id="de590-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span></span> <span data-ttu-id="de590-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span><span class="sxs-lookup"><span data-stu-id="de590-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Imagem da extensibilidade dos suplementos do Office](../images/addins-overview.png)

<span data-ttu-id="de590-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span><span class="sxs-lookup"><span data-stu-id="de590-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span></span> <span data-ttu-id="de590-109">Use the Office Add-ins platform to:</span><span class="sxs-lookup"><span data-stu-id="de590-109">Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="de590-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span><span class="sxs-lookup"><span data-stu-id="de590-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span></span> <span data-ttu-id="de590-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span><span class="sxs-lookup"><span data-stu-id="de590-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="de590-112">**Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** ‒ Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="de590-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="de590-113">Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO?</span><span class="sxs-lookup"><span data-stu-id="de590-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="de590-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span><span class="sxs-lookup"><span data-stu-id="de590-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span></span> <span data-ttu-id="de590-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span><span class="sxs-lookup"><span data-stu-id="de590-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span></span> <span data-ttu-id="de590-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span><span class="sxs-lookup"><span data-stu-id="de590-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span></span> <span data-ttu-id="de590-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span><span class="sxs-lookup"><span data-stu-id="de590-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Imagem dos motivos para usar os suplementos do Office](../images/why.png)

<span data-ttu-id="de590-119">Os suplementos do Office fornecem as seguintes vantagens em relação aos suplementos criados usando o VBA, COM ou VSTO:</span><span class="sxs-lookup"><span data-stu-id="de590-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="de590-120">Cross-platform support.</span><span class="sxs-lookup"><span data-stu-id="de590-120">Cross-platform support.</span></span> <span data-ttu-id="de590-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span><span class="sxs-lookup"><span data-stu-id="de590-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="de590-122">Centralized deployment and distribution.</span><span class="sxs-lookup"><span data-stu-id="de590-122">Centralized deployment and distribution.</span></span> <span data-ttu-id="de590-123">Admins can deploy Office Add-ins centrally across an organization.</span><span class="sxs-lookup"><span data-stu-id="de590-123">Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="de590-124">Easy access via AppSource.</span><span class="sxs-lookup"><span data-stu-id="de590-124">Easy access via AppSource.</span></span> <span data-ttu-id="de590-125">You can make your solution available to a broad audience by submitting it to AppSource.</span><span class="sxs-lookup"><span data-stu-id="de590-125">You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="de590-126">Based on standard web technology.</span><span class="sxs-lookup"><span data-stu-id="de590-126">Based on standard web technology.</span></span> <span data-ttu-id="de590-127">You can use any library you like to build Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="de590-127">You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="de590-128">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="de590-128">Components of an Office Add-in</span></span>

<span data-ttu-id="de590-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span><span class="sxs-lookup"><span data-stu-id="de590-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span></span> <span data-ttu-id="de590-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span><span class="sxs-lookup"><span data-stu-id="de590-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span></span> <span data-ttu-id="de590-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="de590-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="de590-132">Manifesto</span><span class="sxs-lookup"><span data-stu-id="de590-132">Manifest</span></span>

<span data-ttu-id="de590-133">O manifesto é um arquivo XML que especifica configurações e recursos do suplemento, como os seguintes:</span><span class="sxs-lookup"><span data-stu-id="de590-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="de590-134">O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.</span><span class="sxs-lookup"><span data-stu-id="de590-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="de590-135">Como o suplemento se integra ao Office.</span><span class="sxs-lookup"><span data-stu-id="de590-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="de590-136">O nível de permissão e os requisitos de acesso a dados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="de590-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="de590-137">Aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="de590-137">Web app</span></span>

<span data-ttu-id="de590-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span><span class="sxs-lookup"><span data-stu-id="de590-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span></span> <span data-ttu-id="de590-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span><span class="sxs-lookup"><span data-stu-id="de590-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span></span> <span data-ttu-id="de590-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span><span class="sxs-lookup"><span data-stu-id="de590-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="de590-141">*Figura 2. Componentes de um suplemento Hello World do Office*</span><span class="sxs-lookup"><span data-stu-id="de590-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Componentes de um suplemento Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="de590-143">Estender os clientes do Office e interagir com eles</span><span class="sxs-lookup"><span data-stu-id="de590-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="de590-144">Os suplementos do Office podem fazer o seguinte em um aplicativo de host do Office:</span><span class="sxs-lookup"><span data-stu-id="de590-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="de590-145">Estender a funcionalidade (qualquer aplicativo do Office)</span><span class="sxs-lookup"><span data-stu-id="de590-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="de590-146">Criar novos objetos (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="de590-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="de590-147">Estender a funcionalidade do Office</span><span class="sxs-lookup"><span data-stu-id="de590-147">Extend Office functionality</span></span>

<span data-ttu-id="de590-148">Você pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:</span><span class="sxs-lookup"><span data-stu-id="de590-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="de590-149">Botões de faixa de opções e comandos de menu personalizados (coletivamente chamados "comandos de suplemento")</span><span class="sxs-lookup"><span data-stu-id="de590-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="de590-150">Painéis de tarefas inseríveis</span><span class="sxs-lookup"><span data-stu-id="de590-150">Insertable task panes</span></span>

<span data-ttu-id="de590-151">Painéis personalizados de interface do usuário e de tarefa são especificados no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="de590-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="de590-152">Botões e comandos de menu personalizados</span><span class="sxs-lookup"><span data-stu-id="de590-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="de590-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span><span class="sxs-lookup"><span data-stu-id="de590-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span></span> <span data-ttu-id="de590-154">This makes it easy for users to access your add-in directly from their Office application.</span><span class="sxs-lookup"><span data-stu-id="de590-154">This makes it easy for users to access your add-in directly from their Office application.</span></span> <span data-ttu-id="de590-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span><span class="sxs-lookup"><span data-stu-id="de590-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="de590-156">*Figura 3. Comandos do suplemento na faixa de opções*</span><span class="sxs-lookup"><span data-stu-id="de590-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Botões e comandos de menu personalizados](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="de590-158">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="de590-158">Task panes</span></span>  

<span data-ttu-id="de590-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span><span class="sxs-lookup"><span data-stu-id="de590-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span></span> <span data-ttu-id="de590-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span><span class="sxs-lookup"><span data-stu-id="de590-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span></span> <span data-ttu-id="de590-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span><span class="sxs-lookup"><span data-stu-id="de590-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="de590-162">*Figura 4. Painel de tarefas*</span><span class="sxs-lookup"><span data-stu-id="de590-162">*Figure 4. Task pane*</span></span>

![Usar painéis de tarefas, além dos comandos do suplemento](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="de590-164">Estender a funcionalidade do Outlook</span><span class="sxs-lookup"><span data-stu-id="de590-164">Extend Outlook functionality</span></span>

<span data-ttu-id="de590-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span><span class="sxs-lookup"><span data-stu-id="de590-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span></span> <span data-ttu-id="de590-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span><span class="sxs-lookup"><span data-stu-id="de590-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="de590-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span><span class="sxs-lookup"><span data-stu-id="de590-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span></span> <span data-ttu-id="de590-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span><span class="sxs-lookup"><span data-stu-id="de590-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="de590-169">Confira a visão geral dos suplementos do Outlook em [Visão geral dos suplementos do Outlook](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="de590-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="de590-170">Criar novos objetos nos documentos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-170">Create new objects in Office documents</span></span>

<span data-ttu-id="de590-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span><span class="sxs-lookup"><span data-stu-id="de590-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span></span> <span data-ttu-id="de590-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span><span class="sxs-lookup"><span data-stu-id="de590-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="de590-173">*Figura 5. Suplemento de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="de590-173">*Figure 5. Content add-in*</span></span>

![Inserir objetos baseado na Web chamados suplementos de conteúdo](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="de590-175">APIs JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="de590-175">Office JavaScript APIs</span></span>

<span data-ttu-id="de590-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span><span class="sxs-lookup"><span data-stu-id="de590-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span></span> <span data-ttu-id="de590-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span><span class="sxs-lookup"><span data-stu-id="de590-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span></span> <span data-ttu-id="de590-178">There are also more extensive host-specific object models for Excel and Word.</span><span class="sxs-lookup"><span data-stu-id="de590-178">There are also more extensive host-specific object models for Excel and Word.</span></span> <span data-ttu-id="de590-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span><span class="sxs-lookup"><span data-stu-id="de590-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="de590-180">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="de590-180">Next steps</span></span>

<span data-ttu-id="de590-181">Para obter uma introdução mais detalhada ao desenvolvimento de Suplementos do Office, confira [Criando Suplementos do Offices](../overview/office-add-ins-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="de590-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="de590-182">Confira também</span><span class="sxs-lookup"><span data-stu-id="de590-182">See also</span></span>

- [<span data-ttu-id="de590-183">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="de590-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="de590-184">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="de590-185">Desenvolver Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="de590-186">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="de590-187">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="de590-188">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de590-188">Publish Office Add-ins</span></span>](../publish/publish.md)
