---
title: Vis?o geral da plataforma de Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f0f20371eee759a449773effaff1ce365e32bf48
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/25/2018
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="8a3d2-102">Vis?o geral da plataforma de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-102">Office Add-ins platform overview</span></span>

<span data-ttu-id="8a3d2-p101">Voc? pode usar a plataforma de suplementos do Office para criar solu??es que estendem os aplicativos do Office e interagem com conte?do nos documentos do Office. Com os suplementos do Office, voc? pode usar tecnologias web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook. Sua solu??o pode ser executada no Office atrav?s de v?rias plataformas, incluindo Office para Windows, Office Online, Office para Mac e Office para iPad.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>

<span data-ttu-id="8a3d2-p102">Os suplementos do Office podem fazer quase tudo que uma p?gina da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="8a3d2-p103">**Adicionar novas funcionalidades para os clientes do Office** ? trazer dados externos para o Office, automatizar documentos do Office, expor a funcionalidade de terceiros em clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar aos dados que orientam a produtividade.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span> 
    
-  <span data-ttu-id="8a3d2-110">**Crie novos objetos avan?ados e interativos que podem ser integrados em documentos do Office** ? Mapas, gr?ficos e visualiza??es interativas integrados que os usu?rios podem adicionar a suas pr?prias planilhas do Excel e apresenta??es do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-110">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span> 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a><span data-ttu-id="8a3d2-111">Quais s?o as diferen?as entre os suplementos do Office e os suplementos de COM e VSTO?</span><span class="sxs-lookup"><span data-stu-id="8a3d2-111">How are Office Add-ins different than COM and VSTO add-ins?</span></span> 

<span data-ttu-id="8a3d2-p104">Os suplementos de COM ou VSTO s?o solu??es de integra??o anteriores do Office que s?o executadas apenas no Office para Windows. Ao contr?rio de suplementos de COM, os suplementos do Office n?o envolvem c?digo executado no dispositivo do usu?rio ou no cliente do Office. Para um suplemento Office, o aplicativo do host, por exemplo, o Excel, l? o manifesto do suplemento e conecta os comandos do menu e os bot?es da faixa de op??es personalizada do suplemento ? interface de usu?rio. Quando necess?rio, ele carrega o c?digo de HTML e o JavaScript, que s?o executados no contexto de um navegador em uma ?rea restrita.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in?s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span> 

<span data-ttu-id="8a3d2-116">Os suplementos do Office fornecem as seguintes vantagens em rela??o aos suplementos criados usando o VBA, COM ou VSTO:</span><span class="sxs-lookup"><span data-stu-id="8a3d2-116">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span> 

- <span data-ttu-id="8a3d2-p105">Suporte ? plataforma cruzada. Os suplementos do Office podem ser executados no Office para Windows, Mac, iOS e Office Online.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p105">Cross-platform support. Office Add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span> 

- <span data-ttu-id="8a3d2-p106">SSO (logon ?nico). Os suplementos do Office integram-se facilmente com contas do Office 365 dos usu?rios.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p106">Single sign-on (SSO). Office Add-ins integrate easily with users' Office 365 accounts.</span></span> 

- <span data-ttu-id="8a3d2-p107">Implanta??o e distribui??o centralizada. Os administradores podem implantar suplementos do Office centralmente em uma organiza??o.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p107">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span> 

- <span data-ttu-id="8a3d2-p108">Acesso f?cil atrav?s da AppSource. Voc? pode disponibilizar sua solu??o para um p?blico amplo ao envi?-la para o AppSource.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p108">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span> 

- <span data-ttu-id="8a3d2-p109">Com base na tecnologia de Internet padr?o. Voc? pode usar qualquer biblioteca que gosta para criar suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p109">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span> 

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="8a3d2-127">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-127">Components of an Office Add-in</span></span> 

<span data-ttu-id="8a3d2-p110">Um suplemento do Office inclui dois componentes b?sicos: um arquivo de manifesto XML e seu pr?prio aplicativo Web. O manifesto define v?rias configura??es, incluindo como o suplemento ? integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou servi?o de hospedagem na Web, como o Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p110">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="8a3d2-131">*Figura 1. Manifesto + p?gina da Web = um Suplemento do Office*</span><span class="sxs-lookup"><span data-stu-id="8a3d2-131">*Figure 1. Manifest + webpage = an Office Add-in*</span></span>

![Manifesto mais p?gina da Web ? igual a suplemento do Office](../images/dk2-agave-overview-01.png)

### <a name="manifest"></a><span data-ttu-id="8a3d2-133">Manifesto</span><span class="sxs-lookup"><span data-stu-id="8a3d2-133">Manifest</span></span> 

<span data-ttu-id="8a3d2-134">O manifesto ? um arquivo XML que especifica configura??es e recursos do suplemento, como os seguintes:</span><span class="sxs-lookup"><span data-stu-id="8a3d2-134">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span> 

- <span data-ttu-id="8a3d2-135">O nome de exibi??o, a descri??o, a ID, a vers?o e a localidade padr?o do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-135">The add-in's display name, description, ID, version, and default locale.</span></span> 

- <span data-ttu-id="8a3d2-136">Como o suplemento se integra ao Office.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-136">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="8a3d2-137">O n?vel de permiss?o e os requisitos de acesso a dados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-137">The permission level and data access requirements for the add-in.</span></span> 

### <a name="web-app"></a><span data-ttu-id="8a3d2-138">Aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="8a3d2-138">Web app</span></span> 

<span data-ttu-id="8a3d2-p111">O Suplemento do Office mais b?sico consiste em uma p?gina HTML est?tica que ? exibida dentro de um aplicativo do Office, mas n?o interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experi?ncia que interaja com os documentos do Office ou permita que o usu?rio interaja com os recursos online de um aplicativo de host do Office, voc? pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem d? suporte (como ASP.NET, PHP ou N?.js). Para interagir com clientes e documentos do Office, voc? usa as APIs Office.js e JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p111">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span> 

<span data-ttu-id="8a3d2-142">*Figura 2. Componentes de um suplemento Hello World do Office*</span><span class="sxs-lookup"><span data-stu-id="8a3d2-142">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Componentes de um suplemento Hello World](../images/dk2-agave-overview-07.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="8a3d2-144">Estender os clientes do Office e interagir com eles</span><span class="sxs-lookup"><span data-stu-id="8a3d2-144">Extending and interacting with Office clients</span></span> 

<span data-ttu-id="8a3d2-145">Os suplementos do Office podem fazer o seguinte em um aplicativo de host do Office:</span><span class="sxs-lookup"><span data-stu-id="8a3d2-145">Office Add-ins can do the following within an Office host application:</span></span> 

-  <span data-ttu-id="8a3d2-146">Estender a funcionalidade (qualquer aplicativo do Office)</span><span class="sxs-lookup"><span data-stu-id="8a3d2-146">Extend functionality (any Office application)</span></span> 

-  <span data-ttu-id="8a3d2-147">Criar novos objetos (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="8a3d2-147">Create new objects (Excel or PowerPoint)</span></span> 
 
### <a name="extend-office-functionality"></a><span data-ttu-id="8a3d2-148">Estender a funcionalidade do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-148">Extend Office functionality</span></span> 

<span data-ttu-id="8a3d2-149">Voc? pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:</span><span class="sxs-lookup"><span data-stu-id="8a3d2-149">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="8a3d2-150">Bot?es de faixa de op??es e comandos de menu personalizados (coletivamente chamados "comandos de suplemento")</span><span class="sxs-lookup"><span data-stu-id="8a3d2-150">Custom ribbon buttons and menu commands (collectively called ?add-in commands?)</span></span> 

-  <span data-ttu-id="8a3d2-151">Pain?is de tarefas inser?veis</span><span class="sxs-lookup"><span data-stu-id="8a3d2-151">Insertable task panes</span></span> 

<span data-ttu-id="8a3d2-152">Pain?is personalizados de interface do usu?rio e de tarefa s?o especificados no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-152">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="8a3d2-153">Bot?es e comandos de menu personalizados</span><span class="sxs-lookup"><span data-stu-id="8a3d2-153">Custom buttons and menu commands</span></span>  

<span data-ttu-id="8a3d2-p112">Voc? pode adicionar itens de menu e bot?es da faixa de op??es personalizados ? faixa de op??es, tanto no Office para ?rea de Trabalho do Windows quanto no Office Online. Isso facilita aos usu?rios o acesso ao suplemento diretamente do aplicativo do Office. Bot?es de comando podem iniciar diferentes a??es, como mostrar um painel de tarefas com c?digo HTML personalizado ou executar uma fun??o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p112">You can add custom ribbon buttons and menu items to the ribbon in Office for Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="8a3d2-157">*Figura 3. Comandos do suplemento em execu??o na ?rea de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="8a3d2-157">*Figure 3. Add-in commands running in Excel Desktop*</span></span>

![Bot?es e comandos de menu personalizados](../images/add-in-commands-overview.png)

#### <a name="task-panes"></a><span data-ttu-id="8a3d2-159">Pain?is de tarefas</span><span class="sxs-lookup"><span data-stu-id="8a3d2-159">Task panes</span></span>  

<span data-ttu-id="8a3d2-p113">Voc? pode usar pain?is de tarefas, al?m dos comandos de suplemento, para permitir que os usu?rios interajam com sua solu??o. Os clientes que n?o d?o suporte aos comandos de suplemento (Office 2013 e Office para iPad) executar?o seu suplemento como um painel de tarefas. Os usu?rios iniciam os suplementos do painel de tarefas atrav?s do bot?o **Meus suplementos** na guia **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p113">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span> 

<span data-ttu-id="8a3d2-163">*Figura 4. Painel de tarefas*</span><span class="sxs-lookup"><span data-stu-id="8a3d2-163">*Figure 4. Task pane*</span></span>

![Painel de tarefas](../images/task-pane-overview.jpg)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="8a3d2-165">Estender a funcionalidade do Outlook</span><span class="sxs-lookup"><span data-stu-id="8a3d2-165">Extend Outlook functionality</span></span> 

<span data-ttu-id="8a3d2-p114">Os suplementos do Outlook podem estender a faixa de op??es do Office e tamb?m ser exibidos contextualmente ao lado de um item do Outlook quando voc? o exibe ou redige. Eles podem trabalhar com uma mensagem de email, uma solicita??o de reuni?o, uma resposta de reuni?o, um cancelamento de reuni?o ou um compromisso quando um usu?rio est? visualizando um item recebido, ou respondendo ou criando um novo item.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p114">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="8a3d2-p115">Os suplementos do Outlook podem acessar informa??o contextual do item, como o endere?o ou a ID de rastreamento, e, em seguida, usar estes dados para acessarem informa??es adicionais sobre o servidor e de servi?os da Web para criar experi?ncias do usu?rio envolventes. Na maioria dos casos, um suplemento do Outlook ? executado sem modifica??o nos v?rios aplicativos host com suporte, incluindo Outlook, Outlook para Mac, Outlook Web App e Outlook Web App para Dispositivos para fornecer uma experi?ncia perfeita na ?rea de trabalho, na Web e em tablets e dispositivos m?veis.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p115">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span> 

<span data-ttu-id="8a3d2-170">Confira a vis?o geral dos suplementos do Outlook em [Vis?o geral dos suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="8a3d2-170">For an overview of Outlook add-ins, see [Outlook add-ins overview](https://docs.microsoft.com/en-us/outlook/add-ins/).</span></span> 

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="8a3d2-171">Criar novos objetos nos documentos do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-171">Create new objects in Office documents</span></span> 

<span data-ttu-id="8a3d2-p116">Voc? pode inserir objetos baseados na web, chamados de suplementos de conte?do, em documentos do Excel e PowerPoint. Com os suplementos de conte?do, voc? pode integrar visualiza??es de dados avan?adas e baseadas na Web, m?dia (como um player de v?deo do YouTube ou uma galeria de imagens) e outros tipos de conte?do externo.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p116">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="8a3d2-174">*Figura 5. Suplemento de conte?do*</span><span class="sxs-lookup"><span data-stu-id="8a3d2-174">*Figure 5. Content add-in*</span></span>

![Suplemento de conte?do](../images/dk2-agave-overview-05.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="8a3d2-176">APIs JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-176">Office JavaScript APIs</span></span> 

<span data-ttu-id="8a3d2-p117">As APIs JavaScript para Office cont?m objetos e membros para a cria??o de suplementos e a intera??o com conte?do do Office e servi?os Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Tamb?m existem modelos de objeto espec?ficos de host mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como par?grafos e pastas de trabalho, o que facilita a cria??o de um suplemento para um host espec?fico.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-p117">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="8a3d2-181">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="8a3d2-181">Next steps</span></span> 

<span data-ttu-id="8a3d2-182">Para saber mais sobre como come?ar a criar o seu Suplemento do Office, experimente o nosso [In?cios R?pidos de 5 minutos](https://docs.microsoft.com/en-us/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="8a3d2-182">To learn more about how to start building your Office Add-in, try out our [5-minute Quickstarts](https://docs.microsoft.com/en-us/office/dev/add-ins/). You can start building add-ins right away using Visual Studio or any other editor.</span></span> <span data-ttu-id="8a3d2-183">Voc? pode come?ar a criar suplementos imediatamente usando o Visual Studio ou qualquer outro editor.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-183">To learn more about how to start building your Office Add-in, try out our 5-minute Quickstarts. You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="8a3d2-184">Para come?ar a planejar solu??es que criem experi?ncias de usu?rio eficazes e atraentes, familiarize-se com as [diretrizes de design](../design/add-in-design.md) e as [pr?ticas recomendadas](../concepts/add-in-development-best-practices.md) para suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8a3d2-184">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>    
   
## <a name="see-also"></a><span data-ttu-id="8a3d2-185">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="8a3d2-185">See also</span></span>

- [<span data-ttu-id="8a3d2-186">Exemplos de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-186">Office Add-in samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="8a3d2-187">No??es b?sicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-187">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="8a3d2-188">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8a3d2-188">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)


    
