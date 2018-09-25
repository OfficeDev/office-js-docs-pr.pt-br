---
title: Visão geral da plataforma de suplementos do Office | Documentos da Microsoft
description: Use tecnologias da Web familiares, como HTML, CSS e JavaScript para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook.
ms.date: 01/23/2018
ms.openlocfilehash: 4249e675f4f6a0371e13c5933b3277e7e38841d0
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004975"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="c5bf4-103">Visão geral da plataforma de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="c5bf4-p101">Você pode usar a plataforma de suplementos do Office para criar soluções que estendem os aplicativos do Office e interagem com conteúdo nos documentos do Office. Com os suplementos do Office, você pode usar tecnologias web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook. Sua solução pode ser executada no Office através de várias plataformas, incluindo Office para Windows, Office Online, Office para Mac e Office para iPad.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>

<span data-ttu-id="c5bf4-p102">Os suplementos do Office podem fazer quase tudo que uma página da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="c5bf4-p103">**Adicionar novas funcionalidades para os clientes do Office** – trazer dados externos para o Office, automatizar documentos do Office, expor a funcionalidade de terceiros em clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar aos dados que orientam a produtividade.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span> 
    
-  <span data-ttu-id="c5bf4-111">**Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** ‒ Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-111">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span> 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a><span data-ttu-id="c5bf4-112">Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO?</span><span class="sxs-lookup"><span data-stu-id="c5bf4-112">How are Office Add-ins different than COM and VSTO add-ins?</span></span> 

<span data-ttu-id="c5bf4-p104">Os suplementos de COM ou VSTO são soluções de integração anteriores do Office que são executadas apenas no Office para Windows. Ao contrário de suplementos de COM, os suplementos do Office não envolvem código executado no dispositivo do usuário ou no cliente do Office. Para um suplemento Office, o aplicativo do host, por exemplo, o Excel, lê o manifesto do suplemento e conecta os comandos do menu e os botões da faixa de opções personalizada do suplemento à interface de usuário. Quando necessário, ele carrega o código de HTML e o JavaScript, que são executados no contexto de um navegador em uma área restrita.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span> 

<span data-ttu-id="c5bf4-117">Os suplementos do Office fornecem as seguintes vantagens em relação aos suplementos criados usando o VBA, COM ou VSTO:</span><span class="sxs-lookup"><span data-stu-id="c5bf4-117">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span> 

- <span data-ttu-id="c5bf4-p105">Suporte à plataforma cruzada. Os suplementos do Office podem ser executados no Office para Windows, Mac, iOS e Office Online.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p105">Cross-platform support. Office Add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span> 

- <span data-ttu-id="c5bf4-p106">SSO (logon único). Os suplementos do Office integram-se facilmente com contas do Office 365 dos usuários.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p106">Single sign-on (SSO). Office Add-ins integrate easily with users' Office 365 accounts.</span></span> 

- <span data-ttu-id="c5bf4-p107">Implantação e distribuição centralizada. Os administradores podem implantar suplementos do Office centralmente em uma organização.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p107">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span> 

- <span data-ttu-id="c5bf4-p108">Acesso fácil através da AppSource. Você pode disponibilizar sua solução para um público amplo ao enviá-la para o AppSource.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p108">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span> 

- <span data-ttu-id="c5bf4-p109">Com base na tecnologia de Internet padrão. Você pode usar qualquer biblioteca que gosta para criar suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p109">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span> 

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="c5bf4-128">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-128">Components of an Office Add-in</span></span> 

<span data-ttu-id="c5bf4-p110">Um suplemento do Office inclui dois componentes básicos: um arquivo de manifesto XML e seu próprio aplicativo Web. O manifesto define várias configurações, incluindo como o suplemento é integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p110">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="c5bf4-132">*Figura 1. Manifesto do suplemento (XML) + página da Web (HTML, JS) = um suplemento do Office*</span><span class="sxs-lookup"><span data-stu-id="c5bf4-132">*Figure 1. Add-in manifest (XML) + webpage (HTML, JS) = an Office Add-in*</span></span>

![Manifesto mais página da Web é igual a suplemento do Office](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a><span data-ttu-id="c5bf4-134">Manifesto</span><span class="sxs-lookup"><span data-stu-id="c5bf4-134">Manifest</span></span> 

<span data-ttu-id="c5bf4-135">O manifesto é um arquivo XML que especifica configurações e recursos do suplemento, como os seguintes:</span><span class="sxs-lookup"><span data-stu-id="c5bf4-135">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span> 

- <span data-ttu-id="c5bf4-136">O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-136">The add-in's display name, description, ID, version, and default locale.</span></span> 

- <span data-ttu-id="c5bf4-137">Como o suplemento se integra ao Office.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-137">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="c5bf4-138">O nível de permissão e os requisitos de acesso a dados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-138">The permission level and data access requirements for the add-in.</span></span> 

### <a name="web-app"></a><span data-ttu-id="c5bf4-139">Aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="c5bf4-139">Web app</span></span> 

<span data-ttu-id="c5bf4-p111">O Suplemento do Office mais básico consiste em uma página HTML estática que é exibida dentro de um aplicativo do Office, mas não interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experiência que interaja com os documentos do Office ou permita que o usuário interaja com os recursos online de um aplicativo de host do Office, você pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem dá suporte (como ASP.NET, PHP ou Nó.js). Para interagir com clientes e documentos do Office, você usa as APIs Office.js e JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p111">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span> 

<span data-ttu-id="c5bf4-143">*Figura 2. Componentes de um suplemento Hello World do Office*</span><span class="sxs-lookup"><span data-stu-id="c5bf4-143">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Componentes de um suplemento Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="c5bf4-145">Estender os clientes do Office e interagir com eles</span><span class="sxs-lookup"><span data-stu-id="c5bf4-145">Extending and interacting with Office clients</span></span> 

<span data-ttu-id="c5bf4-146">Os suplementos do Office podem fazer o seguinte em um aplicativo de host do Office:</span><span class="sxs-lookup"><span data-stu-id="c5bf4-146">Office Add-ins can do the following within an Office host application:</span></span> 

-  <span data-ttu-id="c5bf4-147">Estender a funcionalidade (qualquer aplicativo do Office)</span><span class="sxs-lookup"><span data-stu-id="c5bf4-147">Extend functionality (any Office application)</span></span> 

-  <span data-ttu-id="c5bf4-148">Criar novos objetos (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="c5bf4-148">Create new objects (Excel or PowerPoint)</span></span> 
 
### <a name="extend-office-functionality"></a><span data-ttu-id="c5bf4-149">Estender a funcionalidade do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-149">Extend Office functionality</span></span> 

<span data-ttu-id="c5bf4-150">Você pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:</span><span class="sxs-lookup"><span data-stu-id="c5bf4-150">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="c5bf4-151">Botões de faixa de opções e comandos de menu personalizados (coletivamente chamados "comandos de suplemento")</span><span class="sxs-lookup"><span data-stu-id="c5bf4-151">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span> 

-  <span data-ttu-id="c5bf4-152">Painéis de tarefas inseríveis</span><span class="sxs-lookup"><span data-stu-id="c5bf4-152">Insertable task panes</span></span> 

<span data-ttu-id="c5bf4-153">Painéis personalizados de interface do usuário e de tarefa são especificados no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-153">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="c5bf4-154">Botões e comandos de menu personalizados</span><span class="sxs-lookup"><span data-stu-id="c5bf4-154">Custom buttons and menu commands</span></span>  

<span data-ttu-id="c5bf4-p112">Você pode adicionar itens de menu e botões da faixa de opções personalizados à faixa de opções, tanto no Office para Área de Trabalho do Windows quanto no Office Online. Isso facilita aos usuários o acesso ao suplemento diretamente do aplicativo do Office. Botões de comando podem iniciar diferentes ações, como mostrar um painel de tarefas com código HTML personalizado ou executar uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p112">You can add custom ribbon buttons and menu items to the ribbon in Office for Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="c5bf4-158">*Figura 3. Comandos do suplemento na faixa de opções*</span><span class="sxs-lookup"><span data-stu-id="c5bf4-158">*Figure 3. Add-in commands in the ribbon*</span></span>

![Botões e comandos de menu personalizados](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="c5bf4-160">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="c5bf4-160">Task panes</span></span>  

<span data-ttu-id="c5bf4-p113">Você pode usar painéis de tarefas, além dos comandos de suplemento, para permitir que os usuários interajam com sua solução. Os clientes que não dão suporte aos comandos de suplemento (Office 2013 e Office para iPad) executarão seu suplemento como um painel de tarefas. Os usuários iniciam os suplementos do painel de tarefas através do botão **Meus suplementos** na guia **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p113">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span> 

<span data-ttu-id="c5bf4-164">*Figura 4. Painel de tarefas*</span><span class="sxs-lookup"><span data-stu-id="c5bf4-164">*Figure 4. Task pane*</span></span>

![Usar painéis de tarefas além dos comandos de suplemento](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="c5bf4-166">Estender a funcionalidade do Outlook</span><span class="sxs-lookup"><span data-stu-id="c5bf4-166">Extend Outlook functionality</span></span> 

<span data-ttu-id="c5bf4-p114">Os suplementos do Outlook podem estender a faixa de opções do Office e também ser exibidos contextualmente ao lado de um item do Outlook quando você o exibe ou redige. Eles podem trabalhar com uma mensagem de email, uma solicitação de reunião, uma resposta de reunião, um cancelamento de reunião ou um compromisso quando um usuário está visualizando um item recebido, ou respondendo ou criando um novo item.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p114">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="c5bf4-p115">Os suplementos do Outlook podem acessar informação contextual do item, como o endereço ou a ID de rastreamento, e, em seguida, usar estes dados para acessarem informações adicionais sobre o servidor e de serviços da Web para criar experiências do usuário envolventes. Na maioria dos casos, um suplemento do Outlook é executado sem modificação nos vários aplicativos host com suporte, incluindo Outlook, Outlook para Mac, Outlook Web App e Outlook Web App para Dispositivos para fornecer uma experiência perfeita na área de trabalho, na Web e em tablets e dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p115">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span> 

<span data-ttu-id="c5bf4-171">Confira a visão geral dos suplementos do Outlook em [Visão geral dos suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="c5bf4-171">For an overview of Outlook add-ins, see [Outlook add-ins overview](https://docs.microsoft.com/outlook/add-ins/).</span></span> 

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="c5bf4-172">Criar novos objetos nos documentos do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-172">Create new objects in Office documents</span></span> 

<span data-ttu-id="c5bf4-p116">Você pode inserir objetos baseados na web, chamados de suplementos de conteúdo, em documentos do Excel e PowerPoint. Com os suplementos de conteúdo, você pode integrar visualizações de dados avançadas e baseadas na Web, mídia (como um player de vídeo do YouTube ou uma galeria de imagens) e outros tipos de conteúdo externo.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p116">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="c5bf4-175">*Figura 5. Suplemento de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="c5bf4-175">*Figure 5. Content add-in*</span></span>

![Incorporar objetos baseados na web chamados suplementos de conteúdo](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="c5bf4-177">APIs JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-177">Office JavaScript APIs</span></span> 

<span data-ttu-id="c5bf4-p117">As APIs JavaScript para Office contêm objetos e membros para a criação de suplementos e a interação com conteúdo do Office e serviços Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Também existem modelos de objeto específicos de host mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como parágrafos e pastas de trabalho, o que facilita a criação de um suplemento para um host específico.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-p117">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="c5bf4-182">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="c5bf4-182">Next steps</span></span> 

<span data-ttu-id="c5bf4-183">Para saber mais sobre como começar a compilar seu suplemento do Office, experimente nossos [Inícios Rápidos de 5 minutos](https://docs.microsoft.com/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="c5bf4-183">To learn more about how to start building your Office Add-in, try out our [5-minute Quickstarts](https://docs.microsoft.com/office/dev/add-ins/). You can start building add-ins right away using Visual Studio or any other editor.</span></span> <span data-ttu-id="c5bf4-184">Você pode começar a compilar suplementos imediatamente usando o Visual Studio ou qualquer outro editor.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-184">To learn more about how to start building your Office Add-in, try out our 5-minute Quickstarts. You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="c5bf4-185">Para começar a planejar soluções que criem experiências de usuário eficazes e atraentes, familiarize-se com as [diretrizes de design](../design/add-in-design.md) e as [práticas recomendadas](../concepts/add-in-development-best-practices.md) para suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="c5bf4-185">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>    
   
## <a name="see-also"></a><span data-ttu-id="c5bf4-186">Confira também</span><span class="sxs-lookup"><span data-stu-id="c5bf4-186">See also</span></span>

- [<span data-ttu-id="c5bf4-187">Exemplos de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-187">Office Add-in samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="c5bf4-188">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-188">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="c5bf4-189">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c5bf4-189">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)


    
