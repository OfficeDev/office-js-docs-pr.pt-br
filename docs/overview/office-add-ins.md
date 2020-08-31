---
title: Visão geral da plataforma Suplementos do Office | Microsoft Docs
description: Use tecnologias da Web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com os aplicativos Word, Excel, PowerPoint, OneNote, Project e Outlook.
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 9a504a81bb15e36f937328e2f7cbb674f416842d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292409"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="6976a-103">Visão geral da plataforma de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="6976a-p101">Você pode usar a plataforma de suplementos do Office para criar soluções que estendem os aplicativos do Office e interagem com conteúdo nos documentos do Office. Com os suplementos do Office, você pode usar tecnologias de web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook. Sua solução pode ser executada no Office através de várias plataformas, incluindo Windows, Mac, iPad e em um navegador.</span><span class="sxs-lookup"><span data-stu-id="6976a-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Imagem da extensibilidade dos suplementos do Office](../images/addins-overview.png)

<span data-ttu-id="6976a-p102">Os suplementos do Office podem fazer quase tudo que uma página da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:</span><span class="sxs-lookup"><span data-stu-id="6976a-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="6976a-p103">**Adicionar novas funcionalidades para os clientes do Office** – trazer dados externos para o Office, automatizar documentos do Office, expor a funcionalidade de terceiros em clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar aos dados que orientam a produtividade.</span><span class="sxs-lookup"><span data-stu-id="6976a-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="6976a-112">**Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** ‒ Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6976a-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="6976a-113">Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO?</span><span class="sxs-lookup"><span data-stu-id="6976a-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="6976a-p104">Os suplementos de COM ou VSTO são soluções de integração anteriores do Office que são executadas apenas no Office no Windows. Ao contrário de suplementos de COM, os suplementos do Office não envolvem código executado no dispositivo do usuário ou no cliente do Office. Para um suplemento do Office, o aplicativo do cliente (por exemplo, o Excel), lê o manifesto do suplemento e conecta os comandos do menu e os botões da faixa de opções personalizada do suplemento à interface de usuário. Quando necessário, ele carrega o código de HTML e o JavaScript, que são executados no contexto de um navegador em uma área restrita.</span><span class="sxs-lookup"><span data-stu-id="6976a-p104">COM or VSTO add-ins are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the application (for example, Excel), reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![Imagem dos motivos para usar os suplementos do Office](../images/why.png)

<span data-ttu-id="6976a-119">Os suplementos do Office fornecem as seguintes vantagens em relação aos suplementos criados usando o VBA, COM ou VSTO:</span><span class="sxs-lookup"><span data-stu-id="6976a-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="6976a-p105">Suporte à plataforma cruzada. Os suplementos do Office podem ser executados no Office na Web, Windows, Mac e iPad.</span><span class="sxs-lookup"><span data-stu-id="6976a-p105">Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="6976a-p106">Implantação e distribuição centralizadas. Os administradores podem implantar suplementos do Office centralmente em uma organização.</span><span class="sxs-lookup"><span data-stu-id="6976a-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="6976a-p107">Acesso fácil através da AppSource. Você pode disponibilizar sua solução para um público amplo ao enviá-la para o AppSource.</span><span class="sxs-lookup"><span data-stu-id="6976a-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="6976a-p108">Com base na tecnologia de Internet padrão. Você pode usar qualquer biblioteca que gosta para criar suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="6976a-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="6976a-128">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-128">Components of an Office Add-in</span></span>

<span data-ttu-id="6976a-p109">Um suplemento do Office inclui dois componentes básicos: um arquivo de manifesto XML e seu próprio aplicativo Web. O manifesto define várias configurações, incluindo como o suplemento é integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="6976a-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="6976a-132">Manifesto</span><span class="sxs-lookup"><span data-stu-id="6976a-132">Manifest</span></span>

<span data-ttu-id="6976a-133">O manifesto é um arquivo XML que especifica configurações e recursos do suplemento, como os seguintes:</span><span class="sxs-lookup"><span data-stu-id="6976a-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="6976a-134">O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6976a-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="6976a-135">Como o suplemento se integra ao Office.</span><span class="sxs-lookup"><span data-stu-id="6976a-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="6976a-136">O nível de permissão e os requisitos de acesso a dados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6976a-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="6976a-137">Aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="6976a-137">Web app</span></span>

<span data-ttu-id="6976a-p110">O Suplemento do Office mais básico consiste em uma página HTML estática que é exibida dentro de um aplicativo do Office, mas não interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experiência que interaja com os documentos do Office ou permita que o usuário interaja com os recursos online de um aplicativo cliente do Office, você pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem dá suporte (como ASP.NET, PHP ou Node.js). Para interagir com clientes e documentos do Office, você usa as APIs Office.js e JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6976a-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office client application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="6976a-141">*Figura 2. Componentes de um suplemento Hello World do Office*</span><span class="sxs-lookup"><span data-stu-id="6976a-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Componentes de um suplemento Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="6976a-143">Estender os clientes do Office e interagir com eles</span><span class="sxs-lookup"><span data-stu-id="6976a-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="6976a-144">Os suplementos do Office podem fazer o seguinte em um aplicativo cliente do Office:</span><span class="sxs-lookup"><span data-stu-id="6976a-144">Office Add-ins can do the following within an Office client application:</span></span>

-  <span data-ttu-id="6976a-145">Estender a funcionalidade (qualquer aplicativo do Office)</span><span class="sxs-lookup"><span data-stu-id="6976a-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="6976a-146">Criar novos objetos (Excel ou PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="6976a-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="6976a-147">Estender a funcionalidade do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-147">Extend Office functionality</span></span>

<span data-ttu-id="6976a-148">Você pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:</span><span class="sxs-lookup"><span data-stu-id="6976a-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="6976a-149">Botões de faixa de opções e comandos de menu personalizados (coletivamente chamados "comandos de suplemento")</span><span class="sxs-lookup"><span data-stu-id="6976a-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="6976a-150">Painéis de tarefas inseríveis</span><span class="sxs-lookup"><span data-stu-id="6976a-150">Insertable task panes</span></span>

<span data-ttu-id="6976a-151">Painéis personalizados de interface do usuário e de tarefa são especificados no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6976a-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="6976a-152">Botões e comandos de menu personalizados</span><span class="sxs-lookup"><span data-stu-id="6976a-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="6976a-p111">Você pode adicionar itens de menu e botões da faixa de opções personalizados à faixa de opções, tanto no Office para Área de Trabalho do Windows quanto no Office Online. Isso facilita o acesso dos usuários ao suplemento diretamente do aplicativo do Office. Botões de comando podem iniciar diferentes ações, como mostrar um painel de tarefas com código HTML personalizado ou executar uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6976a-p111">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="6976a-156">*Figura 3. Comandos do suplemento na faixa de opções*</span><span class="sxs-lookup"><span data-stu-id="6976a-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Botões e comandos de menu personalizados](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="6976a-158">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="6976a-158">Task panes</span></span>  

<span data-ttu-id="6976a-p112">Você pode usar painéis de tarefas, além dos comandos de suplemento, para permitir que os usuários interajam com sua solução. Os clientes que não dão suporte aos comandos de suplemento (Office 2013 e Office para iPad) executarão seu suplemento como um painel de tarefas. Os usuários iniciam os suplementos do painel de tarefas através do botão **Meus suplementos** na guia **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="6976a-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="6976a-162">*Figura 4. Painel de tarefas*</span><span class="sxs-lookup"><span data-stu-id="6976a-162">*Figure 4. Task pane*</span></span>

![Usar painéis de tarefas, além dos comandos do suplemento](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="6976a-164">Estender a funcionalidade do Outlook</span><span class="sxs-lookup"><span data-stu-id="6976a-164">Extend Outlook functionality</span></span>

<span data-ttu-id="6976a-p113">Os suplementos do Outlook podem estender a faixa de opções do aplicativo do Office e também serem exibidos contextualmente ao lado de um item do Outlook quando você o exibe ou redige. Eles podem funcionar com uma mensagem de email, uma solicitação de reunião, uma resposta de reunião, um cancelamento de reunião ou um compromisso quando um usuário estiver visualizando um item recebido, respondendo ou criando um novo item.</span><span class="sxs-lookup"><span data-stu-id="6976a-p113">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="6976a-p114">Os suplementos do Outlook podem acessar informações contextuais do item, como o endereço ou a ID de rastreamento, e usar esses dados para acessar informações adicionais no servidor e de serviços Web para criar experiências de usuário envolventes. Na maioria dos casos, um suplemento do Outlook é executado sem modificação no aplicativo cliente do Outlook para fornecer uma experiência perfeita na área de trabalho, na Web e em dispositivos móveis e tablet.</span><span class="sxs-lookup"><span data-stu-id="6976a-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="6976a-169">Confira a visão geral dos suplementos do Outlook em [Visão geral dos suplementos do Outlook](../outlook/outlook-add-ins-overview.md).</span><span class="sxs-lookup"><span data-stu-id="6976a-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="6976a-170">Criar novos objetos nos documentos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-170">Create new objects in Office documents</span></span>

<span data-ttu-id="6976a-p115">Você pode inserir objetos baseados na web, chamados de suplementos de conteúdo, em documentos do Excel e PowerPoint. Com os suplementos de conteúdo, você pode integrar visualizações de dados avançadas e baseadas na Web, mídia (como um player de vídeo do YouTube ou uma galeria de imagens) e outros tipos de conteúdo externo.</span><span class="sxs-lookup"><span data-stu-id="6976a-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="6976a-173">*Figura 5. Suplemento de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="6976a-173">*Figure 5. Content add-in*</span></span>

![Inserir objetos baseado na Web chamados suplementos de conteúdo](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="6976a-175">APIs JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="6976a-175">Office JavaScript APIs</span></span>

<span data-ttu-id="6976a-p116">As APIs JavaScript para Office contêm objetos e membros para a criação de suplementos e a interação com conteúdo do Office e serviços Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Também existem modelos de objeto específicos de aplicativo mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como parágrafos e pastas de trabalho, o que facilita a criação de um suplemento para um aplicativo específico.</span><span class="sxs-lookup"><span data-stu-id="6976a-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive application-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific application.</span></span>

## <a name="next-steps"></a><span data-ttu-id="6976a-180">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="6976a-180">Next steps</span></span>

<span data-ttu-id="6976a-181">Para obter uma introdução mais detalhada ao desenvolvimento de Suplementos do Office, confira [Criando Suplementos do Offices](../overview/office-add-ins-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="6976a-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6976a-182">Confira também</span><span class="sxs-lookup"><span data-stu-id="6976a-182">See also</span></span>

- [<span data-ttu-id="6976a-183">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="6976a-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="6976a-184">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="6976a-185">Desenvolver Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="6976a-186">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="6976a-187">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="6976a-188">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6976a-188">Publish Office Add-ins</span></span>](../publish/publish.md)
