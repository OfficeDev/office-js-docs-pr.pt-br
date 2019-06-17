---
title: Visão geral dos suplementos do Word
description: ''
ms.date: 06/12/2019
localization_priority: Priority
ms.openlocfilehash: bf0b481ef9c982f06ea04aaecb70a76a790846d2
ms.sourcegitcommit: ee5b4935b5ee1db567a13627b2f87471ee8b8165
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/13/2019
ms.locfileid: "34933769"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="e59a4-102">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="e59a4-102">Word add-ins overview</span></span>

<span data-ttu-id="e59a4-p101">Você deseja criar uma solução que estenda a funcionalidade do Word? Por exemplo, uma solução que envolva conjuntos de documentos automatizados? Ou uma solução que vincule essas associações aos dados e os acesse em um documento do Word a partir de outras fontes de dados? É possível usar a plataforma de suplementos do Office, que inclui a API JavaScript do Word e a API JavaScript para Office, para estender os clientes do Word que executam em uma área de trabalho do Windows, em um Mac ou na nuvem.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="e59a4-p102">Os suplementos do Word são uma das várias opções de desenvolvimento disponíveis na [plataforma de suplementos do Office](../overview/office-add-ins.md). Você pode usar comandos de suplemento para estender a interface do usuário do Word e iniciar os painéis de tarefas que executam JavaScript que interage com o conteúdo em um documento do Word. Qualquer código que você pode executar em um navegador, pode ser executado em um suplemento do Word. Suplementos que interagem com conteúdo em um documento do Word criam solicitações para agir em objetos do Word e sincronizar o estado do objeto.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="e59a4-p103">Caso pretenda [publicar](../publish/publish.md) o suplemento no AppSource depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade do suplemento do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="e59a4-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="e59a4-113">A figura a seguir mostra um exemplo de um suplemento do Word que é executado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e59a4-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="e59a4-114">*Figura 1. Suplemento em execução em um painel de tarefas no Word*</span><span class="sxs-lookup"><span data-stu-id="e59a4-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Suplemento em execução em um painel de tarefas no Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="e59a4-p104">O suplemento do Word (1) pode enviar solicitações para o documento do Word (2) e usar o JavaScript para acessar o objeto parágrafo e atualizar, excluir ou mover o parágrafo. Por exemplo, o código a seguir mostra como acrescentar uma nova sentença a esse parágrafo.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="e59a4-p105">É possível usar qualquer tecnologia de servidor Web para hospedar o suplemento do Word, como ASP.NET, NodeJS ou Python. Use a estrutura de cliente de sua preferência (Ember, Backbone, Angular, React) ou use o VanillaJS para desenvolver a solução. É possível usar serviços como o Azure para [autenticar](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) e hospedar seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="e59a4-p106">As APIs JavaScript do Word proporcionam ao seu aplicativo o acesso aos objetos e metadados encontrado em um documento do Word. Você pode usar essas APIs para criar suplementos que têm como objetivo:</span><span class="sxs-lookup"><span data-stu-id="e59a4-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="e59a4-122">Word 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="e59a4-122">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="e59a4-123">Word Online</span><span class="sxs-lookup"><span data-stu-id="e59a4-123">Word Online</span></span>
* <span data-ttu-id="e59a4-124">Word 2016 ou posterior para Mac</span><span class="sxs-lookup"><span data-stu-id="e59a4-124">Word 2016 or later for Mac</span></span>
* <span data-ttu-id="e59a4-125">Word para iPad</span><span class="sxs-lookup"><span data-stu-id="e59a4-125">Word for iPad</span></span>

<span data-ttu-id="e59a4-p107">Redija seu suplemento uma vez e ele será executado em todas as versões do Word em várias plataformas. Para obter detalhes, consulte [Disponibilidade de Suplementos do Office em hosts e plataformas](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="e59a4-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="e59a4-128">APIs JavaScript para Word</span><span class="sxs-lookup"><span data-stu-id="e59a4-128">JavaScript APIs for Word</span></span>

<span data-ttu-id="e59a4-129">Você pode usar dois conjuntos de APIs JavaScript para interagir com metadados e objetos em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="e59a4-129">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="e59a4-130">O primeiro é a [API comuns](../reference/javascript-api-for-office.md), que foi introduzido no Office 2013.</span><span class="sxs-lookup"><span data-stu-id="e59a4-130">The first is the [Common API](../reference/javascript-api-for-office.md), which was introduced in Office 2013.</span></span> <span data-ttu-id="e59a4-131">Muitos dos objetos na API comuns podem ser usados em suplementos hospedados por dois ou mais clientes do Office. </span><span class="sxs-lookup"><span data-stu-id="e59a4-131">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="e59a4-132">Essa API usa retornos de chamadas de maneira ampla.</span><span class="sxs-lookup"><span data-stu-id="e59a4-132">This API uses callbacks extensively.</span></span>

<span data-ttu-id="e59a4-p109">O segundo é a [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md). Este é um modelo de objeto fortemente tipado que você pode usar para criar suplementos do Word que se destinam ao Word 2016 para Mac e Windows. Este modelo de objeto usa promessas e fornece acesso a objetos específicos do Word como [corpo](/javascript/api/word/word.body), [controles de conteúdo](/javascript/api/word/word.contentcontrol), [imagens embutidas](/javascript/api/word/word.inlinepicture) e [parágrafos](/javascript/api/word/word.paragraph). A API JavaScript do Word inclui definições do TypeScript e arquivos vsdoc para que você possa obter dicas de código em seu IDE.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p109">The second is the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="e59a4-p110">Atualmente, todos os clientes do Word oferecem suporte à API JavaScript para Office compartilhada, e a maioria dos clientes oferece suporte à API JavaScript do Word. Para obter detalhes sobre clientes com suporte, consulte a [documentação de referência da API](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span><span class="sxs-lookup"><span data-stu-id="e59a4-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="e59a4-p111">Recomendamos que você comece com a API JavaScript do Word porque o modelo de objeto é mais fácil de usar. Use a API JavaScript do Word se precisar:</span><span class="sxs-lookup"><span data-stu-id="e59a4-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="e59a4-141">Acessar os objetos em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="e59a4-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="e59a4-142">Use a API JavaScript para Office compartilhada quando precisar:</span><span class="sxs-lookup"><span data-stu-id="e59a4-142">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="e59a4-143">Direcionar o Word 2013.</span><span class="sxs-lookup"><span data-stu-id="e59a4-143">Target Word 2013.</span></span>
* <span data-ttu-id="e59a4-144">Executar ações iniciais do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="e59a4-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="e59a4-145">Verificar o conjunto requisitos com suporte.</span><span class="sxs-lookup"><span data-stu-id="e59a4-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="e59a4-146">Acessar metadados, configurações e informações do ambiente para o documento.</span><span class="sxs-lookup"><span data-stu-id="e59a4-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="e59a4-147">Vincular a seções em um documento e capturar eventos.</span><span class="sxs-lookup"><span data-stu-id="e59a4-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="e59a4-148">Usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e59a4-148">Use custom XML parts.</span></span>
* <span data-ttu-id="e59a4-149">Abrir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="e59a4-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e59a4-150">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e59a4-150">Next steps</span></span>

<span data-ttu-id="e59a4-p112">Pronto para criar seu primeiro suplemento do Word? Confira [Criar seu primeiro suplemento do Word](word-add-ins.md). Use o [manifesto de suplemento](../develop/add-in-manifests.md) para descrever onde seu suplemento está hospedado e como ele é exibido, bem como para definir permissões e outras informações.</span><span class="sxs-lookup"><span data-stu-id="e59a4-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](../develop/add-in-manifests.md). Use the add-in manifest to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="e59a4-154">Para saber mais sobre como projetar um suplemento do Word de classe internacional que cria uma ótima experiência para seus usuários, consulte [Diretrizes de design](../design/add-in-design.md) e [Práticas recomendadas](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="e59a4-154">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="e59a4-155">Depois de desenvolver seu suplemento, é possível [publicá-lo](../publish/publish.md) em um compartilhamento de rede, um catálogo de aplicativos ou no AppSource.</span><span class="sxs-lookup"><span data-stu-id="e59a4-155">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="e59a4-156">Confira também</span><span class="sxs-lookup"><span data-stu-id="e59a4-156">See also</span></span>

* [<span data-ttu-id="e59a4-157">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e59a4-157">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="e59a4-158">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="e59a4-158">Word JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
