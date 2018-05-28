---
title: Vis?o geral dos suplementos do Word
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 63605c18f7e1b3eae2c542aef236372819bc2e6f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="ed87b-102">Vis?o geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="ed87b-102">Word add-ins overview</span></span>

<span data-ttu-id="ed87b-p101">Voc? deseja criar uma solu??o que estenda a funcionalidade do Word? Por exemplo, uma solu??o que envolva conjuntos de documentos automatizados? Ou uma solu??o que vincule essas associa??es aos dados e os acesse em um documento do Word a partir de outras fontes de dados? ? poss?vel usar a plataforma de suplementos do Office, que inclui a API JavaScript do Word e a API JavaScript para Office, para estender os clientes do Word que executam em uma ?rea de trabalho do Windows, em um Mac ou na nuvem.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="ed87b-p102">Os suplementos do Word s?o uma das v?rias op??es de desenvolvimento dispon?veis na [plataforma de suplementos do Office](../overview/office-add-ins.md). Voc? pode usar comandos de suplemento para estender a interface do usu?rio do Word e iniciar os pain?is de tarefas que executam JavaScript que interage com o conte?do em um documento do Word. Qualquer c?digo que voc? pode executar em um navegador, pode ser executado em um suplemento do Word. Suplementos que interagem com conte?do em um documento do Word criam solicita??es para agir em objetos do Word e sincronizar o estado do objeto.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="ed87b-p103">Caso pretenda [publicar](../publish/publish.md) o suplemento no AppSource depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade do suplemento do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="ed87b-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="ed87b-113">A figura a seguir mostra um exemplo de um suplemento do Word que ? executado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ed87b-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="ed87b-114">*Figura 1. Suplemento em execu??o em um painel de tarefas no Word*</span><span class="sxs-lookup"><span data-stu-id="ed87b-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Suplemento em execu??o em um painel de tarefas no Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="ed87b-p104">O suplemento do Word (1) pode enviar solicita??es para o documento do Word (2) e usar o JavaScript para acessar o objeto par?grafo e atualizar, excluir ou mover o par?grafo. Por exemplo, o c?digo a seguir mostra como acrescentar uma nova senten?a a esse par?grafo.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

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

<span data-ttu-id="ed87b-p105">? poss?vel usar qualquer tecnologia de servidor Web para hospedar o suplemento do Word, como ASP.NET, NodeJS ou Python. Use a estrutura de cliente de sua prefer?ncia (Ember, Backbone, Angular, React) ou use o VanillaJS para desenvolver a solu??o. ? poss?vel usar servi?os como o Azure para [autenticar](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) e hospedar seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="ed87b-p106">As APIs JavaScript do Word proporcionam ao seu aplicativo o acesso aos objetos e metadados encontrado em um documento do Word. Voc? pode usar essas APIs para criar suplementos que t?m como objetivo:</span><span class="sxs-lookup"><span data-stu-id="ed87b-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="ed87b-122">Word 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="ed87b-122">Word 2013 for Windows</span></span>
* <span data-ttu-id="ed87b-123">Word 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="ed87b-123">Word 2016 for Windows</span></span>
* <span data-ttu-id="ed87b-124">Word Online</span><span class="sxs-lookup"><span data-stu-id="ed87b-124">Word Online</span></span>
* <span data-ttu-id="ed87b-125">Word 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="ed87b-125">Word 2016 for Mac</span></span>
* <span data-ttu-id="ed87b-126">Word para iOS</span><span class="sxs-lookup"><span data-stu-id="ed87b-126">Word for iOS</span></span>

<span data-ttu-id="ed87b-p107">Redija seu suplemento uma vez e ele ser? executado em todas as vers?es do Word em v?rias plataformas. Para obter detalhes, consulte [Disponibilidade de Suplementos do Office em hosts e plataformas](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="ed87b-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="ed87b-129">APIs JavaScript para Word</span><span class="sxs-lookup"><span data-stu-id="ed87b-129">JavaScript APIs for Word</span></span>

<span data-ttu-id="ed87b-p108">Voc? pode usar dois conjuntos de APIs JavaScript para interagir com metadados e objetos em um documento do Word. O primeiro ? o [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word), que foi introduzido no Office 2013. Esta ? uma API compartilhada ? muitos dos objetos podem ser usados em suplementos hospedados por dois ou mais clientes do Office. Essa API usa retornos de chamadas de maneira ampla.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p108">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word), which was introduced in Office 2013. This is a shared API -- many of the objects can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.</span></span> 

<span data-ttu-id="ed87b-p109">O segundo ? a [API JavaScript do Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). Este ? um modelo de objeto fortemente tipado que voc? pode usar para criar suplementos do Word que se destinam ao Word 2016 para Mac e Windows. Este modelo de objeto usa promessas e fornece acesso a objetos espec?ficos do Word como [corpo](https://dev.office.com/reference/add-ins/word/body), [controles de conte?do](https://dev.office.com/reference/add-ins/word/contentcontrol), [imagens embutidas](https://dev.office.com/reference/add-ins/word/inlinepicture) e [par?grafos](https://dev.office.com/reference/add-ins/word/paragraph). A API JavaScript do Word inclui defini??es do TypeScript e arquivos vsdoc para que voc? possa obter dicas de c?digo em seu IDE.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p109">The second is the [Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](https://dev.office.com/reference/add-ins/word/body), [content controls](https://dev.office.com/reference/add-ins/word/contentcontrol), [inline pictures](https://dev.office.com/reference/add-ins/word/inlinepicture), and [paragraphs](https://dev.office.com/reference/add-ins/word/paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="ed87b-p110">Atualmente, todos os clientes do Word oferecem suporte ? API JavaScript para Office compartilhada, e a maioria dos clientes oferece suporte ? API JavaScript do Word. Para obter detalhes sobre clientes com suporte, consulte a [documenta??o de refer?ncia da API](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).</span><span class="sxs-lookup"><span data-stu-id="ed87b-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="ed87b-p111">Recomendamos que voc? comece com a API JavaScript do Word porque o modelo de objeto ? mais f?cil de usar. Use a API JavaScript do Word se precisar:</span><span class="sxs-lookup"><span data-stu-id="ed87b-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="ed87b-142">Acessar os objetos em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="ed87b-142">Access the objects in a Word document.</span></span>

<span data-ttu-id="ed87b-143">Use a API JavaScript para Office compartilhada quando precisar:</span><span class="sxs-lookup"><span data-stu-id="ed87b-143">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="ed87b-144">Direcionar o Word 2013.</span><span class="sxs-lookup"><span data-stu-id="ed87b-144">Target Word 2013.</span></span>
* <span data-ttu-id="ed87b-145">Executar a??es iniciais do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="ed87b-145">Perform initial actions for the application.</span></span>
* <span data-ttu-id="ed87b-146">Verificar o conjunto requisitos com suporte.</span><span class="sxs-lookup"><span data-stu-id="ed87b-146">Check the supported requirement set.</span></span>
* <span data-ttu-id="ed87b-147">Acessar metadados, configura??es e informa??es do ambiente para o documento.</span><span class="sxs-lookup"><span data-stu-id="ed87b-147">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="ed87b-148">Vincular a se??es em um documento e capturar eventos.</span><span class="sxs-lookup"><span data-stu-id="ed87b-148">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="ed87b-149">Usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ed87b-149">Use custom XML parts.</span></span>
* <span data-ttu-id="ed87b-150">Abrir uma caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="ed87b-150">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ed87b-151">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="ed87b-151">Next steps</span></span>

<span data-ttu-id="ed87b-p112">Pronto para criar seu primeiro suplemento do Word? Confira [Compilar seu primeiro suplemento do Word](word-add-ins.md). Tamb?m ? poss?vel tentar nossa [Experi?ncia de introdu??o](http://dev.office.com/getting-started/addins?product=Word) interativa. Use um [manifesto do suplemento](../develop/add-in-manifests.md) para descrever onde seu suplemento est? hospedado e como ele ? exibido, al?m de definir permiss?es e outras informa??es.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](http://dev.office.com/getting-started/addins?product=Word). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="ed87b-156">Para saber mais sobre como projetar um suplemento do Word de classe internacional que cria uma ?tima experi?ncia para seus usu?rios, consulte [Diretrizes de design](../design/add-in-design.md) e [Pr?ticas recomendadas](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="ed87b-156">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="ed87b-157">Depois de desenvolver seu suplemento, ? poss?vel [public?-lo](../publish/publish.md) em um compartilhamento de rede, um cat?logo de aplicativos ou no AppSource.</span><span class="sxs-lookup"><span data-stu-id="ed87b-157">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="whats-coming-up-for-word-add-ins"></a><span data-ttu-id="ed87b-158">O que est? surgindo para os suplementos do Word?</span><span class="sxs-lookup"><span data-stu-id="ed87b-158">What's coming up for Word add-ins?</span></span>

<span data-ttu-id="ed87b-p113">? medida que criamos e desenvolvemos novas APIs para suplementos do Word, elas ficam dispon?veis na nossa p?gina [Especifica??es abertas da API](https://dev.office.com/reference/add-ins/openspec) para voc? deixar seus coment?rios. Descubra que novos recursos est?o no pipeline para as APIs JavaScript do Word e forne?a coment?rios sobre nossas especifica??es de design.</span><span class="sxs-lookup"><span data-stu-id="ed87b-p113">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [API open specifications](https://dev.office.com/reference/add-ins/openspec) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="ed87b-161">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="ed87b-161">See also</span></span>

* [<span data-ttu-id="ed87b-162">Vis?o geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ed87b-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="ed87b-163">Refer?ncias da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="ed87b-163">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

