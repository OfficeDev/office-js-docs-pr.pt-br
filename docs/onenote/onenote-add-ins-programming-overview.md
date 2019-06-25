---
title: Visão geral da programação da API JavaScript do OneNote
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 852c68bc9edf370d0eef687fb4869b23d4f59fe4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128633"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="a8c5a-102">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8c5a-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="a8c5a-103">O OneNote introduz uma API do JavaScript para suplementos do OneNote na Web.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-103">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="a8c5a-104">Você pode criar suplementos de painel de tarefas e de conteúdo e comandos de suplemento que interagem com objetos do OneNote e conectam-se a serviços Web ou a outros recursos baseados na Web.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-104">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="a8c5a-p102">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="a8c5a-107">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="a8c5a-107">Components of an Office Add-in</span></span>

<span data-ttu-id="a8c5a-108">Os suplementos consistem de dois componentes básicos:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-108">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="a8c5a-109">Um **aplicativo Web** consiste em uma página da Web e em JavaScript, CSS ou outros arquivos necessários.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-109">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="a8c5a-110">Estes arquivos podem ser hospedados em qualquer servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-110">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="a8c5a-111">No OneNote online, o aplicativo Web exibe um controle de navegação ou iframe.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-111">In OneNote Online, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="a8c5a-p104">Um **manifesto XML** que especifica a URL da página da Web do suplemento e os requisitos de acesso, as configurações e os recursos para o suplemento. Este arquivo é armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](../develop/add-in-manifests.md) como outros suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="a8c5a-115">**Suplemento do Office = manifesto + página da Web**</span><span class="sxs-lookup"><span data-stu-id="a8c5a-115">**Office Add-in = Manifest + Webpage**</span></span>

![Um suplemento do Office consiste em um manifesto e uma página da Web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="a8c5a-117">Usar a API JavaScript</span><span class="sxs-lookup"><span data-stu-id="a8c5a-117">Using the JavaScript API</span></span>

<span data-ttu-id="a8c5a-p105">Os suplementos usam o contexto de tempo de execução do aplicativo host para acessar a API JavaScript. A API tem duas camadas:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="a8c5a-120">Uma **API avançada** para operações específicas do OneNote, acessada por meio do objeto **Aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-120">A **host-specific API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="a8c5a-121">Uma **API comum** compartilhada entre os aplicativos do Office, acessada por meio do objeto **Documento**.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-121">A **Common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="a8c5a-122">Acessar uma API avançada por meio do objeto *Aplicativo*.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-122">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="a8c5a-123">Use o objeto **Aplicativo** para acessar os objetos do OneNote, como **Bloco de anotações**, **Seção** e **Página**.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-123">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="a8c5a-124">Com as APIs avançadas, você executa operações em lotes em objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-124">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="a8c5a-125">O fluxo básico será semelhante a:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-125">The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="a8c5a-126">Obtenha a instância do aplicativo do contexto.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-126">Get the application instance from the context.</span></span>

2. <span data-ttu-id="a8c5a-p107">Crie um proxy que representa o objeto do OneNote com o qual você deseja trabalhar. Você interage com sincronia com os objetos proxy lendo e gravar suas propriedades e chamando seus métodos.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="a8c5a-p108">Chame **load** no proxy para preenchê-lo com valores de propriedade especificados no parâmetro. Essa chamada é adicionada à fila de comandos.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p108">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="a8c5a-131">Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas à fila.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-131">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="a8c5a-p109">Chame **context.sync** para executar todos os comandos na fila na ordem em que eles estão. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais estão sendo executados, além de recuperar as propriedades dos objetos do OneNote carregados para uso no seu script. Você pode usar o objeto promessa retornado para o encadeamento ações adicionais.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-p109">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="a8c5a-135">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-135">For example:</span></span>

```js
function getPagesInSection() {
    OneNote.run(function (context) {

        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Read the id and title of each page.
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId);
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

<span data-ttu-id="a8c5a-136">Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="a8c5a-136">You can find supported OneNote objects and operations in the [API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="a8c5a-137">Acessar a API comum por meio do objeto *Documento*</span><span class="sxs-lookup"><span data-stu-id="a8c5a-137">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="a8c5a-138">Use o objeto **Documento** para acessar a API comum, como os métodos [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) e [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="a8c5a-138">Use the **Document** object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span> 


<span data-ttu-id="a8c5a-139">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-139">For example:</span></span>  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

<span data-ttu-id="a8c5a-140">Os suplementos do OneNote são compatíveis apenas com as seguintes APIs comuns:</span><span class="sxs-lookup"><span data-stu-id="a8c5a-140">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="a8c5a-141">API</span><span class="sxs-lookup"><span data-stu-id="a8c5a-141">API</span></span> | <span data-ttu-id="a8c5a-142">Observações</span><span class="sxs-lookup"><span data-stu-id="a8c5a-142">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="a8c5a-143">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a8c5a-143">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="a8c5a-144">Apenas **Office.CoercionType.Text** e **Office.CoercionType.Matrix**</span><span class="sxs-lookup"><span data-stu-id="a8c5a-144">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="a8c5a-145">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a8c5a-145">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="a8c5a-146">Apenas **Office.CoercionType.Text**, **Office.CoercionType.Image** e **Office.CoercionType.Html**</span><span class="sxs-lookup"><span data-stu-id="a8c5a-146">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="a8c5a-147">var mySetting = Office.context.document.settings.get(nome);</span><span class="sxs-lookup"><span data-stu-id="a8c5a-147">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="a8c5a-148">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="a8c5a-148">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="a8c5a-149">Office.context.document.settings.set(nome, valor);</span><span class="sxs-lookup"><span data-stu-id="a8c5a-149">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="a8c5a-150">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="a8c5a-150">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="a8c5a-151">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="a8c5a-151">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="a8c5a-152">Em geral, você só pode usar a API comum para fazer algo que não seja compatível com a API avançada.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-152">In general, you only use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="a8c5a-153">Para saber mais sobre como usar a API comum, confira a [documentação](../overview/office-add-ins.md) e a [referência](../reference/javascript-api-for-office.md) dos suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-153">To learn more about using the Common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](../reference/javascript-api-for-office.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="a8c5a-154">Diagrama do modelo de objeto do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8c5a-154">OneNote object model diagram</span></span> 
<span data-ttu-id="a8c5a-155">O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.</span><span class="sxs-lookup"><span data-stu-id="a8c5a-155">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="a8c5a-157">Confira também</span><span class="sxs-lookup"><span data-stu-id="a8c5a-157">See also</span></span>

- [<span data-ttu-id="a8c5a-158">Criar seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8c5a-158">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="a8c5a-159">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8c5a-159">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="a8c5a-160">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="a8c5a-160">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="a8c5a-161">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a8c5a-161">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
