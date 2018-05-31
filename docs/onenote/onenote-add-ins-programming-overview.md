---
title: Visão geral da programação da API JavaScript do OneNote
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: aded1210abc11a80c6200a207d3896df8ef4218b
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19439009"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="6021c-102">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="6021c-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="6021c-103">O OneNote introduz uma API JavaScript para os suplementos do OneNote Online. Você pode criar suplementos de painel de tarefas e de conteúdo e comandos de suplemento que interagem com objetos do OneNote e conectam-se a serviços Web ou a outros recursos baseados na Web.</span><span class="sxs-lookup"><span data-stu-id="6021c-103">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="6021c-p101">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="6021c-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="6021c-106">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="6021c-106">Components of an Office Add-in</span></span>

<span data-ttu-id="6021c-107">Os suplementos consistem de dois componentes básicos:</span><span class="sxs-lookup"><span data-stu-id="6021c-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="6021c-p102">Um **aplicativo Web** consiste em uma página da Web e em JavaScript, CSS ou outros arquivos necessários. Estes arquivos podem ser hospedados em qualquer servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure. No OneNote Online, o aplicativo Web exibe um controle de navegação ou iframe.</span><span class="sxs-lookup"><span data-stu-id="6021c-p102">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote Online, the web application displays in a browser control or iframe.</span></span>
    
- <span data-ttu-id="6021c-p103">Um **manifesto XML** que especifica a URL da página da Web do suplemento e os requisitos de acesso, as configurações e os recursos para o suplemento. Este arquivo é armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](../develop/add-in-manifests.md) como outros suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="6021c-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="6021c-114">**Suplemento do Office = manifesto + página da Web**</span><span class="sxs-lookup"><span data-stu-id="6021c-114">**Office Add-in = Manifest + Webpage**</span></span>

![Um suplemento do Office consiste em um manifesto e uma página da Web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="6021c-116">Usar a API JavaScript</span><span class="sxs-lookup"><span data-stu-id="6021c-116">Using the JavaScript API</span></span>

<span data-ttu-id="6021c-p104">Os suplementos usam o contexto de tempo de execução do aplicativo host para acessar a API JavaScript. A API tem duas camadas:</span><span class="sxs-lookup"><span data-stu-id="6021c-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="6021c-119">Uma **API avançada** para operações específicas do OneNote, acessada por meio do objeto **Application**.</span><span class="sxs-lookup"><span data-stu-id="6021c-119">A **rich API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="6021c-120">Uma **API comum** compartilhada entre os aplicativos do Office, acessada por meio do objeto **Document**.</span><span class="sxs-lookup"><span data-stu-id="6021c-120">A **common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-rich-api-through-the-application-object"></a><span data-ttu-id="6021c-121">Acessar uma API avançada por meio do objeto *Application*.</span><span class="sxs-lookup"><span data-stu-id="6021c-121">Accessing the rich API through the *Application* object</span></span>

<span data-ttu-id="6021c-p105">Use o objeto **Application** para acessar os objetos do OneNote, como **Notebook**, **Section** e **Page**. Com as APIs avançadas, você executa operações em lotes em objetos proxy. O fluxo básico será semelhante a:</span><span class="sxs-lookup"><span data-stu-id="6021c-p105">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With rich APIs, you run batch operations on proxy objects. The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="6021c-125">Obtenha a instância do aplicativo do contexto.</span><span class="sxs-lookup"><span data-stu-id="6021c-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="6021c-p106">Crie um proxy que representa o objeto do OneNote com o qual você deseja trabalhar. Você interage com sincronia com os objetos proxy lendo e gravar suas propriedades e chamando seus métodos.</span><span class="sxs-lookup"><span data-stu-id="6021c-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span> 

3. <span data-ttu-id="6021c-p107">Chame **load** no proxy para preenchê-lo com valores de propriedade especificados no parâmetro. Essa chamada é adicionada à fila de comandos.</span><span class="sxs-lookup"><span data-stu-id="6021c-p107">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="6021c-130">Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas à fila.</span><span class="sxs-lookup"><span data-stu-id="6021c-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="6021c-p108">Chame **context.sync** para executar todos os comandos na fila na ordem em que eles estão. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais estão sendo executados, além de recuperar as propriedades dos objetos do OneNote carregados para uso no seu script. Você pode usar o objeto promessa retornado para o encadeamento ações adicionais.</span><span class="sxs-lookup"><span data-stu-id="6021c-p108">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="6021c-134">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="6021c-134">For example:</span></span> 

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

<span data-ttu-id="6021c-135">Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span><span class="sxs-lookup"><span data-stu-id="6021c-135">You can find supported OneNote objects and operations in the [API reference](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="6021c-136">Acessar a API comum por meio do objeto *Document*</span><span class="sxs-lookup"><span data-stu-id="6021c-136">Accessing the common API through the *Document* object</span></span>

<span data-ttu-id="6021c-137">Use o objeto **Document** para acessar a API comum, como os métodos [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) e [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync).</span><span class="sxs-lookup"><span data-stu-id="6021c-137">Use the **Document** object to access the common API, such as the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) and [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) methods.</span></span> 

<span data-ttu-id="6021c-138">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="6021c-138">For example:</span></span>  

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
<span data-ttu-id="6021c-139">Os suplementos do OneNote são compatíveis apenas com as seguintes APIs comuns:</span><span class="sxs-lookup"><span data-stu-id="6021c-139">OneNote add-ins support only the following common APIs:</span></span>

| <span data-ttu-id="6021c-140">API</span><span class="sxs-lookup"><span data-stu-id="6021c-140">API</span></span> | <span data-ttu-id="6021c-141">Observações</span><span class="sxs-lookup"><span data-stu-id="6021c-141">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="6021c-142">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6021c-142">Office.context.document.getSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) | <span data-ttu-id="6021c-143">Apenas **Office.CoercionType.Text** e **Office.CoercionType.Matrix**</span><span class="sxs-lookup"><span data-stu-id="6021c-143">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="6021c-144">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6021c-144">Office.context.document.setSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) | <span data-ttu-id="6021c-145">Apenas **Office.CoercionType.Text**, **Office.CoercionType.Image** e **Office.CoercionType.Html**</span><span class="sxs-lookup"><span data-stu-id="6021c-145">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="6021c-146">var mySetting = Office.context.document.settings.get(nome);</span><span class="sxs-lookup"><span data-stu-id="6021c-146">var mySetting = Office.context.document.settings.get(name);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.get) | <span data-ttu-id="6021c-147">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="6021c-147">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="6021c-148">Office.context.document.settings.set(nome, valor);</span><span class="sxs-lookup"><span data-stu-id="6021c-148">Office.context.document.settings.set(name, value);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.set) | <span data-ttu-id="6021c-149">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="6021c-149">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="6021c-150">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="6021c-150">Office.EventType.DocumentSelectionChanged</span></span>](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

<span data-ttu-id="6021c-p109">Em geral, você só pode usar a API comum para fazer algo que não seja compatível com a API avançada. Para saber mais sobre como usar a API comum, confira os suplementos do Office [documentação](../overview/office-add-ins.md) e [referência](https://dev.office.com/reference/add-ins/javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="6021c-p109">In general, you only use the common API to do something that isn't supported in the rich API. To learn more about using the common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](https://dev.office.com/reference/add-ins/javascript-api-for-office).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="6021c-153">Diagrama do modelo de objeto do OneNote</span><span class="sxs-lookup"><span data-stu-id="6021c-153">OneNote object model diagram</span></span> 
<span data-ttu-id="6021c-154">O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.</span><span class="sxs-lookup"><span data-stu-id="6021c-154">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="6021c-156">Veja também</span><span class="sxs-lookup"><span data-stu-id="6021c-156">See also</span></span>

- [<span data-ttu-id="6021c-157">Criar seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="6021c-157">Build your first OneNote add-in</span></span>](onenote-add-ins-getting-started.md)
- [<span data-ttu-id="6021c-158">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="6021c-158">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="6021c-159">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="6021c-159">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="6021c-160">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6021c-160">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
