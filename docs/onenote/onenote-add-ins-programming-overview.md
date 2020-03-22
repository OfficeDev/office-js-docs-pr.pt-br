---
title: Visão geral da programação da API JavaScript do OneNote
description: Saiba mais sobre a API JavaScript do OneNote para suplementos do OneNote na Web.
ms.date: 03/18/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: ae88c2bba6c23a2c3ec3358db121a2ca3630f09d
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891051"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="4bacf-103">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4bacf-103">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="4bacf-104">O OneNote introduz uma API do JavaScript para suplementos do OneNote na Web.</span><span class="sxs-lookup"><span data-stu-id="4bacf-104">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="4bacf-105">Você pode criar suplementos de painel de tarefas e de conteúdo e comandos de suplemento que interagem com objetos do OneNote e conectam-se a serviços Web ou a outros recursos baseados na Web.</span><span class="sxs-lookup"><span data-stu-id="4bacf-105">You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="4bacf-106">Componentes de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="4bacf-106">Components of an Office Add-in</span></span>

<span data-ttu-id="4bacf-107">Os suplementos consistem de dois componentes básicos:</span><span class="sxs-lookup"><span data-stu-id="4bacf-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="4bacf-108">Um **aplicativo Web** consiste em uma página da Web e em JavaScript, CSS ou outros arquivos necessários.</span><span class="sxs-lookup"><span data-stu-id="4bacf-108">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="4bacf-109">Estes arquivos podem ser hospedados em qualquer servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="4bacf-109">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="4bacf-110">No OneNote online, o aplicativo Web exibe um controle de navegação ou iframe.</span><span class="sxs-lookup"><span data-stu-id="4bacf-110">In OneNote on the web, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="4bacf-p103">Um **manifesto XML** que especifica a URL da página da Web do suplemento e os requisitos de acesso, as configurações e os recursos para o suplemento. Este arquivo é armazenado no cliente. Os suplementos do OneNote usam o mesmo formato de [manifesto](../develop/add-in-manifests.md) como outros suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="4bacf-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="4bacf-114">**Suplemento do Office = manifesto + página da Web**</span><span class="sxs-lookup"><span data-stu-id="4bacf-114">**Office Add-in = Manifest + Webpage**</span></span>

![Um suplemento do Office consiste em um manifesto e uma página da Web](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="4bacf-116">Usar a API JavaScript</span><span class="sxs-lookup"><span data-stu-id="4bacf-116">Using the JavaScript API</span></span>

<span data-ttu-id="4bacf-p104">Os suplementos usam o contexto de tempo de execução do aplicativo host para acessar a API JavaScript. A API tem duas camadas:</span><span class="sxs-lookup"><span data-stu-id="4bacf-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span>

- <span data-ttu-id="4bacf-119">Uma **API específica do host** para operações específicas do OneNote, acessada por meio do objeto `Application`.</span><span class="sxs-lookup"><span data-stu-id="4bacf-119">A **host-specific API** for OneNote-specific operations, accessed through the `Application` object.</span></span>
- <span data-ttu-id="4bacf-120">Uma **API comum** compartilhada entre aplicativos do Office, acessada por meio do objeto `Document`.</span><span class="sxs-lookup"><span data-stu-id="4bacf-120">A **Common API** that's shared across Office applications, accessed through the `Document` object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="4bacf-121">Acessar uma API avançada por meio do objeto *Aplicativo*.</span><span class="sxs-lookup"><span data-stu-id="4bacf-121">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="4bacf-122">Use o objeto `Application` para acessar objetos do OneNote, como **Bloco de Anotações**, **Seção** e **Página**.</span><span class="sxs-lookup"><span data-stu-id="4bacf-122">Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="4bacf-123">Com as APIs avançadas, você executa operações em lotes em objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="4bacf-123">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="4bacf-124">O fluxo básico será semelhante a:</span><span class="sxs-lookup"><span data-stu-id="4bacf-124">The basic flow goes something like this:</span></span>

1. <span data-ttu-id="4bacf-125">Obtenha a instância do aplicativo do contexto.</span><span class="sxs-lookup"><span data-stu-id="4bacf-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="4bacf-p106">Crie um proxy que representa o objeto do OneNote com o qual você deseja trabalhar. Você interage com sincronia com os objetos proxy lendo e gravar suas propriedades e chamando seus métodos.</span><span class="sxs-lookup"><span data-stu-id="4bacf-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="4bacf-p107">Chame `load` no proxy para preenchê-lo com valores de propriedade especificados no parâmetro. Essa chamada é adicionada à fila de comandos.</span><span class="sxs-lookup"><span data-stu-id="4bacf-p107">Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="4bacf-130">Chamadas de método para a API (como `context.application.getActiveSection().pages;`) também são adicionadas à fila.</span><span class="sxs-lookup"><span data-stu-id="4bacf-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="4bacf-p108">Chame `context.sync` para executar todos os comandos na fila na ordem em que eles estão. Isso sincroniza o estado entre o momento em que os scripts e os objetos reais estão sendo executados, além de recuperar as propriedades dos objetos do OneNote carregados para uso no seu script. Você pode usar o objeto promessa retornado para o encadeamento ações adicionais.</span><span class="sxs-lookup"><span data-stu-id="4bacf-p108">Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="4bacf-134">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="4bacf-134">For example:</span></span>

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

<span data-ttu-id="4bacf-135">Você pode encontrar objetos do OneNote e operações compatíveis na [Referência API](../reference/overview/onenote-add-ins-javascript-reference.md).</span><span class="sxs-lookup"><span data-stu-id="4bacf-135">You can find supported OneNote objects and operations in the [API reference](../reference/overview/onenote-add-ins-javascript-reference.md).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="4bacf-136">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4bacf-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="4bacf-137">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="4bacf-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="4bacf-138">Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento.</span><span class="sxs-lookup"><span data-stu-id="4bacf-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="4bacf-139">Para saber mais sobre conjuntos de requisitos da API JavaScript do OneNote, consulte [Conjuntos de requisitos da API JavaScript do OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="4bacf-139">For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="4bacf-140">Acessar a API comum por meio do objeto *Documento*</span><span class="sxs-lookup"><span data-stu-id="4bacf-140">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="4bacf-141">Use o objeto `Document`para acessar a API comum, como os métodos [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) e [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="4bacf-141">Use the `Document` object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span>


<span data-ttu-id="4bacf-142">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="4bacf-142">For example:</span></span>  

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

<span data-ttu-id="4bacf-143">Os suplementos do OneNote são compatíveis apenas com as seguintes APIs comuns:</span><span class="sxs-lookup"><span data-stu-id="4bacf-143">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="4bacf-144">API</span><span class="sxs-lookup"><span data-stu-id="4bacf-144">API</span></span> | <span data-ttu-id="4bacf-145">Observações</span><span class="sxs-lookup"><span data-stu-id="4bacf-145">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="4bacf-146">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bacf-146">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="4bacf-147">Somente `Office.CoercionType.Text` e `Office.CoercionType.Matrix`</span><span class="sxs-lookup"><span data-stu-id="4bacf-147">`Office.CoercionType.Text` and `Office.CoercionType.Matrix` only</span></span> |
| [<span data-ttu-id="4bacf-148">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bacf-148">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="4bacf-149">Somente `Office.CoercionType.Text`, `Office.CoercionType.Image` e `Office.CoercionType.Html`</span><span class="sxs-lookup"><span data-stu-id="4bacf-149">`Office.CoercionType.Text`, `Office.CoercionType.Image`, and `Office.CoercionType.Html` only</span></span> | 
| [<span data-ttu-id="4bacf-150">var mySetting = Office.context.document.settings.get(nome);</span><span class="sxs-lookup"><span data-stu-id="4bacf-150">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="4bacf-151">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="4bacf-151">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="4bacf-152">Office.context.document.settings.set(nome, valor);</span><span class="sxs-lookup"><span data-stu-id="4bacf-152">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="4bacf-153">As configurações são compatíveis apenas com os suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="4bacf-153">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="4bacf-154">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="4bacf-154">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="4bacf-155">Em geral, você usa a API comum para fazer algo que não é compatível com a API específica do host.</span><span class="sxs-lookup"><span data-stu-id="4bacf-155">In general, you use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="4bacf-156">Para obter mais informações sobre como usar a API comum, confira [Modelo do objeto do JavaScript API para Office](../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="4bacf-156">To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="4bacf-157">Diagrama do modelo de objeto do OneNote</span><span class="sxs-lookup"><span data-stu-id="4bacf-157">OneNote object model diagram</span></span> 
<span data-ttu-id="4bacf-158">O diagrama a seguir representa o que está disponível atualmente na API JavaScript do OneNote.</span><span class="sxs-lookup"><span data-stu-id="4bacf-158">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![Diagrama do modelo de objeto do OneNote](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="4bacf-160">Confira também</span><span class="sxs-lookup"><span data-stu-id="4bacf-160">See also</span></span>

- [<span data-ttu-id="4bacf-161">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="4bacf-161">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="4bacf-162">Criar seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="4bacf-162">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="4bacf-163">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4bacf-163">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="4bacf-164">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="4bacf-164">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="4bacf-165">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4bacf-165">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
