---
title: Visão geral da API JavaScript do Visio
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: f530e1541146fb2f38209845791132b789577640
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432932"
---
# <a name="visio-javascript-api-overview"></a><span data-ttu-id="ec429-102">Visão geral da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="ec429-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="ec429-103">Você pode usar as APIs JavaScript do Visio para inserir diagramas do Visio no SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="ec429-103">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="ec429-104">Um diagrama integrado do Visio é um diagrama armazenado em uma biblioteca de documentos do SharePoint e exibido em uma página do SharePoint.</span><span class="sxs-lookup"><span data-stu-id="ec429-104">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="ec429-105">Para integrar um diagrama do Visio, exiba-o em um elemento `<iframe>` de HTML.</span><span class="sxs-lookup"><span data-stu-id="ec429-105">To embed a Visio diagram, display it in an HTML `<iframe>`iframe element.</span></span> <span data-ttu-id="ec429-106">Em seguida, você pode usar APIs JavaScript do Visio para trabalhar via programação com o diagrama integrado.</span><span class="sxs-lookup"><span data-stu-id="ec429-106">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Diagrama do Visio em um iframe na página do SharePoint junto com a Web Part do editor de script](../images/visio-api-block-diagram.png)


<span data-ttu-id="ec429-108">É possível usar as APIs JavaScript do Visio para:</span><span class="sxs-lookup"><span data-stu-id="ec429-108">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="ec429-109">Interagir com os elementos de diagrama do Visio, como páginas e formas.</span><span class="sxs-lookup"><span data-stu-id="ec429-109">Interact with Visio diagram elements like pages and shapes</span></span>
* <span data-ttu-id="ec429-110">Criar uma marcação visual na tela do diagrama do Visio.</span><span class="sxs-lookup"><span data-stu-id="ec429-110">Create visual markup on the Visio diagram canvas</span></span>
* <span data-ttu-id="ec429-111">Adicionar manipuladores personalizados para eventos com o mouse no desenho.</span><span class="sxs-lookup"><span data-stu-id="ec429-111">Write custom handlers for mouse events within the drawing</span></span>
* <span data-ttu-id="ec429-112">Expôr dados de diagrama, como texto da forma, dados da forma e hiperlinks, em sua solução.</span><span class="sxs-lookup"><span data-stu-id="ec429-112">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="ec429-p102">Este artigo descreve como usar as APIs JavaScript do Visio com o Visio Online para desenvolver suas soluções para o SharePoint Online. Ele apresenta os principais conceitos que são fundamentais para o uso das APIs, como **EmbeddedSession**, **RequestContext**, e dos objetos proxy do JavaScript, além dos métodos **sync()**, **Visio.run()** e **load()**. Os exemplos de código mostram como aplicar esses conceitos.</span><span class="sxs-lookup"><span data-stu-id="ec429-p102">This article describes how to use the Visio JavaScript APIs with Visio Online to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as **EmbeddedSession**, **RequestContext**, and JavaScript proxy objects, and the **sync()**, **Visio.run()**, and **load()** methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="ec429-116">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="ec429-116">EmbeddedSession</span></span>

<span data-ttu-id="ec429-117">O objeto EmbeddedSession inicia a comunicação entre o quadro do desenvolvedor e o quadro do Visio Online.</span><span class="sxs-lookup"><span data-stu-id="ec429-117">The EmbeddedSession object initializes communication between the developer frame and the Visio Online frame.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="ec429-118">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="ec429-118">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="ec429-119">O método **Visio.run()** executa um script em lotes que realiza ações no modelo de objeto do Visio.</span><span class="sxs-lookup"><span data-stu-id="ec429-119">**Visio.run()** executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="ec429-120">Os comandos em lotes incluem definições de objetos proxy JavaScript locais e métodos **sync()** que sincronizam o estado entre objetos locais e do Visio, e a resolução de promessa.</span><span class="sxs-lookup"><span data-stu-id="ec429-120">The batch commands include definitions of local JavaScript proxy objects and **sync()** methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="ec429-121">A vantagem do envio de solicitações em lotes com o método **Visio.run()** é que, quando a promessa é resolvida, todos os objetos de página controlados que foram alocados durante a execução são automaticamente liberados.</span><span class="sxs-lookup"><span data-stu-id="ec429-121">The advantage of batching requests in **Visio.run()** is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="ec429-122">O método de execução recebe a sessão e o objeto RequestContext e retorna uma promessa (normalmente, apenas o resultado de **context.sync()**).</span><span class="sxs-lookup"><span data-stu-id="ec429-122">The run method takes in RequestContext and returns a promise (typically, just the result of **ctx.sync()**).</span></span> <span data-ttu-id="ec429-123">É possível executar a operação em lotes fora do **Visio.run()**.</span><span class="sxs-lookup"><span data-stu-id="ec429-123">It is possible to run the batch operation outside of the **Visio.run()**.</span></span> <span data-ttu-id="ec429-124">No entanto, todas as referências aos objetos de página devem ser rastreadas e gerenciadas manualmente nesse cenário.</span><span class="sxs-lookup"><span data-stu-id="ec429-124">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="ec429-125">RequestContext</span><span class="sxs-lookup"><span data-stu-id="ec429-125">RequestContext</span></span>

<span data-ttu-id="ec429-126">O objeto RequestContext facilita as solicitações para o aplicativo do Visio.</span><span class="sxs-lookup"><span data-stu-id="ec429-126">Request Context: The RequestContext object facilitates requests to the Excel application.</span></span> <span data-ttu-id="ec429-127">Como o quadro de desenvolvedor e o aplicativo do Visio Online são executados em dois iframes diferentes, o objeto RequestContext (contexto no próximo exemplo) é necessário para obter acesso ao Visio e aos objetos relacionados do quadro de desenvolvedor, como páginas e formas.</span><span class="sxs-lookup"><span data-stu-id="ec429-127">The RequestContext object facilitates requests to the Visio application. Because the developer frame and the Visio Online application run in two different iframes, request context is required to get access to Visio and related objects such as pages and shapes, from the developer frame. The following example shows how to create a request context.</span></span>

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a><span data-ttu-id="ec429-128">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="ec429-128">Proxy objects</span></span>

<span data-ttu-id="ec429-p106">Os objetos JavaScript do Visio declarados e usados em um suplemento são objetos proxy dos objetos reais de um documento do Visio. Todas as ações executadas em objetos proxy não são percebidas no Visio, e o estado do documento do Visio não é percebido em objetos proxy, até que o estado do documento tenha sido sincronizado. O estado do documento é sincronizado quando `context.sync()` é executado.</span><span class="sxs-lookup"><span data-stu-id="ec429-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="ec429-132">Por exemplo, o objeto JavaScript local getActivePage é declarado para fazer referência à página selecionada.</span><span class="sxs-lookup"><span data-stu-id="ec429-132">For example, the local JavaScript object  is declared to reference the selected range.</span></span> <span data-ttu-id="ec429-133">Você pode usá-lo para colocar a configuração das respectivas propriedades em fila e para invocar métodos.</span><span class="sxs-lookup"><span data-stu-id="ec429-133">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="ec429-134">As ações são realizadas sobre esses objetos apenas quando o método **sync()** é executado.</span><span class="sxs-lookup"><span data-stu-id="ec429-134">The actions on such objects are not realized until the sync() method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="ec429-135">sync()</span><span class="sxs-lookup"><span data-stu-id="ec429-135">sync()</span></span>

<span data-ttu-id="ec429-136">O método **sync()** sincroniza o estado entre objetos proxy JavaScript e objetos reais no Visio, com a execução de instruções enfileiradas no contexto e com a recuperação de propriedades de objetos carregados do Office para uso no código.</span><span class="sxs-lookup"><span data-stu-id="ec429-136">The **sync()** method, available on the request context, synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="ec429-137">Este método retorna uma promessa, que é resolvida quando o sistema conclui a sincronização.</span><span class="sxs-lookup"><span data-stu-id="ec429-137">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="ec429-138">load()</span><span class="sxs-lookup"><span data-stu-id="ec429-138">load()</span></span>

<span data-ttu-id="ec429-p109">O método **load()** é usado para preencher os objetos proxy criados na camada JavaScript do suplemento. Ao tentar recuperar um objeto, como um documento, um objeto proxy local é criado inicialmente na camada JavaScript. Você pode usar esse objeto para colocar a configuração das respectivas propriedades em fila e para invocar métodos. No entanto, você deve invocar inicialmente os métodos **load()** e **sync()** para as relações ou propriedades do objeto de leitura. O método load() realizado nas propriedades e relações que devem ser carregadas quando você chama o método **sync()**.</span><span class="sxs-lookup"><span data-stu-id="ec429-p109">The **load()** method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the **load()** and **sync()** methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the **sync()** method is called.</span></span>

<span data-ttu-id="ec429-144">O seguinte mostra a sintaxe para o método **load()**.</span><span class="sxs-lookup"><span data-stu-id="ec429-144">The following shows the syntax for the **load()** method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="ec429-145">**properties** é a lista de nomes de propriedades a carregar, especificados como cadeias de caracteres delimitadas por vírgulas ou por uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="ec429-145">**properties** is the list of properties and/or relationship names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="ec429-146">Consulte os métodos **.load()** em cada objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="ec429-146">See **.load()** methods under each object for details.</span></span>

2. <span data-ttu-id="ec429-p111">**loadOption** especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Consulte as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="ec429-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="ec429-149">Exemplo: imprimir todo o texto de formas na página ativa</span><span class="sxs-lookup"><span data-stu-id="ec429-149">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="ec429-150">O exemplo a seguir mostra como imprimir o valor de texto de forma de um objeto de formas de matriz.</span><span class="sxs-lookup"><span data-stu-id="ec429-150">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="ec429-151">O método **Visio.run()** inclui um lote de instruções.</span><span class="sxs-lookup"><span data-stu-id="ec429-151">The **Visio.run()** method contains a batch of instructions.</span></span> <span data-ttu-id="ec429-152">Como parte deste lote, o sistema cria um objeto proxy que faz referência a formas no documento ativo.</span><span class="sxs-lookup"><span data-stu-id="ec429-152">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="ec429-153">Todos esses comandos são enfileirados e executados ao chamar o método **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="ec429-153">All these commands are queued and run when **ctx.sync()** is called.</span></span> <span data-ttu-id="ec429-154">O método **sync()** retorna uma promessa que pode ser usada para encadeamento com outras operações.</span><span class="sxs-lookup"><span data-stu-id="ec429-154">The **sync()** method returns a promise that can be used to chain it with other operations.</span></span>

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a><span data-ttu-id="ec429-155">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="ec429-155">Error Messages</span></span>

<span data-ttu-id="ec429-p114">O sistema retorna erros usando um objeto Error composto por um código e uma mensagem. A tabela a seguir fornece uma lista de possíveis condições de erro que podem ocorrer.</span><span class="sxs-lookup"><span data-stu-id="ec429-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="ec429-158">error.code</span><span class="sxs-lookup"><span data-stu-id="ec429-158">error.code</span></span>            | <span data-ttu-id="ec429-159">error.message</span><span class="sxs-lookup"><span data-stu-id="ec429-159">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="ec429-160">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="ec429-160">InvalidArgument</span></span>       | <span data-ttu-id="ec429-161">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="ec429-161">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="ec429-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ec429-162">GeneralException</span></span>      | <span data-ttu-id="ec429-163">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="ec429-163">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="ec429-164">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="ec429-164">NotImplemented</span></span>        | <span data-ttu-id="ec429-165">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="ec429-165">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="ec429-166">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="ec429-166">UnsupportedOperation</span></span>  | <span data-ttu-id="ec429-167">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="ec429-167">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="ec429-168">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="ec429-168">AccessDenied</span></span>          | <span data-ttu-id="ec429-169">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="ec429-169">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="ec429-170">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="ec429-170">ItemNotFound</span></span>          | <span data-ttu-id="ec429-171">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="ec429-171">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="ec429-172">Introdução</span><span class="sxs-lookup"><span data-stu-id="ec429-172">Get started</span></span>

<span data-ttu-id="ec429-173">Você pode usar o exemplo nesta seção para começar.</span><span class="sxs-lookup"><span data-stu-id="ec429-173">You can use the example in this section to get started.</span></span> <span data-ttu-id="ec429-174">Este exemplo mostra como exibir o texto da forma selecionada em um diagrama do Visio via programação.</span><span class="sxs-lookup"><span data-stu-id="ec429-174">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="ec429-175">Para começar, crie uma página clássica no SharePoint Online ou edite uma página existente.</span><span class="sxs-lookup"><span data-stu-id="ec429-175">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="ec429-176">Adicione uma Web Part de editor de script à página e copie e cole o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="ec429-176">Add a script editor webpart on the page and copy-paste the following code.</span></span>

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

<span data-ttu-id="ec429-177">Depois disso, você só precisa da URL de um diagrama do Visio com o qual deseja trabalhar.</span><span class="sxs-lookup"><span data-stu-id="ec429-177">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="ec429-178">Basta carregar o diagrama do Visio no SharePoint Online e abri-lo no Visio Online.</span><span class="sxs-lookup"><span data-stu-id="ec429-178">Just upload the Visio diagram to SharePoint Online and open it in Visio Online.</span></span> <span data-ttu-id="ec429-179">A partir daí, abra a caixa de diálogo Inserir e use a URL de integração do exemplo acima.</span><span class="sxs-lookup"><span data-stu-id="ec429-179">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Copiar a URL do arquivo do Visio da caixa de diálogo Inserir](../images/Visio-embed-url.png)

<span data-ttu-id="ec429-181">Se você estiver usando o Visio Online no modo de edição, abra a caixa de diálogo Inserir escolhendo **Arquivo** > **Compartilhar** > **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="ec429-181">If you are using Visio Online in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="ec429-182">Se você estiver usando o Visio Online no modo de exibição, abra a caixa de diálogo Inserir escolhendo '...' e **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="ec429-182">If you are using Visio Online in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="ec429-183">Especificações abertas da API</span><span class="sxs-lookup"><span data-stu-id="ec429-183">Open API specifications</span></span>

<span data-ttu-id="ec429-p118">À medida que criamos e desenvolvemos novas APIs, elas ficam disponíveis na nossa página [Especificações abertas da API](../openspec.md) para recebe seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.</span><span class="sxs-lookup"><span data-stu-id="ec429-p118">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="ec429-186">Referência da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="ec429-186">OneNote JavaScript API reference</span></span>

<span data-ttu-id="ec429-187">Para saber mais sobre a API JavaScript do Visio, consulte a [Documentação de referência da API JavaScript do Visio](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="ec429-187">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/visio).</span></span>
