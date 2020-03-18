---
title: Visão geral da API JavaScript do Visio
description: Visão geral da API JavaScript do Visio.
ms.date: 06/20/2019
ms.prod: visio
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 5a544d93c1a41f6c913381ee8d67d375646b2883
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717527"
---
# <a name="visio-javascript-api-overview"></a><span data-ttu-id="b59e3-103">Visão geral da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="b59e3-103">Visio JavaScript API overview</span></span>

<span data-ttu-id="b59e3-104">Você pode usar as APIs JavaScript do Visio para inserir diagramas do Visio no SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="b59e3-104">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="b59e3-105">Um diagrama integrado do Visio é um diagrama armazenado em uma biblioteca de documentos do SharePoint e exibido em uma página do SharePoint.</span><span class="sxs-lookup"><span data-stu-id="b59e3-105">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="b59e3-106">Para integrar um diagrama do Visio, exiba-o em um elemento `<iframe>` de HTML.</span><span class="sxs-lookup"><span data-stu-id="b59e3-106">To embed a Visio diagram, display it in an HTML `<iframe>` element.</span></span> <span data-ttu-id="b59e3-107">Em seguida, você pode usar APIs JavaScript do Visio para trabalhar via programação com o diagrama integrado.</span><span class="sxs-lookup"><span data-stu-id="b59e3-107">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Diagrama do Visio em um iframe na página do SharePoint junto com a Web Part do editor de script](../images/visio-api-block-diagram.png)


<span data-ttu-id="b59e3-109">É possível usar as APIs JavaScript do Visio para:</span><span class="sxs-lookup"><span data-stu-id="b59e3-109">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="b59e3-110">Interagir com os elementos de diagrama do Visio, como páginas e formas.</span><span class="sxs-lookup"><span data-stu-id="b59e3-110">Interact with Visio diagram elements like pages and shapes.</span></span>
* <span data-ttu-id="b59e3-111">Criar uma marcação visual na tela do diagrama do Visio.</span><span class="sxs-lookup"><span data-stu-id="b59e3-111">Create visual markup on the Visio diagram canvas.</span></span>
* <span data-ttu-id="b59e3-112">Adicionar manipuladores personalizados para eventos com o mouse no desenho.</span><span class="sxs-lookup"><span data-stu-id="b59e3-112">Write custom handlers for mouse events within the drawing.</span></span>
* <span data-ttu-id="b59e3-113">Expôr dados de diagrama, como texto da forma, dados da forma e hiperlinks, em sua solução.</span><span class="sxs-lookup"><span data-stu-id="b59e3-113">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="b59e3-p102">Este artigo descreve como usar as APIs JavaScript do Visio com o Visio na Web para desenvolver suas soluções para o SharePoint Online. Ele apresenta os principais conceitos que são fundamentais para o uso das APIs, como `EmbeddedSession`, `RequestContext` e dos objetos proxy do JavaScript, além dos métodos `sync()`, `Visio.run()`, and `load()`. Os exemplos de código mostram como aplicar esses conceitos.</span><span class="sxs-lookup"><span data-stu-id="b59e3-p102">This article describes how to use the Visio JavaScript APIs with Visio on the web to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as `EmbeddedSession`, `RequestContext`, and JavaScript proxy objects, and the `sync()`, `Visio.run()`, and `load()` methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="b59e3-117">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="b59e3-117">EmbeddedSession</span></span>

<span data-ttu-id="b59e3-118">O objeto EmbeddedSession inicia a comunicação entre o quadro do desenvolvedor e o quadro do Visio no navegador.</span><span class="sxs-lookup"><span data-stu-id="b59e3-118">The EmbeddedSession object initializes communication between the developer frame and the Visio frame in the browser.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="b59e3-119">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="b59e3-119">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="b59e3-120">`Visio.run()` executa um script em lote que executa ações no modelo de objeto do Visio.</span><span class="sxs-lookup"><span data-stu-id="b59e3-120">`Visio.run()` executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="b59e3-121">Os comandos em lotes incluem definições de objetos proxy JavaScript locais e métodos `sync()` que sincronizam o estado entre objetos locais e do Visio, e a resolução de promessa.</span><span class="sxs-lookup"><span data-stu-id="b59e3-121">The batch commands include definitions of local JavaScript proxy objects and `sync()` methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="b59e3-122">A vantagem do envio de solicitações em lotes com o método `Visio.run()` é que, quando a promessa é resolvida, todos os objetos de página controlados que foram alocados durante a execução são automaticamente liberados.</span><span class="sxs-lookup"><span data-stu-id="b59e3-122">The advantage of batching requests in `Visio.run()` is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="b59e3-123">O método run recebe a sessão e o objeto RequestContext e retorna uma promessa (normalmente, apenas o resultado de `context.sync()`).</span><span class="sxs-lookup"><span data-stu-id="b59e3-123">The run method takes in session and RequestContext object and returns a promise (typically, just the result of `context.sync()`).</span></span> <span data-ttu-id="b59e3-124">É possível executar a operação em lote fora do `Visio.run()`.</span><span class="sxs-lookup"><span data-stu-id="b59e3-124">It is possible to run the batch operation outside of the `Visio.run()`.</span></span> <span data-ttu-id="b59e3-125">No entanto, todas as referências aos objetos de página devem ser rastreadas e gerenciadas manualmente nesse cenário.</span><span class="sxs-lookup"><span data-stu-id="b59e3-125">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="b59e3-126">RequestContext</span><span class="sxs-lookup"><span data-stu-id="b59e3-126">RequestContext</span></span>

<span data-ttu-id="b59e3-127">O objeto RequestContext facilita as solicitações para o aplicativo do Visio.</span><span class="sxs-lookup"><span data-stu-id="b59e3-127">The RequestContext object facilitates requests to the Visio application.</span></span> <span data-ttu-id="b59e3-128">Como o quadro do desenvolvedor e o cliente Web do Visio são executados em dois iframes diferentes, o objeto RequestContext (contexto no próximo exemplo) é necessário para obter acesso ao Visio e aos objetos relacionados do quadro de desenvolvedor, como páginas e formas.</span><span class="sxs-lookup"><span data-stu-id="b59e3-128">Because the developer frame and the Visio web client run in two different iframes, the RequestContext object (context in next example) is required to get access to Visio and related objects such as pages and shapes, from the developer frame.</span></span>

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

## <a name="proxy-objects"></a><span data-ttu-id="b59e3-129">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="b59e3-129">Proxy objects</span></span>

<span data-ttu-id="b59e3-p106">Os objetos JavaScript do Visio declarados e usados em um suplemento são objetos proxy dos objetos reais de um documento do Visio. Todas as ações executadas em objetos proxy não são percebidas no Visio, e o estado do documento do Visio não é percebido em objetos proxy, até que o estado do documento tenha sido sincronizado. O estado do documento é sincronizado quando `context.sync()` é executado.</span><span class="sxs-lookup"><span data-stu-id="b59e3-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="b59e3-133">Por exemplo, o objeto JavaScript local getActivePage é declarado para fazer referência à página selecionada.</span><span class="sxs-lookup"><span data-stu-id="b59e3-133">For example, the local JavaScript object getActivePage is declared to reference the selected page.</span></span> <span data-ttu-id="b59e3-134">Você pode usá-lo para colocar a configuração das respectivas propriedades em fila e para invocar métodos.</span><span class="sxs-lookup"><span data-stu-id="b59e3-134">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="b59e3-135">As ações nesses objetos não são realizadas até que o método `sync()` seja executado.</span><span class="sxs-lookup"><span data-stu-id="b59e3-135">The actions on such objects are not realized until the `sync()` method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="b59e3-136">sync()</span><span class="sxs-lookup"><span data-stu-id="b59e3-136">sync()</span></span>

<span data-ttu-id="b59e3-137">O método `sync()` sincroniza o estado entre objetos proxy JavaScript e objetos reais no Visio, com a execução de instruções enfileiradas no contexto e com a recuperação de propriedades de objetos carregados do Office para uso no código.</span><span class="sxs-lookup"><span data-stu-id="b59e3-137">The `sync()` method synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="b59e3-138">Este método retorna uma promessa, que é resolvida quando o sistema conclui a sincronização.</span><span class="sxs-lookup"><span data-stu-id="b59e3-138">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="b59e3-139">load()</span><span class="sxs-lookup"><span data-stu-id="b59e3-139">load()</span></span>

<span data-ttu-id="b59e3-p109">O método `load()` é usado para preencher os objetos proxy criados na camada JavaScript do suplemento. Ao tentar recuperar um objeto, como um documento, um objeto proxy local é criado inicialmente na camada JavaScript. Você pode usar esse objeto para colocar a configuração das respectivas propriedades em fila e para invocar métodos. No entanto, você deve invocar inicialmente os métodos `load()` e `sync()` para as relações ou propriedades do objeto de leitura. O método load() realizado nas propriedades e relações que devem ser carregadas quando você chama o método `sync()`.</span><span class="sxs-lookup"><span data-stu-id="b59e3-p109">The `load()` method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the `load()` and `sync()` methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the `sync()` method is called.</span></span>

<span data-ttu-id="b59e3-145">A seguir, é mostrada a sintaxe do método `load()`.</span><span class="sxs-lookup"><span data-stu-id="b59e3-145">The following shows the syntax for the `load()` method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="b59e3-146">**properties** é a lista de nomes de propriedades a carregar, especificados como cadeias de caracteres delimitadas por vírgulas ou por uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="b59e3-146">**properties** is the list of property names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="b59e3-147">Veja os métodos `.load()` em cada objeto para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="b59e3-147">See `.load()` methods under each object for details.</span></span>

2. <span data-ttu-id="b59e3-p111">**loadOption** especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Consulte as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="b59e3-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="b59e3-150">Exemplo: imprimir todo o texto de formas na página ativa</span><span class="sxs-lookup"><span data-stu-id="b59e3-150">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="b59e3-151">O exemplo a seguir mostra como imprimir o valor de texto de forma de um objeto de formas de matriz.</span><span class="sxs-lookup"><span data-stu-id="b59e3-151">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="b59e3-152">O método `Visio.run()` contém um lote de instruções.</span><span class="sxs-lookup"><span data-stu-id="b59e3-152">The `Visio.run()` method contains a batch of instructions.</span></span> <span data-ttu-id="b59e3-153">Como parte deste lote, o sistema cria um objeto proxy que faz referência a formas no documento ativo.</span><span class="sxs-lookup"><span data-stu-id="b59e3-153">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="b59e3-154">Todos esses comandos são enfileirados e executados quando `context.sync()` é chamado.</span><span class="sxs-lookup"><span data-stu-id="b59e3-154">All these commands are queued and run when `context.sync()` is called.</span></span> <span data-ttu-id="b59e3-155">O método `sync()` retorna uma promessa que pode ser usada para encadeá-lo com outras operações.</span><span class="sxs-lookup"><span data-stu-id="b59e3-155">The `sync()` method returns a promise that can be used to chain it with other operations.</span></span>

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

## <a name="error-messages"></a><span data-ttu-id="b59e3-156">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="b59e3-156">Error messages</span></span>

<span data-ttu-id="b59e3-p114">O sistema retorna erros usando um objeto Error composto por um código e uma mensagem. A tabela a seguir fornece uma lista de possíveis condições de erro que podem ocorrer.</span><span class="sxs-lookup"><span data-stu-id="b59e3-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="b59e3-159">error.code</span><span class="sxs-lookup"><span data-stu-id="b59e3-159">error.code</span></span>            | <span data-ttu-id="b59e3-160">error.message</span><span class="sxs-lookup"><span data-stu-id="b59e3-160">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="b59e3-161">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="b59e3-161">InvalidArgument</span></span>       | <span data-ttu-id="b59e3-162">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="b59e3-162">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="b59e3-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b59e3-163">GeneralException</span></span>      | <span data-ttu-id="b59e3-164">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="b59e3-164">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="b59e3-165">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="b59e3-165">NotImplemented</span></span>        | <span data-ttu-id="b59e3-166">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="b59e3-166">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="b59e3-167">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="b59e3-167">UnsupportedOperation</span></span>  | <span data-ttu-id="b59e3-168">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="b59e3-168">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="b59e3-169">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="b59e3-169">AccessDenied</span></span>          | <span data-ttu-id="b59e3-170">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="b59e3-170">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="b59e3-171">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="b59e3-171">ItemNotFound</span></span>          | <span data-ttu-id="b59e3-172">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="b59e3-172">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="b59e3-173">Introdução</span><span class="sxs-lookup"><span data-stu-id="b59e3-173">Get started</span></span>

<span data-ttu-id="b59e3-174">Você pode usar o exemplo nesta seção para começar.</span><span class="sxs-lookup"><span data-stu-id="b59e3-174">You can use the example in this section to get started.</span></span> <span data-ttu-id="b59e3-175">Este exemplo mostra como exibir o texto da forma selecionada em um diagrama do Visio via programação.</span><span class="sxs-lookup"><span data-stu-id="b59e3-175">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="b59e3-176">Para começar, crie uma página clássica no SharePoint Online ou edite uma página existente.</span><span class="sxs-lookup"><span data-stu-id="b59e3-176">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="b59e3-177">Adicione uma Web Part de editor de script à página e copie e cole o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="b59e3-177">Add a script editor webpart on the page and copy-paste the following code.</span></span>

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

<span data-ttu-id="b59e3-178">Depois disso, você só precisa da URL de um diagrama do Visio com o qual deseja trabalhar.</span><span class="sxs-lookup"><span data-stu-id="b59e3-178">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="b59e3-179">Basta carregar o diagrama do Visio no SharePoint Online e abri-lo no Visio na Web.</span><span class="sxs-lookup"><span data-stu-id="b59e3-179">Just upload the Visio diagram to SharePoint Online and open it in Visio on the web.</span></span> <span data-ttu-id="b59e3-180">A partir daí, abra a caixa de diálogo Inserir e use a URL de integração do exemplo acima.</span><span class="sxs-lookup"><span data-stu-id="b59e3-180">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Copiar a URL do arquivo do Visio da caixa de diálogo Inserir](../images/Visio-embed-url.png)

<span data-ttu-id="b59e3-182">Se você estiver usando o Visio na Web no modo de edição, abra a caixa de diálogo Inserir escolhendo **Arquivo** > **Compartilhar** > **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="b59e3-182">If you are using Visio on the web in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="b59e3-183">Se você estiver usando o Visio na Web no modo de exibição, abra a caixa de diálogo Inserir escolhendo '...' e, em seguida, **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="b59e3-183">If you are using Visio on the web in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="b59e3-184">Referência da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="b59e3-184">Visio JavaScript API reference</span></span>

<span data-ttu-id="b59e3-185">Para saber mais sobre a API JavaScript do Visio, consulte a [Documentação de referência da API JavaScript do Visio](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="b59e3-185">For detailed information about Visio JavaScript API, see the [Visio JavaScript API reference documentation](/javascript/api/visio).</span></span>
