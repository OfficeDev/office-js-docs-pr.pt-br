---
title: Visão geral da API JavaScript do Visio
description: Visão geral da API JavaScript do Visio.
ms.date: 06/03/2020
ms.prod: visio
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 7f706d8f566a747468c4c8d676bd54882bb2a6bf
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076438"
---
# <a name="visio-javascript-api-overview"></a><span data-ttu-id="a840b-103">Visão geral da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="a840b-103">Visio JavaScript API overview</span></span>

<span data-ttu-id="a840b-104">Você pode usar as APIs do Visio JavaScript para incorporar diagramas do Visio em páginas *clássicas* do SharePoint no Microsoft Office SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="a840b-104">You can use the Visio JavaScript APIs to embed Visio diagrams in *classic* SharePoint pages in SharePoint Online.</span></span> <span data-ttu-id="a840b-105">(Este recurso de extensibilidade não é compatível com o Microsoft Office SharePoint Online local ou nas páginas do SharePoint Framework.)</span><span class="sxs-lookup"><span data-stu-id="a840b-105">(This extensibility feature is not supported in on-premise SharePoint or on SharePoint Framework pages.)</span></span>

<span data-ttu-id="a840b-106">Um diagrama integrado do Visio é um diagrama armazenado em uma biblioteca de documentos do SharePoint e exibido em uma página do SharePoint.</span><span class="sxs-lookup"><span data-stu-id="a840b-106">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="a840b-107">Para integrar um diagrama do Visio, exiba-o em um elemento `<iframe>` de HTML.</span><span class="sxs-lookup"><span data-stu-id="a840b-107">To embed a Visio diagram, display it in an HTML `<iframe>` element.</span></span> <span data-ttu-id="a840b-108">Em seguida, você pode usar APIs JavaScript do Visio para trabalhar via programação com o diagrama integrado.</span><span class="sxs-lookup"><span data-stu-id="a840b-108">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Diagrama do Visio em um iframe na página do SharePoint junto com a Web Part do editor de script.](../images/visio-api-block-diagram.png)

<span data-ttu-id="a840b-110">É possível usar as APIs JavaScript do Visio para:</span><span class="sxs-lookup"><span data-stu-id="a840b-110">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="a840b-111">Interagir com os elementos de diagrama do Visio, como páginas e formas.</span><span class="sxs-lookup"><span data-stu-id="a840b-111">Interact with Visio diagram elements like pages and shapes.</span></span>
* <span data-ttu-id="a840b-112">Criar uma marcação visual na tela do diagrama do Visio.</span><span class="sxs-lookup"><span data-stu-id="a840b-112">Create visual markup on the Visio diagram canvas.</span></span>
* <span data-ttu-id="a840b-113">Adicionar manipuladores personalizados para eventos com o mouse no desenho.</span><span class="sxs-lookup"><span data-stu-id="a840b-113">Write custom handlers for mouse events within the drawing.</span></span>
* <span data-ttu-id="a840b-114">Expôr dados de diagrama, como texto da forma, dados da forma e hiperlinks, em sua solução.</span><span class="sxs-lookup"><span data-stu-id="a840b-114">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="a840b-p103">Este artigo descreve como usar as APIs JavaScript do Visio com o Visio na Web para desenvolver suas soluções para o SharePoint Online. Ele apresenta os principais conceitos que são fundamentais para o uso das APIs, como `EmbeddedSession`, `RequestContext` e dos objetos proxy do JavaScript, além dos métodos `sync()`, `Visio.run()`, and `load()`. Os exemplos de código mostram como aplicar esses conceitos.</span><span class="sxs-lookup"><span data-stu-id="a840b-p103">This article describes how to use the Visio JavaScript APIs with Visio on the web to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as `EmbeddedSession`, `RequestContext`, and JavaScript proxy objects, and the `sync()`, `Visio.run()`, and `load()` methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="a840b-118">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="a840b-118">EmbeddedSession</span></span>

<span data-ttu-id="a840b-119">O objeto EmbeddedSession inicia a comunicação entre o quadro do desenvolvedor e o quadro do Visio no navegador.</span><span class="sxs-lookup"><span data-stu-id="a840b-119">The EmbeddedSession object initializes communication between the developer frame and the Visio frame in the browser.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="a840b-120">Visio.run(session, function(context) { batch })</span><span class="sxs-lookup"><span data-stu-id="a840b-120">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="a840b-121">`Visio.run()` executa um script em lote que executa ações no modelo de objeto do Visio.</span><span class="sxs-lookup"><span data-stu-id="a840b-121">`Visio.run()` executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="a840b-122">Os comandos em lotes incluem definições de objetos proxy JavaScript locais e métodos `sync()` que sincronizam o estado entre objetos locais e do Visio, e a resolução de promessa.</span><span class="sxs-lookup"><span data-stu-id="a840b-122">The batch commands include definitions of local JavaScript proxy objects and `sync()` methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="a840b-123">A vantagem do envio de solicitações em lotes com o método `Visio.run()` é que, quando a promessa é resolvida, todos os objetos de página controlados que foram alocados durante a execução são automaticamente liberados.</span><span class="sxs-lookup"><span data-stu-id="a840b-123">The advantage of batching requests in `Visio.run()` is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="a840b-124">O método run recebe a sessão e o objeto RequestContext e retorna uma promessa (normalmente, apenas o resultado de `context.sync()`).</span><span class="sxs-lookup"><span data-stu-id="a840b-124">The run method takes in session and RequestContext object and returns a promise (typically, just the result of `context.sync()`).</span></span> <span data-ttu-id="a840b-125">É possível executar a operação em lote fora do `Visio.run()`.</span><span class="sxs-lookup"><span data-stu-id="a840b-125">It is possible to run the batch operation outside of the `Visio.run()`.</span></span> <span data-ttu-id="a840b-126">No entanto, todas as referências aos objetos de página devem ser rastreadas e gerenciadas manualmente nesse cenário.</span><span class="sxs-lookup"><span data-stu-id="a840b-126">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="a840b-127">RequestContext</span><span class="sxs-lookup"><span data-stu-id="a840b-127">RequestContext</span></span>

<span data-ttu-id="a840b-p106">O objeto RequestContext facilita as solicitações para o aplicativo Visio. Como o quadro do desenvolvedor e o cliente Web do Visio são executados em dois iframes diferentes, o objeto RequestContext (contexto no próximo exemplo) é necessário para obter acesso ao Visio e a objetos relacionados, como páginas e formas, do quadro do desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="a840b-p106">The RequestContext object facilitates requests to the Visio application. Because the developer frame and the Visio web client run in two different iframes, the RequestContext object (context in next example) is required to get access to Visio and related objects such as pages and shapes, from the developer frame.</span></span>

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

## <a name="proxy-objects"></a><span data-ttu-id="a840b-130">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="a840b-130">Proxy objects</span></span>

<span data-ttu-id="a840b-131">Os objetos JavaScript do Visio declarados e usados em uma sessão incorporada são objetos proxy para os objetos reais em um documento do Visio.</span><span class="sxs-lookup"><span data-stu-id="a840b-131">The Visio JavaScript objects declared and used in an embedded session are proxy objects for the real objects in a Visio document.</span></span> <span data-ttu-id="a840b-132">Todas as ações executadas em objetos proxy não são percebidas no Visio, e o estado do documento do Visio não é percebido em objetos proxy, até que o estado do documento tenha sido sincronizado.</span><span class="sxs-lookup"><span data-stu-id="a840b-132">All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized.</span></span> <span data-ttu-id="a840b-133">O estado do documento é sincronizado quando `context.sync()` é executado.</span><span class="sxs-lookup"><span data-stu-id="a840b-133">The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="a840b-134">Por exemplo, o objeto JavaScript local getActivePage é declarado para fazer referência à página selecionada.</span><span class="sxs-lookup"><span data-stu-id="a840b-134">For example, the local JavaScript object getActivePage is declared to reference the selected page.</span></span> <span data-ttu-id="a840b-135">Você pode usá-lo para colocar a configuração das respectivas propriedades em fila e para invocar métodos.</span><span class="sxs-lookup"><span data-stu-id="a840b-135">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="a840b-136">As ações nesses objetos não são realizadas até que o método `sync()` seja executado.</span><span class="sxs-lookup"><span data-stu-id="a840b-136">The actions on such objects are not realized until the `sync()` method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="a840b-137">sync()</span><span class="sxs-lookup"><span data-stu-id="a840b-137">sync()</span></span>

<span data-ttu-id="a840b-138">O método `sync()` sincroniza o estado entre objetos proxy JavaScript e objetos reais no Visio, com a execução de instruções enfileiradas no contexto e com a recuperação de propriedades de objetos carregados do Office para uso no código.</span><span class="sxs-lookup"><span data-stu-id="a840b-138">The `sync()` method synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="a840b-139">Este método retorna uma promessa, que é resolvida quando o sistema conclui a sincronização.</span><span class="sxs-lookup"><span data-stu-id="a840b-139">This method returns a promise, which is resolved when synchronization is complete.</span></span>

## <a name="load"></a><span data-ttu-id="a840b-140">load()</span><span class="sxs-lookup"><span data-stu-id="a840b-140">load()</span></span>

<span data-ttu-id="a840b-141">O método `load()` é usado para preencher os objetos proxy criados na camada JavaScript do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a840b-141">The `load()` method is used to fill in the proxy objects created in the JavaScript layer.</span></span> <span data-ttu-id="a840b-142">Ao tentar recuperar um objeto, como um documento, um objeto proxy local é criado inicialmente na camada JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a840b-142">When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer.</span></span> <span data-ttu-id="a840b-143">Você pode usar esse objeto para colocar a configuração das respectivas propriedades em fila e para invocar métodos.</span><span class="sxs-lookup"><span data-stu-id="a840b-143">Such an object can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="a840b-144">No entanto, para ler as propriedades ou relações do objeto, os métodos `load()`e `sync()` precisam ser chamados primeiro.</span><span class="sxs-lookup"><span data-stu-id="a840b-144">However, for reading object properties or relations, the `load()` and `sync()` methods need to be invoked first.</span></span> <span data-ttu-id="a840b-145">O método load() leva nas propriedades e relações que precisam ser carregadas quando o método `sync()` é chamado.</span><span class="sxs-lookup"><span data-stu-id="a840b-145">The load() method takes in the properties and relations that need to be loaded when the `sync()` method is called.</span></span>

<span data-ttu-id="a840b-146">A seguir, é mostrada a sintaxe do método `load()`.</span><span class="sxs-lookup"><span data-stu-id="a840b-146">The following shows the syntax for the `load()` method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="a840b-147">**properties** é a lista de nomes de propriedades a carregar, especificados como cadeias de caracteres delimitadas por vírgulas ou por uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="a840b-147">**properties** is the list of property names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="a840b-148">Veja os métodos `.load()` em cada objeto para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="a840b-148">See `.load()` methods under each object for details.</span></span>

2. <span data-ttu-id="a840b-p112">**loadOption** especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Consulte as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="a840b-p112">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="a840b-151">Exemplo: imprimir todo o texto de formas na página ativa</span><span class="sxs-lookup"><span data-stu-id="a840b-151">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="a840b-152">O exemplo a seguir mostra como imprimir o valor de texto de forma de um objeto de formas de matriz.</span><span class="sxs-lookup"><span data-stu-id="a840b-152">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="a840b-153">O método `Visio.run()` contém um lote de instruções.</span><span class="sxs-lookup"><span data-stu-id="a840b-153">The `Visio.run()` method contains a batch of instructions.</span></span> <span data-ttu-id="a840b-154">Como parte deste lote, o sistema cria um objeto proxy que faz referência a formas no documento ativo.</span><span class="sxs-lookup"><span data-stu-id="a840b-154">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="a840b-155">Todos esses comandos são enfileirados e executados quando `context.sync()` é chamado.</span><span class="sxs-lookup"><span data-stu-id="a840b-155">All these commands are queued and run when `context.sync()` is called.</span></span> <span data-ttu-id="a840b-156">O método `sync()` retorna uma promessa que pode ser usada para encadeá-lo com outras operações.</span><span class="sxs-lookup"><span data-stu-id="a840b-156">The `sync()` method returns a promise that can be used to chain it with other operations.</span></span>

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

## <a name="error-messages"></a><span data-ttu-id="a840b-157">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="a840b-157">Error messages</span></span>

<span data-ttu-id="a840b-p115">O sistema retorna erros usando um objeto Error composto por um código e uma mensagem. A tabela a seguir fornece uma lista de possíveis condições de erro que podem ocorrer.</span><span class="sxs-lookup"><span data-stu-id="a840b-p115">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="a840b-160">error.code</span><span class="sxs-lookup"><span data-stu-id="a840b-160">error.code</span></span>            | <span data-ttu-id="a840b-161">error.message</span><span class="sxs-lookup"><span data-stu-id="a840b-161">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="a840b-162">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="a840b-162">InvalidArgument</span></span>       | <span data-ttu-id="a840b-163">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="a840b-163">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="a840b-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="a840b-164">GeneralException</span></span>      | <span data-ttu-id="a840b-165">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="a840b-165">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="a840b-166">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="a840b-166">NotImplemented</span></span>        | <span data-ttu-id="a840b-167">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="a840b-167">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="a840b-168">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="a840b-168">UnsupportedOperation</span></span>  | <span data-ttu-id="a840b-169">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="a840b-169">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="a840b-170">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="a840b-170">AccessDenied</span></span>          | <span data-ttu-id="a840b-171">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="a840b-171">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="a840b-172">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="a840b-172">ItemNotFound</span></span>          | <span data-ttu-id="a840b-173">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="a840b-173">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="a840b-174">Introdução</span><span class="sxs-lookup"><span data-stu-id="a840b-174">Get started</span></span>

<span data-ttu-id="a840b-175">Você pode usar o exemplo nesta seção para começar.</span><span class="sxs-lookup"><span data-stu-id="a840b-175">You can use the example in this section to get started.</span></span> <span data-ttu-id="a840b-176">Este exemplo mostra como exibir o texto da forma selecionada em um diagrama do Visio via programação.</span><span class="sxs-lookup"><span data-stu-id="a840b-176">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="a840b-177">Para começar, crie uma página clássica no SharePoint Online ou edite uma página existente.</span><span class="sxs-lookup"><span data-stu-id="a840b-177">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="a840b-178">Adicione uma Web Part de editor de script à página e copie e cole o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a840b-178">Add a script editor webpart on the page and copy-paste the following code.</span></span>

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

<span data-ttu-id="a840b-179">Depois disso, você só precisa da URL de um diagrama do Visio com o qual deseja trabalhar.</span><span class="sxs-lookup"><span data-stu-id="a840b-179">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="a840b-180">Basta carregar o diagrama do Visio no SharePoint Online e abri-lo no Visio na Web.</span><span class="sxs-lookup"><span data-stu-id="a840b-180">Just upload the Visio diagram to SharePoint Online and open it in Visio on the web.</span></span> <span data-ttu-id="a840b-181">A partir daí, abra a caixa de diálogo Inserir e use a URL de integração do exemplo acima.</span><span class="sxs-lookup"><span data-stu-id="a840b-181">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Copiar a URL do arquivo do Visio da caixa de diálogo Inserir.](../images/Visio-embed-url.png)

<span data-ttu-id="a840b-183">Se você estiver usando o Visio na Web no modo de edição, abra a caixa de diálogo Inserir escolhendo **Arquivo** > **Compartilhar** > **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="a840b-183">If you are using Visio on the web in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="a840b-184">Se você estiver usando o Visio na Web no modo de exibição, abra a caixa de diálogo Inserir escolhendo '...' e, em seguida, **Inserir**.</span><span class="sxs-lookup"><span data-stu-id="a840b-184">If you are using Visio on the web in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="a840b-185">Referência da API JavaScript do Visio</span><span class="sxs-lookup"><span data-stu-id="a840b-185">Visio JavaScript API reference</span></span>

<span data-ttu-id="a840b-186">Para saber mais sobre a API JavaScript do Visio, consulte a [Documentação de referência da API JavaScript do Visio](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="a840b-186">For detailed information about Visio JavaScript API, see the [Visio JavaScript API reference documentation](/javascript/api/visio).</span></span>
