---
title: Programação assíncrona em Suplementos do Office
description: Saiba como a biblioteca JavaScript do Office usa a programação assíncrona em suplementos do Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 0ffeb6a0745c59df5130cb742ee9e51752396d42
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131938"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="65f0f-103">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="65f0f-103">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="65f0f-104">Por que a API de Suplementos do Office usa a programação assíncrona?</span><span class="sxs-lookup"><span data-stu-id="65f0f-104">Why does the Office Add-ins API use asynchronous programming?</span></span> <span data-ttu-id="65f0f-105">Como o JavaScript é uma linguagem de thread único, se o script invocar um processo síncrono demorado, todas as execuções subsequentes do script serão bloqueadas até que o processo seja concluído.</span><span class="sxs-lookup"><span data-stu-id="65f0f-105">Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes.</span></span> <span data-ttu-id="65f0f-106">Como determinadas operações com clientes Web do Office (mas também clientes avançados) podem bloquear a execução se forem executadas de forma síncrona, a maioria das APIs JavaScript do Office é projetada para ser executada de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="65f0f-106">Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously.</span></span> <span data-ttu-id="65f0f-107">Isso garante que os suplementos do Office sejam responsivos e rápidos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-107">This makes sure that Office Add-ins are responsive and fast.</span></span> <span data-ttu-id="65f0f-108">Em geral, isso também requer que você escreva funções de retorno de chamada ao trabalhar com esses métodos assíncronos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-108">It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="65f0f-109">Os nomes de todos os métodos assíncronos na API terminam com "Async", como `Document.getSelectedDataAsync` os `Binding.getDataAsync` métodos, ou `Item.loadCustomPropertiesAsync` .</span><span class="sxs-lookup"><span data-stu-id="65f0f-109">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods.</span></span> <span data-ttu-id="65f0f-110">Quando um método "Async" é chamado, ele é executado imediatamente e qualquer execução subsequente do script poderá continuar.</span><span class="sxs-lookup"><span data-stu-id="65f0f-110">When an "Async" method is called, it executes immediately and any subsequent script execution can continue.</span></span> <span data-ttu-id="65f0f-111">A função de retorno de chamada opcional que você passar para um método de "Async" é executada assim que os dados ou a operação solicitada está pronta.</span><span class="sxs-lookup"><span data-stu-id="65f0f-111">The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready.</span></span> <span data-ttu-id="65f0f-112">Isso geralmente ocorre imediatamente, mas pode haver um pequeno atraso antes de retornar.</span><span class="sxs-lookup"><span data-stu-id="65f0f-112">This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="65f0f-113">O diagrama a seguir mostra o fluxo de execução de uma chamada para um método "Async" que lê os dados que o usuário selecionou em um documento aberto no Word baseado no servidor ou no Excel.</span><span class="sxs-lookup"><span data-stu-id="65f0f-113">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel.</span></span> <span data-ttu-id="65f0f-114">No ponto em que a chamada "Async" é feita, o thread de execução do JavaScript é livre para executar qualquer processamento adicional no lado do cliente (embora nenhum seja mostrado no diagrama).</span><span class="sxs-lookup"><span data-stu-id="65f0f-114">At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram).</span></span> <span data-ttu-id="65f0f-115">Quando o método "Async" retorna, o retorno de chamada continua a execução no thread e o suplemento pode acessar os dados, fazer algo com ele e exibir o resultado.</span><span class="sxs-lookup"><span data-stu-id="65f0f-115">When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result.</span></span> <span data-ttu-id="65f0f-116">O mesmo padrão de execução assíncrona é executado ao trabalhar com os aplicativos cliente avançados do Office, como o Word 2013 ou o Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="65f0f-116">The same asynchronous execution pattern holds when working with the Office rich client applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="65f0f-117">*Figura 1. Fluxo de execução da programação assíncrona*</span><span class="sxs-lookup"><span data-stu-id="65f0f-117">*Figure 1. Asynchronous programming execution flow*</span></span>

![Diagrama mostrando a interação de execução do comando ao longo do tempo com o usuário, a página de suplemento e o servidor de aplicativos Web que hospeda o suplemento](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="65f0f-p104">O suporte a este design assíncrono em clientes Web e avançados faz parte das metas de design "gravar plataforma cruzada já executada" do modelo de desenvolvimento de Suplementos do Office. Por exemplo, você pode criar um suplemento do painel de tarefas ou conteúdo com uma única base de código que será executada no Excel 2013 e Excel Online.</span><span class="sxs-lookup"><span data-stu-id="65f0f-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="65f0f-121">Gravar a função de retorno de chamada para um método "Async"</span><span class="sxs-lookup"><span data-stu-id="65f0f-121">Writing the callback function for an "Async" method</span></span>

<span data-ttu-id="65f0f-122">A função de retorno de chamada passada como o argumento de _retorno de chamada_ para um método "Async" deve declarar um único parâmetro que o tempo de execução do suplemento usará para fornecer acesso a um objeto [AsyncResult](/javascript/api/office/office.asyncresult) quando a função de retorno de chamada for executada.</span><span class="sxs-lookup"><span data-stu-id="65f0f-122">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes.</span></span> <span data-ttu-id="65f0f-123">Você pode gravar:</span><span class="sxs-lookup"><span data-stu-id="65f0f-123">You can write:</span></span>

- <span data-ttu-id="65f0f-124">Uma função anônima que deve ser gravada e transmitida diretamente em linha com a chamada para o método "Async" como o parâmetro _callback_ do método "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-124">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="65f0f-125">Uma função nomeada, passando o nome dessa função como o parâmetro _callback_ de um método "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-125">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="65f0f-p106">Uma função anônima é útil se você só for usar seu código uma vez – porque ele não possui um nome, você não pode referenciá-la em outra parte do seu código. Uma função nomeada é útil se você quiser reutilizar a função retorno de chamada para mais de um método "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>

### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="65f0f-128">Gravar uma função de retorno de chamada anônima</span><span class="sxs-lookup"><span data-stu-id="65f0f-128">Writing an anonymous callback function</span></span>

<span data-ttu-id="65f0f-129">A seguinte função de retorno de chamada anônima declara um único parâmetro chamado `result` que recupera os dados da propriedade [AsyncResult. Value](/javascript/api/office/office.asyncresult#value) quando o retorno de chamada retorna.</span><span class="sxs-lookup"><span data-stu-id="65f0f-129">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="65f0f-130">O exemplo a seguir mostra como passar essa função de retorno de chamada anônima em linha no contexto de uma chamada de método "Async" completa para o `Document.getSelectedDataAsync` método.</span><span class="sxs-lookup"><span data-stu-id="65f0f-130">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>

- <span data-ttu-id="65f0f-131">O primeiro argumento _coercionType_ , `Office.CoercionType.Text` , especifica para retornar os dados selecionados como uma cadeia de texto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-131">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="65f0f-132">O segundo argumento de _retorno de chamada_ é a função anônima passada na linha para o método.</span><span class="sxs-lookup"><span data-stu-id="65f0f-132">The second _callback_ argument is the anonymous function passed in-line to the method.</span></span> <span data-ttu-id="65f0f-133">Quando a função é executada, ela usa o parâmetro _Result_ para acessar a `value` Propriedade do `AsyncResult` objeto para exibir os dados selecionados pelo usuário no documento.</span><span class="sxs-lookup"><span data-stu-id="65f0f-133">When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="65f0f-134">Você também pode usar o parâmetro de sua função de retorno de chamada para acessar outras propriedades do `AsyncResult` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-134">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object.</span></span> <span data-ttu-id="65f0f-135">Use a propriedade [AsyncResult.status](/javascript/api/office/office.asyncresult#status) para determinar se a chamada teve êxito ou falhou.</span><span class="sxs-lookup"><span data-stu-id="65f0f-135">Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed.</span></span> <span data-ttu-id="65f0f-136">Se sua chamada falhar, você pode usar a propriedade [AsyncResult.error](/javascript/api/office/office.asyncresult#error) para acessar um objeto [Error](/javascript/api/office/office.error) para informações sobre o erro.</span><span class="sxs-lookup"><span data-stu-id="65f0f-136">If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="65f0f-137">Para obter mais informações sobre como usar o `getSelectedDataAsync` método, confira [ler e gravar dados na seleção ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="65f0f-137">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 

### <a name="writing-a-named-callback-function"></a><span data-ttu-id="65f0f-138">Gravar uma função de retorno de chamada nomeada</span><span class="sxs-lookup"><span data-stu-id="65f0f-138">Writing a named callback function</span></span>

<span data-ttu-id="65f0f-139">Como alternativa, você pode escrever uma função nomeada e passar seu nome para o parâmetro _callback_ de um método "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-139">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method.</span></span> <span data-ttu-id="65f0f-140">Por exemplo, o exemplo anterior pode ser reescrito para transmitir uma função chamada `writeDataCallback` como o parâmetro _callback_ assim.</span><span class="sxs-lookup"><span data-stu-id="65f0f-140">For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="65f0f-141">Diferenças entre o que é retornado para a propriedade AsyncResult.value</span><span class="sxs-lookup"><span data-stu-id="65f0f-141">Differences in what's returned to the AsyncResult.value property</span></span>

<span data-ttu-id="65f0f-142">As `asyncContext` `status` Propriedades,, e `error` do `AsyncResult` objeto retornam os mesmos tipos de informações para a função de retorno de chamada passada para todos os métodos "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-142">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods.</span></span> <span data-ttu-id="65f0f-143">No entanto, o que é retornado para a `AsyncResult.value` Propriedade varia dependendo da funcionalidade do método "Async".</span><span class="sxs-lookup"><span data-stu-id="65f0f-143">However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="65f0f-144">Por exemplo, os `addHandlerAsync` métodos (dos objetos [Binding](/javascript/api/office/office.binding), [CustomXMLPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings)e [Settings](/javascript/api/office/office.settings) ) são usados para adicionar funções de manipulador de eventos aos itens representados por esses objetos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-144">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects.</span></span> <span data-ttu-id="65f0f-145">Você pode acessar a `AsyncResult.value` propriedade da função de retorno de chamada passada para qualquer um dos `addHandlerAsync` métodos, mas como nenhum dado ou objeto está sendo acessado quando você adiciona um manipulador de eventos, a `value` propriedade sempre retorna **undefined** se você tentar acessá-lo.</span><span class="sxs-lookup"><span data-stu-id="65f0f-145">You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="65f0f-146">Por outro lado, se você chamar o `Document.getSelectedDataAsync` método, ele retornará os dados que o usuário selecionou no documento para a `AsyncResult.value` propriedade no retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="65f0f-146">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback.</span></span> <span data-ttu-id="65f0f-147">Ou, se você chamar o método [bindings. getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) , ele retornará uma matriz de todos os `Binding` objetos no documento.</span><span class="sxs-lookup"><span data-stu-id="65f0f-147">Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document.</span></span> <span data-ttu-id="65f0f-148">E, se você chamar o método [bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) , ele retornará um único `Binding` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-148">And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="65f0f-149">Para obter uma descrição do que é retornado para a `AsyncResult.value` propriedade de um `Async` método, consulte a seção "valor de retorno de chamada" do tópico de referência desse método.</span><span class="sxs-lookup"><span data-stu-id="65f0f-149">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic.</span></span> <span data-ttu-id="65f0f-150">Para obter um resumo de todos os objetos que fornecem `Async` métodos, consulte a tabela na parte inferior do tópico do objeto [AsyncResult](/javascript/api/office/office.asyncresult) .</span><span class="sxs-lookup"><span data-stu-id="65f0f-150">For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>

## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="65f0f-151">Padrões de programação assíncrona</span><span class="sxs-lookup"><span data-stu-id="65f0f-151">Asynchronous programming patterns</span></span>

<span data-ttu-id="65f0f-152">A API JavaScript do Office oferece suporte a dois tipos de padrões de programação assíncrono:</span><span class="sxs-lookup"><span data-stu-id="65f0f-152">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>

- <span data-ttu-id="65f0f-153">Usando retornos de chamada aninhados</span><span class="sxs-lookup"><span data-stu-id="65f0f-153">Using nested callbacks</span></span>
- <span data-ttu-id="65f0f-154">Usando o padrão de promessas</span><span class="sxs-lookup"><span data-stu-id="65f0f-154">Using the promises pattern</span></span>

<span data-ttu-id="65f0f-p114">A programação assíncrona com funções de retorno de chamada frequentemente exigem que você aninhe o resultado retornado de um retorno de chamada dentro de dois ou mais retornos de chamada. Se você precisar fazer isso, é possível usar retornos de chamada aninhados de todos os métodos "Async" da API.</span><span class="sxs-lookup"><span data-stu-id="65f0f-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="65f0f-157">Usar retornos de chamada aninhados é um padrão de programação familiar para a maioria dos desenvolvedores de JavaScript, mas códigos com retornos de chamada profundamente aninhados podem ser difíceis de ler e entender.</span><span class="sxs-lookup"><span data-stu-id="65f0f-157">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand.</span></span> <span data-ttu-id="65f0f-158">Como alternativa para retornos de chamada aninhados, a API JavaScript do Office também oferece suporte a uma implementação do padrão de promessas.</span><span class="sxs-lookup"><span data-stu-id="65f0f-158">As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern.</span></span>

> [!NOTE]
> <span data-ttu-id="65f0f-159">Na versão atual da API JavaScript do Office, o suporte *interno* para o padrão de promessas só funciona com o código para [associações em documentos do Word e planilhas do Excel](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="65f0f-159">In the current version of the Office JavaScript API, *built-in* support for the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span> <span data-ttu-id="65f0f-160">No entanto, você pode quebrar outras funções que têm retornos de chamada dentro de sua própria função de retorno de promessa personalizada.</span><span class="sxs-lookup"><span data-stu-id="65f0f-160">However, you can wrap other functions that have callbacks inside your own custom Promise-returning function.</span></span> <span data-ttu-id="65f0f-161">Para obter mais informações, consulte [quebra de APIs comuns em funções que retornam a promessa](#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="65f0f-161">For more information, see [Wrap Common APIs in Promise-returning functions](#wrap-common-apis-in-promise-returning-functions).</span></span>

### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="65f0f-162">Programação assíncrona usando funções aninhadas de retorno de chamada</span><span class="sxs-lookup"><span data-stu-id="65f0f-162">Asynchronous programming using nested callback functions</span></span>

<span data-ttu-id="65f0f-p117">Frequentemente, você precisa executar duas ou mais operações assíncronas para concluir uma tarefa. Para fazer isso, você pode aninhar uma chamada "Async" dentro de outra.</span><span class="sxs-lookup"><span data-stu-id="65f0f-p117">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="65f0f-165">O exemplo de código a seguir aninha duas ou mais chamadas assíncronas.</span><span class="sxs-lookup"><span data-stu-id="65f0f-165">The following code example nests two asynchronous calls.</span></span>

- <span data-ttu-id="65f0f-166">Primeiro, o método [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) é chamado para acessar uma associação no documento chamado "MyBinding".</span><span class="sxs-lookup"><span data-stu-id="65f0f-166">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding".</span></span> <span data-ttu-id="65f0f-167">O `AsyncResult` objeto retornado para o `result` parâmetro desse retorno de chamada fornece acesso ao objeto Binding especificado a partir da `AsyncResult.value` propriedade.</span><span class="sxs-lookup"><span data-stu-id="65f0f-167">The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>
- <span data-ttu-id="65f0f-168">Em seguida, o objeto de vinculação acessado do primeiro `result` parâmetro é usado para chamar o método [Binding. getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="65f0f-168">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>
- <span data-ttu-id="65f0f-169">Por fim, o `result2` parâmetro do retorno de chamada passado para o `Binding.getDataAsync` método é usado para exibir os dados na associação.</span><span class="sxs-lookup"><span data-stu-id="65f0f-169">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>

```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="65f0f-170">Este padrão de retorno de chamada aninhado básico pode ser usado para todos os métodos assíncronos na API JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="65f0f-170">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="65f0f-171">As seções a seguir mostram como usar funções anônimas ou nomeadas para retornos de chamada aninhados em métodos assíncronos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-171">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>

#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="65f0f-172">Usando funções anônimas para retornos de chamada aninhados</span><span class="sxs-lookup"><span data-stu-id="65f0f-172">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="65f0f-173">No exemplo a seguir, duas funções anônimas são declaradas embutidas e passadas para os `getByIdAsync` `getDataAsync` métodos e como callbacks aninhados.</span><span class="sxs-lookup"><span data-stu-id="65f0f-173">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks.</span></span> <span data-ttu-id="65f0f-174">Como as funções são simples e embutidas, a intenção da implementação fica imediatamente clara.</span><span class="sxs-lookup"><span data-stu-id="65f0f-174">Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="65f0f-175">Usando funções nomeadas para retornos de chamada aninhados</span><span class="sxs-lookup"><span data-stu-id="65f0f-175">Using named functions for nested callbacks</span></span>

<span data-ttu-id="65f0f-176">Em implementações complexas, pode ser útil usar funções nomeadas para facilitar a leitura, manutenção e reutilização do seu código.</span><span class="sxs-lookup"><span data-stu-id="65f0f-176">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse.</span></span> <span data-ttu-id="65f0f-177">No exemplo a seguir, as duas funções anônimas do exemplo na seção anterior foram reescritas como funções chamadas `deleteAllData` e `showResult` .</span><span class="sxs-lookup"><span data-stu-id="65f0f-177">In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`.</span></span> <span data-ttu-id="65f0f-178">Essas funções nomeadas são então passadas para os `getByIdAsync` `deleteAllDataValuesAsync` métodos e como callbacks por nome.</span><span class="sxs-lookup"><span data-stu-id="65f0f-178">These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>

```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="65f0f-179">Programação assíncrona usando o padrão de promessas para acessar dados em associações</span><span class="sxs-lookup"><span data-stu-id="65f0f-179">Asynchronous programming using the promises pattern to access data in bindings</span></span>

<span data-ttu-id="65f0f-p121">Em vez de transmitir a função de retorno de chamada e aguardar até que a função retorne antes da continuação da execução, o padrão de programação de promessas retorna imediatamente retorna um objeto de promessa que representa o resultado desejado. No entanto, ao contrário da verdadeira programação síncrona, nos bastidores o cumprimento do resultado prometido é, na verdade, adiado até que o ambiente de tempo de execução dos Suplementos do Office possa concluir a solicitação. Um manipulador _onError_ é fornecido para atender a situações em que a solicitação não pode ser cumprida.</span><span class="sxs-lookup"><span data-stu-id="65f0f-p121">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>

<span data-ttu-id="65f0f-183">A API JavaScript do Office fornece o método [Office. Select](/javascript/api/office#office-select-expression--callback-) para dar suporte ao padrão de promessas para trabalhar com objetos Binding existentes.</span><span class="sxs-lookup"><span data-stu-id="65f0f-183">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects.</span></span> <span data-ttu-id="65f0f-184">O objeto Promise retornado para o `Office.select` método suporta apenas os quatro métodos que você pode acessar diretamente a partir do objeto [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)e [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="65f0f-184">The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>

<span data-ttu-id="65f0f-185">O padrão de promessas para funcionar com associações tem este formato:</span><span class="sxs-lookup"><span data-stu-id="65f0f-185">The promises pattern for working with bindings takes this form:</span></span>

<span data-ttu-id="65f0f-186">**Office. Select (**_selectorr_, _OnError_**).** _BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="65f0f-186">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="65f0f-187">O parâmetro _selectoré_ assume o formato `"bindings#bindingId"` , em que _BindingId_ é o nome ( `id` ) de uma associação que você criou anteriormente no documento ou planilha (usando um dos métodos "addfrom" do `Bindings` conjunto: `addFromNamedItemAsync` , `addFromPromptAsync` , ou `addFromSelectionAsync` ).</span><span class="sxs-lookup"><span data-stu-id="65f0f-187">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`).</span></span> <span data-ttu-id="65f0f-188">Por exemplo, a expressão de seletor `bindings#cities` especifica que você deseja acessar a associação com uma **ID** de "cidades".</span><span class="sxs-lookup"><span data-stu-id="65f0f-188">For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="65f0f-189">O parâmetro _OnError_ é uma função de tratamento de erros que usa um único parâmetro do tipo `AsyncResult` que pode ser usado para acessar um `Error` objeto, se o `select` método falhar ao acessar a associação especificada.</span><span class="sxs-lookup"><span data-stu-id="65f0f-189">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding.</span></span> <span data-ttu-id="65f0f-190">O exemplo a seguir mostra uma função de manipulador de erro básica que pode ser transmitida para o parâmetro _onError_.</span><span class="sxs-lookup"><span data-stu-id="65f0f-190">The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>

```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="65f0f-191">Substitua o espaço reservado _BindingObjectAsyncMethod_ por uma chamada para qualquer um dos quatro `Binding` métodos de objeto suportados pelo objeto Promise: `getDataAsync` , `setDataAsync` , `addHandlerAsync` ou `removeHandlerAsync` .</span><span class="sxs-lookup"><span data-stu-id="65f0f-191">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`.</span></span> <span data-ttu-id="65f0f-192">As chamadas para esses métodos não oferecem suporte a promessas adicionais.</span><span class="sxs-lookup"><span data-stu-id="65f0f-192">Calls to these methods don't support additional promises.</span></span> <span data-ttu-id="65f0f-193">Você deve chamá-los usando o [padrão de função de retorno de chamada aninhado](#asynchronous-programming-using-nested-callback-functions).</span><span class="sxs-lookup"><span data-stu-id="65f0f-193">You must call them using the [nested callback function pattern](#asynchronous-programming-using-nested-callback-functions).</span></span>

<span data-ttu-id="65f0f-194">Depois que uma `Binding` promessa de objeto é atendida, ela pode ser reutilizada na chamada do método encadeado como se fosse uma associação (o tempo de execução do suplemento não tentará executar a promessa de forma assíncrona).</span><span class="sxs-lookup"><span data-stu-id="65f0f-194">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise).</span></span> <span data-ttu-id="65f0f-195">Se a `Binding` promessa do objeto não puder ser atendida, o tempo de execução do suplemento tentará novamente acessar o objeto Binding na próxima vez que um de seus métodos assíncronos for chamado.</span><span class="sxs-lookup"><span data-stu-id="65f0f-195">If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="65f0f-196">O exemplo de código a seguir usa o `select` método para recuperar uma associação com o `id` " `cities` " do `Bindings` conjunto e, em seguida, chama o método [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) para adicionar um manipulador de eventos para o evento [DataChanged](/javascript/api/office/office.bindingdatachangedeventargs) da associação.</span><span class="sxs-lookup"><span data-stu-id="65f0f-196">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> <span data-ttu-id="65f0f-197">A `Binding` promessa do objeto retornada pelo `Office.select` método fornece acesso apenas aos quatro métodos do `Binding` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-197">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object.</span></span> <span data-ttu-id="65f0f-198">Se você precisar acessar qualquer um dos outros membros do `Binding` objeto, em vez disso, deverá usar a `Document.bindings` propriedade e `Bindings.getByIdAsync` ou os `Bindings.getAllAsync` métodos para recuperar o `Binding` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-198">If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object.</span></span> <span data-ttu-id="65f0f-199">Por exemplo, se você precisar acessar qualquer uma das `Binding` Propriedades do objeto (as `document` Propriedades, `id` ou, `type` ) ou precisar acessar as propriedades dos objetos [matrixbinding](/javascript/api/office/office.matrixbinding) ou [TableBinding](/javascript/api/office/office.tablebinding) , você deve usar os `getByIdAsync` `getAllAsync` métodos ou para recuperar um `Binding` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-199">For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>

## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="65f0f-200">Transmitir parâmetros opcionais para métodos assíncronos</span><span class="sxs-lookup"><span data-stu-id="65f0f-200">Passing optional parameters to asynchronous methods</span></span>

<span data-ttu-id="65f0f-201">A sintaxe comum para todos os métodos "Async" segue este padrão:</span><span class="sxs-lookup"><span data-stu-id="65f0f-201">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="65f0f-202">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="65f0f-202">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="65f0f-p128">Todos os métodos assíncronos dão suporte parâmetros opcionais, que são passados como um objeto JSON (JavaScript Object Notation) contendo um ou mais parâmetros opcionais. O objeto JSON que contém os parâmetros opcionais é uma coleção desordenada de pares de valores e chaves com o caractere ":" separando os valores e as chaves. Cada par do objeto é separado por vírgula e o conjunto completo de pares é incluído entre chaves. A chave é o nome do parâmetro e o valor é o valor a ser transmitido para esse parâmetro.</span><span class="sxs-lookup"><span data-stu-id="65f0f-p128">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="65f0f-207">Você pode criar o objeto JSON que contém parâmetros opcionais embutidos ou criando um `options` objeto e transmitindo-o como o parâmetro _Options_ .</span><span class="sxs-lookup"><span data-stu-id="65f0f-207">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>

### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="65f0f-208">Transmitir parâmetros opcionais embutidos</span><span class="sxs-lookup"><span data-stu-id="65f0f-208">Passing optional parameters inline</span></span>

<span data-ttu-id="65f0f-209">Por exemplo, a sintaxe para chamar o método [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) com parâmetros opcionais embutidos tem esta aparência:</span><span class="sxs-lookup"><span data-stu-id="65f0f-209">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="65f0f-210">Neste formulário da sintaxe de chamada, os dois parâmetros opcionais, _coercionType_ e _asyncContext_, são definidos como um objeto JSON embutido entre chaves.</span><span class="sxs-lookup"><span data-stu-id="65f0f-210">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="65f0f-211">O exemplo a seguir mostra como chamar o `Document.setSelectedDataAsync` método especificando parâmetros opcionais embutidos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-211">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>

```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> <span data-ttu-id="65f0f-212">É possível especificar parâmetros opcionais em qualquer ordem no objeto JSON desde que seus nomes sejam especificados corretamente.</span><span class="sxs-lookup"><span data-stu-id="65f0f-212">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>

### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="65f0f-213">Transmitir parâmetros opcionais em um objeto de opções</span><span class="sxs-lookup"><span data-stu-id="65f0f-213">Passing optional parameters in an options object</span></span>

<span data-ttu-id="65f0f-214">Como alternativa, você pode criar um objeto nomeado `options` que especifica os parâmetros opcionais separadamente da chamada do método e, em seguida, passar o `options` objeto como o argumento _Options_ .</span><span class="sxs-lookup"><span data-stu-id="65f0f-214">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="65f0f-215">O exemplo a seguir mostra uma maneira de criar o `options` objeto, onde `parameter1` , `value1` e assim por diante, são espaços reservados para os valores e nomes de parâmetro reais.</span><span class="sxs-lookup"><span data-stu-id="65f0f-215">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="65f0f-216">Que é semelhante ao exemplo a seguir quando usado para especificar os parâmetros [ValueFormat](/javascript/api/office/office.valueformat) e [FilterType](/javascript/api/office/office.filtertype).</span><span class="sxs-lookup"><span data-stu-id="65f0f-216">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="65f0f-217">Esta é outra maneira de criar o `options` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-217">Here's another way of creating the `options` object.</span></span>

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="65f0f-218">Que é semelhante ao exemplo a seguir quando usado para especificar `ValueFormat` os `FilterType` parâmetros e:</span><span class="sxs-lookup"><span data-stu-id="65f0f-218">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> <span data-ttu-id="65f0f-219">Ao usar um dos métodos de criação do `options` objeto, você pode especificar parâmetros opcionais em qualquer ordem, desde que seus nomes sejam especificados corretamente.</span><span class="sxs-lookup"><span data-stu-id="65f0f-219">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="65f0f-220">O exemplo a seguir mostra como chamar o `Document.setSelectedDataAsync` método especificando parâmetros opcionais em um `options` objeto.</span><span class="sxs-lookup"><span data-stu-id="65f0f-220">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>

```js
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="65f0f-221">Em ambos os exemplos de parâmetros opcionais, o parâmetro _callback_ é especificado como o último parâmetro (seguindo os parâmetros opcionais embutidos ou seguindo o objeto de argumento _Options_ ).</span><span class="sxs-lookup"><span data-stu-id="65f0f-221">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object).</span></span> <span data-ttu-id="65f0f-222">Como alternativa, você pode especificar o parâmetro _callback_ dentro o objeto JSON embutido ou no objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="65f0f-222">Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object.</span></span> <span data-ttu-id="65f0f-223">No entanto, você pode transmitir o parâmetro _callback_ em um só local: no objeto _options_ (embutido ou criado externamente) ou como o último parâmetro, mas não ambos.</span><span class="sxs-lookup"><span data-stu-id="65f0f-223">However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>

## <a name="wrap-common-apis-in-promise-returning-functions"></a><span data-ttu-id="65f0f-224">Dispor APIs comuns em funções que retornam a promessa</span><span class="sxs-lookup"><span data-stu-id="65f0f-224">Wrap Common APIs in Promise-returning functions</span></span>

<span data-ttu-id="65f0f-225">Os métodos Common API (e do Outlook API) não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="65f0f-225">The Common API (and Outlook API) methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="65f0f-226">Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída.</span><span class="sxs-lookup"><span data-stu-id="65f0f-226">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="65f0f-227">Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.</span><span class="sxs-lookup"><span data-stu-id="65f0f-227">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span> 

<span data-ttu-id="65f0f-228">O padrão básico é criar um método assíncrono que retorna um objeto Promise imediatamente e *resolve* esse objeto Promise quando o método interno é concluído ou *rejeita* o objeto se o método falhar.</span><span class="sxs-lookup"><span data-stu-id="65f0f-228">The basic pattern is to create an asynchronous method that returns a Promise object immediately and *resolves* that Promise object when the inner method completes, or *rejects* the object if the method fails.</span></span> <span data-ttu-id="65f0f-229">Veja a seguir um exemplo simples</span><span class="sxs-lookup"><span data-stu-id="65f0f-229">The following is a simple example</span></span>

```javascript
function getDocumentFilePath() {
    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                resolve(asyncResult.value.url);
            });
        }
        catch (error) {
            reject(WordMarkdownConversion.errorHandler(error));
        }
    })
}
```

<span data-ttu-id="65f0f-230">Quando esse método precisa ser aguardado, ele pode ser chamado com a `await` palavra-chave ou como a função passada para uma `then` função.</span><span class="sxs-lookup"><span data-stu-id="65f0f-230">When this method needs to be awaited, it can be called either with the `await` keyword or as the function passed to a `then` function.</span></span>

> [!NOTE]
> <span data-ttu-id="65f0f-231">Essa técnica é especialmente útil quando você precisa chamar uma das APIs comuns dentro de uma chamada do `run` método em um dos modelos de objeto específicos do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="65f0f-231">This technique is especially useful when you need to call one of the Common APIs inside a call of the `run` method in one of the application-specific object models.</span></span> <span data-ttu-id="65f0f-232">Para obter um exemplo da função acima sendo usada dessa forma, confira o arquivo [Home.js no exemplo de Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).</span><span class="sxs-lookup"><span data-stu-id="65f0f-232">For an example of the function above being used in this way, see the file [Home.js in the sample Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).</span></span>

<span data-ttu-id="65f0f-233">Veja a seguir um exemplo usando TypeScript.</span><span class="sxs-lookup"><span data-stu-id="65f0f-233">The following is an example using TypeScript.</span></span>

```typescript
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

## <a name="see-also"></a><span data-ttu-id="65f0f-234">Confira também</span><span class="sxs-lookup"><span data-stu-id="65f0f-234">See also</span></span>

- [<span data-ttu-id="65f0f-235">Entendendo a API de JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="65f0f-235">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="65f0f-236">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="65f0f-236">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
