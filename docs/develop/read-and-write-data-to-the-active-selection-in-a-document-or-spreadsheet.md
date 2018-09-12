---
title: Ler e gravar dados na seleção ativa em um documento ou em uma planilha
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d1c8fcdeec8d92fd3f77e169dc24715f7c5e9964
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944983"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="81cf2-102">Ler e gravar dados na seleção ativa em um documento ou em uma planilha</span><span class="sxs-lookup"><span data-stu-id="81cf2-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="81cf2-p101">O objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) expõe métodos que permitem ler e gravar a seleção atual do usuário em um documento ou uma planilha. Para fazer isso, o objeto **Document** fornece os métodos **getSelectedDataAsync** e **setSelectedDataAsync**. Este tópico também descreve como ler, gravar e criar manipuladores de eventos para detectar alterações na seleção do usuário.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p101">The [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="81cf2-p102">O método **getSelectedDataAsync** só funciona em relação à seleção atual do usuário. Se você precisar persistir a seleção no documento de forma que a mesma seleção esteja disponível para ler e gravar entre sessões de execução do suplemento, adicione uma associação usando o método[Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) (ou crie uma associação com um dos outros métodos "addFrom" do objeto [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js)). Para saber mais sobre como criar uma associação a uma região de um documento e a leitura e a gravação em uma associação, confira [Associar a regiões em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="81cf2-p102">The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="81cf2-109">Ler dados selecionados</span><span class="sxs-lookup"><span data-stu-id="81cf2-109">Read selected data</span></span>


<span data-ttu-id="81cf2-110">O exemplo a seguir mostra como obter dados de uma seleção em um documento usando o método [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="81cf2-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="81cf2-p103">No exemplo, o primeiro parâmetro _coercionType_ é especificado como **Office.CoercionType.Text** (você também pode especificar esse parâmetro usando a cadeia de caracteres literal `"text"`). Isso significa que a propriedade [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) do objeto [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js), que está disponível por meio do parâmetro _asyncResult_ na função de retorno de chamada, retorna uma **string** que contém o texto selecionado no documento. A especificação de tipos diferentes de coerção resulta em valores diferentes. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) é uma enumeração dos valores de tipos de coerção disponíveis. **Office.CoercionType.Text** é avaliado como a cadeia de caracteres "text".</span><span class="sxs-lookup"><span data-stu-id="81cf2-p103">In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) property of the [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="81cf2-p104">**Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se for preciso que os dados tabulares selecionados cresçam dinamicamente quando linhas e colunas forem adicionadas, e você precisar trabalhar com os cabeçalhos da tabela, use o tipo de dados da tabela (especificando o parâmetro _coercionType_ do método **getSelectedDataAsync** como `"table"` ou **Office.CoercionType.Table**). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não planeja adicionar linhas e colunas, e os dados não exigem a funcionalidade do cabeçalho, use o tipo de dados de matriz (especificando o parâmetro _coercionType_ do método\*\* getSelecteDataAsync\*\* como `"matrix"` ou **Office.CoercionType.Matrix**), que fornece um modelo mais simples para interagir com os dados.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="81cf2-p105">A função anônima que é transmitida para a função como o segundo parâmetro de _callback_ é executada quando a operação **getSelectedDataAsync** é concluída. A função é chamada com um único parâmetro, _asyncResult_, que contém o resultado e o status da chamada. Se a chamada falhar, a propriedade [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) do objeto **AsyncResult** fornece acesso ao objeto [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js). Você pode verificar o valor das propriedades [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) e [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) para determinar por quê a operação set falhou. Caso contrário, o texto selecionado no documento é exibido.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) property of the **AsyncResult** object provides access to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object. You can check the value of the [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) and [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="81cf2-p106">A propriedade [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) é usada na instrução **if** para testar se a chamada foi bem-sucedida. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) é uma enumeração de valores disponíveis da propriedade **AsyncResult.status**. **Office.AsyncResultStatus.Failed** é avaliado na cadeia de caracteres "failed" (e também pode ser especificado como essa cadeia de caracteres literal).</span><span class="sxs-lookup"><span data-stu-id="81cf2-p106">The [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="81cf2-128">Gravar dados na seleção</span><span class="sxs-lookup"><span data-stu-id="81cf2-128">Write data to the selection</span></span>


<span data-ttu-id="81cf2-129">O exemplo a seguir mostra como definir a seleção para mostrar "Hello World!".</span><span class="sxs-lookup"><span data-stu-id="81cf2-129">The following example shows how to set the selection to show "Hello World!".</span></span>


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="81cf2-p107">Passar diferentes tipos de objeto para o parâmetro _data_ terá resultados diferentes. O resultado depende do que está selecionado no documento no momento, qual aplicativo está hospedando o suplemento e se os dados passados podem ser forçados para a seleção atual.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="81cf2-p108">A função anônima passada para o método [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) como o parâmetro _callback_ é executada quando a chamada assíncrona é concluída. Ao gravar dados na seleção usando o método **setSelectedDataAsync**, o parâmetro _asyncResult_ do retorno de chamada fornece acesso somente ao status da chamada e ao objeto [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js), se a chamada falhar.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p108">The anonymous function passed into the [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="81cf2-134">A partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao gravar uma tabela na seleção atual](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="81cf2-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="81cf2-135">Detectar alterações na seleção</span><span class="sxs-lookup"><span data-stu-id="81cf2-135">Detect changes in the selection</span></span>


<span data-ttu-id="81cf2-136">O exemplo a seguir mostra como detectar alterações na seleção usando o método [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) para adicionar um manipulador de eventos ao evento [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) no documento.</span><span class="sxs-lookup"><span data-stu-id="81cf2-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event on the document.</span></span>


```js
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="81cf2-p109">O primeiro parâmetro _eventType_ especifica o nome do evento a ser assinado. Passar a cadeia de caracteres `"documentSelectionChanged"` para esse parâmetro é equivalente a passar o tipo de evento **Office.EventType.DocumentSelectionChanged** da enumeração [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="81cf2-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js) enumeration.</span></span>

<span data-ttu-id="81cf2-p110">A função `myHander()` que é passada para a função como o segundo parâmetro _handler_ é um manipulador de eventos executado ao alterar a seleção no documento. A função é chamada com um único parâmetro, _eventArgs_, que conterá uma referência a um objeto [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) quando a operação assíncrona for concluída. Você pode usar a propriedade [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) para acessar o documento que gerou o evento.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="81cf2-p111">Você pode adicionar vários manipuladores de eventos para um determinado evento chamando o método **addHandlerAsync** novamente e transmitindo uma função de manipulador de eventos adicional para o parâmetro _handler_. Isso funcionará corretamente desde que o nome de cada função do manipulador de eventos seja exclusivo.</span><span class="sxs-lookup"><span data-stu-id="81cf2-p111">You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="81cf2-144">Parar de detectar alterações na seleção</span><span class="sxs-lookup"><span data-stu-id="81cf2-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="81cf2-145">O exemplo a seguir mostra como deixar de ouvir o evento [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) chamando o método [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="81cf2-145">The following example shows how to stop listening to the [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event by calling the [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="81cf2-146">O nome de função `myHandler` que é passado como o segundo parâmetro _handler_ especifica o manipulador de eventos que será removido do evento **SelectionChanged**.</span><span class="sxs-lookup"><span data-stu-id="81cf2-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="81cf2-147">Se o parâmetro _handler_ opcional for omitido quando o método **removeHandlerAsync** for chamado, todos os manipuladores de eventos do _eventType_ especificado serão removidos.</span><span class="sxs-lookup"><span data-stu-id="81cf2-147">If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.</span></span>

