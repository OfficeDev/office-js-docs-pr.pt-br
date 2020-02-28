---
title: Ler e gravar dados na seleção ativa em um documento ou em uma planilha
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 039631e935d2ff6fadb4eab9d99df73ac30dae4d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325000"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Ler e gravar dados na seleção ativa em um documento ou em uma planilha

O objeto [Document](/javascript/api/office/office.document) expõe métodos que permitem que você leia e grave na seleção atual do usuário em um documento ou planilha. Para fazer isso, o `Document` objeto fornece os `getSelectedDataAsync` métodos `setSelectedDataAsync` e. Este tópico também descreve como ler, gravar e criar manipuladores de eventos para detectar alterações na seleção do usuário.

O `getSelectedDataAsync` método funciona apenas na seleção atual do usuário. Se você precisar persistir a seleção no documento, para que a mesma seleção esteja disponível para ler e gravar nas sessões de execução do seu suplemento, você deve adicionar uma associação usando o método [bindings. addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (ou criar uma associação com um dos outros métodos "addfrom" do objeto [bindings](/javascript/api/office/office.bindings) ). Para obter informações sobre como criar uma associação a uma região de um documento e, em seguida, ler e gravar em uma associação, confira [associar a regiões em um documento ou planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Ler dados selecionados


O exemplo a seguir mostra como obter dados de uma seleção em um documento usando o método [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).


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

Neste exemplo, o primeiro parâmetro _coercionType_ é especificado como `Office.CoercionType.Text` (você também pode especificar esse parâmetro usando a cadeia de caracteres `"text"`literal). Isso significa que a propriedade [Value](/javascript/api/office/office.asyncresult#status) do objeto [AsyncResult](/javascript/api/office/office.asyncresult) que está disponível no parâmetro _AsyncResult_ na função de retorno de chamada retornará uma **cadeia de caracteres** que contém o texto selecionado no documento. A especificação de diferentes tipos de coerção resultará em valores diferentes. [Office. CoercionType](/javascript/api/office/office.coerciontype) é uma enumeração dos valores de tipo de coerção disponíveis. `Office.CoercionType.Text` avalia a cadeia de caracteres "text".


> [!TIP]
> **Quando você deve usar a matriz versus a tabela coercionType para acesso a dados?** Se você precisar que os dados tabulares selecionados sejam expandidos dinamicamente quando as linhas e colunas forem adicionadas e você precisar trabalhar com cabeçalhos de tabela, deverá usar o tipo de dados __ Table (especificando o `getSelectedDataAsync` parâmetro coercionType `"table"` do `Office.CoercionType.Table`método como ou). A adição de linhas e colunas dentro da estrutura de dados é suportada em dados de tabela e matriz, mas só é possível acrescentar linhas e colunas para dados de tabela. Se você não estiver planejando a adição de linhas e colunas e seus dados não exigirem a funcionalidade de cabeçalho, você deverá usar o tipo de dados Matrix (especificando __ o parâmetro coercionType `getSelectedDataAsync` do método `"matrix"` como `Office.CoercionType.Matrix`ou), que fornece um modelo mais simples de interagir com os dados.

A função anônima que é passada para a função como o segundo parâmetro de _retorno de chamada_ é `getSelectedDataAsync` executada quando a operação é concluída. A função é chamada com um único parâmetro, _AsyncResult_, que contém o resultado e o status da chamada. Se a chamada falhar, a propriedade [Error](/javascript/api/office/office.asyncresult#asynccontext) do `AsyncResult` objeto fornece acesso ao objeto [Error](/javascript/api/office/office.error) . Você pode verificar o valor das propriedades [Error.Name](/javascript/api/office/office.error#name) e [Error. Message](/javascript/api/office/office.error#message) para determinar por que a operação de definição falhou. Caso contrário, o texto selecionado no documento será exibido.

A propriedade [AsyncResult. status](/javascript/api/office/office.asyncresult#error) é usada na instrução **If** para testar se a chamada foi bem-sucedida. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) é uma enumeração dos valores `AsyncResult.status` de propriedade disponíveis. `Office.AsyncResultStatus.Failed` avalia a cadeia de caracteres "failed" (e, novamente, também pode ser especificada como essa cadeia de caracteres literal).


## <a name="write-data-to-the-selection"></a>Gravar dados na seleção


O exemplo a seguir mostra como definir a seleção para mostrar "Hello World!".


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

Passar diferentes tipos de objeto para o parâmetro _data_ terá resultados diferentes. O resultado depende do que está selecionado no documento no momento, qual aplicativo está hospedando o suplemento e se os dados passados podem ser forçados para a seleção atual.

A função anônima passada para o método [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) como o parâmetro _callback_ é executada quando a chamada assíncrona é concluída. Quando você escreve dados na seleção usando o `setSelectedDataAsync` método, o parâmetro _AsyncResult_ do retorno de chamada fornece acesso somente ao status da chamada e ao objeto [Error](/javascript/api/office/office.error) se a chamada falhar.

> [!NOTE]
> A partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao gravar uma tabela na seleção atual](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Detectar alterações na seleção


O exemplo a seguir mostra como detectar alterações na seleção usando o método [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) para adicionar um manipulador de eventos ao evento [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) no documento.


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

O primeiro parâmetro _EventType_ especifica o nome do evento para se inscrever. Passar a cadeia `"documentSelectionChanged"` de caracteres para esse parâmetro é equivalente a `Office.EventType.DocumentSelectionChanged` passar o tipo de evento da enumeração [Office. EventType](/javascript/api/office/office.eventtype) .

A função `myHander()` que é passada para a função como o segundo parâmetro _handler_ é um manipulador de eventos executado ao alterar a seleção no documento. A função é chamada com um único parâmetro, _eventArgs_, que conterá uma referência a um objeto [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quando a operação assíncrona for concluída. Você pode usar a propriedade [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) para acessar o documento que gerou o evento.


> [!NOTE]
> Você pode adicionar vários manipuladores de eventos para um determinado evento chamando o `addHandlerAsync` método novamente e passando uma função de manipulador de eventos adicional para o parâmetro _Handler_ . Isso funcionará corretamente, desde que o nome de cada função do manipulador de eventos seja exclusivo.


## <a name="stop-detecting-changes-in-the-selection"></a>Parar de detectar alterações na seleção


O exemplo a seguir mostra como deixar de ouvir o evento [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) chamando o método [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

O `myHandler` nome da função que é passado como o segundo parâmetro _Handler_ especifica o manipulador de eventos que será removido do `SelectionChanged` evento.


> [!IMPORTANT]
> Se o parâmetro _Handler_ opcional for omitido quando `removeHandlerAsync` o método for chamado, todos os manipuladores de eventos do _EventType_ especificado serão removidos.
