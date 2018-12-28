---
title: Ler e gravar dados na seleção ativa em um documento ou em uma planilha
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 76f6f5f6a2d117b59e1a7794e35e181383022269
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457884"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Ler e gravar dados na seleção ativa em um documento ou em uma planilha

O objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document) expõe métodos que permitem ler e gravar a seleção atual do usuário em um documento ou uma planilha. Para fazer isso, o objeto **Document** fornece os métodos **getSelectedDataAsync** e **setSelectedDataAsync**. Este tópico também descreve como ler, gravar e criar manipuladores de eventos para detectar alterações na seleção do usuário.

O método **getSelectedDataAsync** só funciona em relação à seleção atual do usuário. Se você precisar persistir a seleção no documento de forma que a mesma seleção esteja disponível para ler e gravar entre sessões de execução do suplemento, adicione uma associação usando o método[Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (ou crie uma associação com um dos outros métodos "addFrom" do objeto [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings)). Para saber mais sobre como criar uma associação a uma região de um documento e a leitura e a gravação em uma associação, confira [Associar a regiões em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Ler dados selecionados


O exemplo a seguir mostra como obter dados de uma seleção em um documento usando o método [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).


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

No exemplo, o primeiro parâmetro _coercionType_ é especificado como **Office.CoercionType.Text** (você também pode especificar esse parâmetro usando a cadeia de caracteres literal `"text"`). Isso significa que a propriedade [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) do objeto [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult), que está disponível por meio do parâmetro _asyncResult_ na função de retorno de chamada, retorna uma **string** que contém o texto selecionado no documento. A especificação de tipos diferentes de coerção resulta em valores diferentes. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype) é uma enumeração dos valores de tipos de coerção disponíveis. **Office.CoercionType.Text** é avaliado como a cadeia de caracteres "text".


> [!TIP]
> **Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se for preciso que os dados tabulares selecionados cresçam dinamicamente quando linhas e colunas forem adicionadas, e você precisar trabalhar com os cabeçalhos da tabela, use o tipo de dados da tabela (especificando o parâmetro _coercionType_ do método **getSelectedDataAsync** como `"table"` ou **Office.CoercionType.Table**). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não planeja adicionar linhas e colunas, e os dados não exigem a funcionalidade do cabeçalho, use o tipo de dados de matriz (especificando o parâmetro _coercionType_ do método** getSelecteDataAsync** como `"matrix"` ou **Office.CoercionType.Matrix**), que fornece um modelo mais simples para interagir com os dados.

A função anônima que é transmitida para a função como o segundo parâmetro de _callback_ é executada quando a operação **getSelectedDataAsync** é concluída. A função é chamada com um único parâmetro, _asyncResult_, que contém o resultado e o status da chamada. Se a chamada falhar, a propriedade [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult#asynccontext) do objeto **AsyncResult** fornece acesso ao objeto [Error](https://docs.microsoft.com/javascript/api/office/office.error). Você pode verificar o valor das propriedades [Error.name](https://docs.microsoft.com/javascript/api/office/office.error#name) e [Error.message](https://docs.microsoft.com/javascript/api/office/office.error#message) para determinar por quê a operação set falhou. Caso contrário, o texto selecionado no documento é exibido.

A propriedade [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult#error) é usada na instrução **if** para testar se a chamada foi bem-sucedida. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) é uma enumeração de valores disponíveis da propriedade **AsyncResult.status**. **Office.AsyncResultStatus.Failed** é avaliado na cadeia de caracteres "failed" (e também pode ser especificado como essa cadeia de caracteres literal).


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

A função anônima passada para o método [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) como o parâmetro _callback_ é executada quando a chamada assíncrona é concluída. Ao gravar dados na seleção usando o método **setSelectedDataAsync**, o parâmetro _asyncResult_ do retorno de chamada fornece acesso somente ao status da chamada e ao objeto [Error](https://docs.microsoft.com/javascript/api/office/office.error), se a chamada falhar.

> [!NOTE]
> A partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao gravar uma tabela na seleção atual](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Detectar alterações na seleção


O exemplo a seguir mostra como detectar alterações na seleção usando o método [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) para adicionar um manipulador de eventos ao evento [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) no documento.


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

O primeiro parâmetro _eventType_ especifica o nome do evento a ser assinado. Passar a cadeia de caracteres `"documentSelectionChanged"` para esse parâmetro é equivalente a passar o tipo de evento **Office.EventType.DocumentSelectionChanged** da enumeração [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype).

A função `myHander()` que é passada para a função como o segundo parâmetro _handler_ é um manipulador de eventos executado ao alterar a seleção no documento. A função é chamada com um único parâmetro, _eventArgs_, que conterá uma referência a um objeto [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) quando a operação assíncrona for concluída. Você pode usar a propriedade [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs#document) para acessar o documento que gerou o evento.


> [!NOTE]
> Você pode adicionar vários manipuladores de eventos para um determinado evento chamando o método **addHandlerAsync** novamente e transmitindo uma função de manipulador de eventos adicional para o parâmetro _handler_. Isso funcionará corretamente desde que o nome de cada função do manipulador de eventos seja exclusivo.


## <a name="stop-detecting-changes-in-the-selection"></a>Parar de detectar alterações na seleção


O exemplo a seguir mostra como deixar de ouvir o evento [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) chamando o método [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

O nome de função `myHandler` que é passado como o segundo parâmetro _handler_ especifica o manipulador de eventos que será removido do evento **SelectionChanged**.


> [!IMPORTANT]
> Se o parâmetro _handler_ opcional for omitido quando o método **removeHandlerAsync** for chamado, todos os manipuladores de eventos do _eventType_ especificado serão removidos.

