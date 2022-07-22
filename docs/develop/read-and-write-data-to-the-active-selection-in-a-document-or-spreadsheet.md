---
title: Ler e gravar dados na seleção ativa em um documento ou em uma planilha
description: Saiba como ler e gravar dados na seleção ativa em um documento do Word ou planilha do Excel.
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 220f768352aa3cf8a077f2e37ec812878cbeffba
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958458"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Ler e gravar dados na seleção ativa em um documento ou em uma planilha

O objeto [Document](/javascript/api/office/office.document) expõe métodos que permitem ler e gravar a seleção atual do usuário em um documento ou uma planilha. Para fazer isso, o `Document` objeto fornece os `getSelectedDataAsync` métodos `setSelectedDataAsync` e os métodos. Este tópico também descreve como ler, gravar e criar manipuladores de eventos para detectar alterações na seleção do usuário.

O `getSelectedDataAsync` método só funciona em relação à seleção atual do usuário. Se você precisar persistir a seleção no documento de forma que a mesma seleção esteja disponível para ler e gravar entre sessões de execução do suplemento, adicione uma associação usando o método[Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) (ou crie uma associação com um dos outros métodos "addFrom" do objeto [Bindings](/javascript/api/office/office.bindings)). Para saber mais sobre como criar uma associação a uma região de um documento e a leitura e a gravação em uma associação, confira [Associar a regiões em um documento ou uma planilha](bind-to-regions-in-a-document-or-spreadsheet.md).

## <a name="read-selected-data"></a>Ler dados selecionados

O exemplo a seguir mostra como obter dados de uma seleção em um documento usando o método [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)).

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

Neste exemplo, o primeiro parâmetro, _coercionType_, `Office.CoercionType.Text` é especificado como (você também pode especificar esse parâmetro usando a cadeia de caracteres literal `"text"`). Isso significa que a propriedade [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) do objeto [AsyncResult](/javascript/api/office/office.asyncresult), que está disponível por meio do parâmetro _asyncResult_ na função de retorno de chamada, retorne uma **string** que contenha o texto selecionado no documento. A especificação de tipos diferentes de coerção resulta em valores diferentes. [Office.CoercionType](/javascript/api/office/office.coerciontype) é uma enumeração dos valores de tipos de coerção disponíveis. `Office.CoercionType.Text` é avaliada como a cadeia de caracteres "text".

> [!TIP]
> **Quando devo usar a matriz ou a tabela coercionType para o acesso aos dados?** Se você precisar que os dados tabulares selecionados cresçam dinamicamente quando linhas e colunas forem adicionadas e precisar trabalhar com cabeçalhos de tabela, você deverá usar o tipo de dados da tabela (especificando o parâmetro _coercionType_ `getSelectedDataAsync` do método como `"table"` `Office.CoercionType.Table`ou ). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não estiver planejando adicionar linhas e colunas e seus dados não exigirem a funcionalidade de cabeçalho, use o tipo de dados de matriz (especificando o parâmetro  _coercionType_ `getSelectedDataAsync` do método como `"matrix"` `Office.CoercionType.Matrix`ou ), que fornece um modelo mais simples de interação com os dados.

A função anônima que é passada para o método como o segundo parâmetro, _retorno_ de chamada, é executada quando a `getSelectedDataAsync` operação é concluída. A função é chamada com um único parâmetro, _asyncResult_, que contém o resultado e o status da chamada. Se a chamada falhar, [a propriedade de](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) erro do `AsyncResult` objeto fornecerá acesso ao [objeto Error](/javascript/api/office/office.error) . Você pode verificar o valor das propriedades [Error.name](/javascript/api/office/office.error#office-office-error-name-member) e [Error.message](/javascript/api/office/office.error#office-office-error-message-member) para determinar por quê a operação set falhou. Caso contrário, o texto selecionado no documento é exibido.

A propriedade [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) é usada na instrução **if** para testar se a chamada foi bem-sucedida. [Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) é uma enumeração de valores de propriedade `AsyncResult.status` disponíveis. `Office.AsyncResultStatus.Failed` avalia a cadeia de caracteres "failed" (e, novamente, também pode ser especificada como essa cadeia de caracteres literal).

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

Passar diferentes tipos de objeto para o parâmetro _data_ terá resultados diferentes. O resultado depende do que está selecionado no documento no momento, qual aplicativo cliente do Office está hospedando seu suplemento e se os dados passados podem ser coaçados para a seleção atual.

A função anônima passada para o método [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) como o parâmetro _callback_ é executada quando a chamada assíncrona é concluída. `setSelectedDataAsync` Quando você grava dados na seleção usando o método, o parâmetro _asyncResult_ do retorno de chamada fornece acesso somente ao status da chamada e ao objeto [Error](/javascript/api/office/office.error) se a chamada falhar.

> [!NOTE]
> A partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao gravar uma tabela na seleção atual](../excel/excel-add-ins-tables.md).

## <a name="detect-changes-in-the-selection"></a>Detectar alterações na seleção

O exemplo a seguir mostra como detectar alterações na seleção usando o método [Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) para adicionar um manipulador de eventos ao evento [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) no documento.

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

O primeiro parâmetro, _eventType_, especifica o nome do evento no qual assinar. Passar a cadeia de `"documentSelectionChanged"` caracteres para esse parâmetro é equivalente a passar o `Office.EventType.DocumentSelectionChanged` tipo de evento da [enumeração Office.EventType](/javascript/api/office/office.eventtype) .

A  `myHandler()` função que é passada para o método como o segundo _parâmetro, manipulador_, é um manipulador de eventos executado quando a seleção é alterada no documento. A função é chamada com um único parâmetro, _eventArgs_, que conterá uma referência a um objeto [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quando a operação assíncrona for concluída. Você pode usar a propriedade [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) para acessar o documento que gerou o evento.

> [!NOTE]
> Você pode adicionar vários manipuladores de eventos para `addHandlerAsync` um determinado evento chamando o método novamente e passando uma função de manipulador de eventos adicional para o parâmetro _do_ manipulador. This will work correctly as long as the name of each event handler function is unique.

## <a name="stop-detecting-changes-in-the-selection"></a>Parar de detectar alterações na seleção

O exemplo a seguir mostra como deixar de ouvir o evento [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) chamando o método [document.removeHandlerAsync](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)).

```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

O  `myHandler` nome da função que é passado como o segundo _parâmetro, manipulador_, especifica o manipulador de eventos que será removido do `SelectionChanged` evento.

> [!IMPORTANT]
> Se o parâmetro  _de manipulador_ opcional for omitido `removeHandlerAsync` quando o método for chamado, todos os manipuladores de eventos para _o eventType_ especificado serão removidos.
