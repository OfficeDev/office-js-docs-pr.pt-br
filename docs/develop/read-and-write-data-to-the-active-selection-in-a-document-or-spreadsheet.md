
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Ler e gravar dados na seleção ativa em um documento ou uma planilha

O objeto [Document](http://dev.office.com/reference/add-ins/shared/document) expõe métodos que permitem ler e gravar a seleção atual do usuário em um documento ou uma planilha. Para fazer isso, o objeto **Document** fornece os métodos **getSelectedDataAsync** e **setSelectedDataAsync**. Este tópico também descreve como ler, gravar e criar manipuladores de eventos para detectar alterações na seleção do usuário.

O método **getSelectedDataAsync** só funciona em relação à seleção atual do usuário. Se você precisar persistir a seleção no documento de forma que a mesma seleção esteja disponível para ler e gravar entre sessões de execução do suplemento, adicione uma associação usando o método[Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) (ou crie uma associação com um dos outros métodos "addFrom" do objeto [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx)). Para saber mais sobre como criar uma associação a uma região de um documento e a leitura e a gravação em uma associação, confira [Associar a regiões em um documento ou uma planilha](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Ler dados selecionados


O exemplo a seguir mostra como obter dados de uma seleção em um documento usando o método [getSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.getselecteddataasync).


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

No exemplo, o primeiro parâmetro _coercionType_ é especificado como **Office.CoercionType.Text** (você também pode especificar esse parâmetro usando a cadeia de caracteres literal `"text"`). Isso significa que a propriedade [value](http://dev.office.com/reference/add-ins/shared/asyncresult.status) do objeto [AsyncResult](http://dev.office.com/reference/add-ins/shared/asyncresult), que está disponível por meio do parâmetro _asyncResult_ na função de retorno de chamada, retorna uma **string** que contém o texto selecionado no documento. A especificação de tipos diferentes de coerção resulta em valores diferentes. [Office.CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) é uma enumeração dos valores de tipos de coerção disponíveis. **Office.CoercionType.Text** é avaliado como a cadeia de caracteres "text".


 >**Dica:**   **Quando devo usar a matriz ou a tabela coercionType para o acesso de dados?** Se for preciso que os dados tabulares selecionados cresçam dinamicamente quando linhas e colunas forem adicionadas, e você precisar trabalhar com os cabeçalhos da tabela, use o tipo de dados da tabela (especificando o parâmetro _coercionType_ do método **getSelectedDataAsync** como `"table"` ou **Office.CoercionType.Table**). A adição de linhas e colunas na estrutura de dados tem suporte nos dados de tabela e matriz, mas o acréscimo de linhas e colunas só tem suporte para dados de tabela. Se você não planeja adicionar linhas e colunas, e os dados não exigem a funcionalidade do cabeçalho, use o tipo de dados de matriz (especificando o parâmetro _coercionType_ do método **getSelecteDataAsync** como `"matrix"` ou **Office.CoercionType.Matrix**), que fornece um modelo mais simples para interagir com os dados.

A função anônima que é transmitida para a função como o segundo parâmetro de _callback_ é executada quando a operação **getSelectedDataAsync** é concluída. A função é chamada com um único parâmetro, _asyncResult_, que contém o resultado e o status da chamada. Se a chamada falhar, a propriedade [error](http://dev.office.com/reference/add-ins/shared/asyncresult.context) do objeto **AsyncResult** fornece acesso ao objeto [Error](http://dev.office.com/reference/add-ins/shared/error). Você pode verificar o valor das propriedades [Error.name](http://dev.office.com/reference/add-ins/shared/error.name) e [Error.message](http://dev.office.com/reference/add-ins/shared/error.message) para determinar por quê a operação set falhou. Caso contrário, o texto selecionado no documento é exibido.

A propriedade [AsyncResult.status](http://dev.office.com/reference/add-ins/shared/asyncresult.error) é usada na instrução **if** para testar se a chamada foi bem-sucedida. [Office.AsyncResultStatus](http://dev.office.com/reference/add-ins/shared/asyncresultstatus-enumeration) é uma enumeração de valores disponíveis da propriedade **AsyncResult.status**. **Office.AsyncResultStatus.Failed** é avaliado na cadeia de caracteres "failed" (e também pode ser especificado como essa cadeia de caracteres literal).


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

A função anônima passada para o método [setSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) como o parâmetro _callback_ é executada quando a chamada assíncrona é concluída. Ao gravar dados na seleção usando o método **setSelectedDataAsync**, o parâmetro _asyncResult_ do retorno de chamada fornece acesso somente ao status da chamada e ao objeto [Error](http://dev.office.com/reference/add-ins/shared/error), se a chamada falhar.

 **Observação:** a partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao gravar uma tabela na seleção atual](../../docs/excel/format-tables-in-add-ins-for-excel.md).


## <a name="detect-changes-in-the-selection"></a>Detectar alterações na seleção


O exemplo a seguir mostra como detectar alterações na seleção usando o método [Document.addHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.addhandlerasync) para adicionar um manipulador de eventos ao evento [SelectionChanged](http://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) no documento.


```
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

O primeiro parâmetro _eventType_ especifica o nome do evento a ser assinado. Passar a cadeia de caracteres `"documentSelectionChanged"` para esse parâmetro é equivalente a passar o tipo de evento **Office.EventType.DocumentSelectionChanged** da enumeração [Office.EventType](http://dev.office.com/reference/add-ins/shared/eventtype-enumeration).

A função `myHander()` que é passada para a função como o segundo parâmetro _handler_ é um manipulador de eventos executado ao alterar a seleção no documento. A função é chamada com um único parâmetro, _eventArgs_, que conterá uma referência a um objeto [DocumentSelectionChangedEventArgs](http://dev.office.com/reference/add-ins/shared/document.selectionchangedeventargs) quando a operação assíncrona for concluída. Você pode usar a propriedade [DocumentSelectionChangedEventArgs.document](http://dev.office.com/reference/add-ins/shared/document.selectionchangedeventargs.document) para acessar o documento que gerou o evento.


 >**Observação**  É possível adicionar vários manipuladores de eventos para um evento específico chamando o método **addHandlerAsync** novamente e passando uma função de manipulador de eventos adicional para o parâmetro _handler_. Isso funcionará corretamente, contanto que o nome de cada função do manipulador de eventos seja exclusivo.


## <a name="stop-detecting-changes-in-the-selection"></a>Parar de detectar alterações na seleção


O exemplo a seguir mostra como deixar de ouvir o evento [Document.SelectionChanged](http://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) chamando o método [document.removeHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.removehandlerasync).


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

O nome de função `myHandler` que é passado como o segundo parâmetro _handler_ especifica o manipulador de eventos que será removido do evento **SelectionChanged**.


 >**Importante:**  Se o parâmetro _handler_ opcional for omitido quando o método **removeHandlerAsync** for chamado, todos os manipuladores de eventos do _eventType_ especificado serão removidos.

