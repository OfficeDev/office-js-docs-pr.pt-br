---
title: Programa??o ass?ncrona em Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d251ebfd03227569b9a24bcd7f17baada6099938
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programa??o ass?ncrona em Suplementos do Office

Por que a API de Suplementos do Office usa a programa??o ass?ncrona? Como o JavaScript ? uma linguagem de thread ?nico, se o script invocar um processo s?ncrono demorado, todas as execu??es subsequentes do script ser?o bloqueadas at? que o processo seja conclu?do. Como certas opera??es para clientes Web do Office (mas tamb?m para clientes avan?ados) podem impedir a execu??o se estiverem sendo executadas em sincronia, a maioria dos m?todos na API do JavaScript para Office foi desenvolvido para execu??o ass?ncrona. Isso garante que os Suplementos do Office sejam responsivos e tenham alto desempenho. Em geral, isso tamb?m requer que voc? escreva fun??es de retorno de chamada ao trabalhar com esses m?todos ass?ncronos.

Os nomes de todos os m?todos ass?ncronos na API terminam com "Async", como os m?todos [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) ou [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item). Quando um m?todo "Async" ? chamado, ele ? executado imediatamente e qualquer execu??o subsequente do script poder? continuar. A fun??o de retorno de chamada opcional que voc? passar para um m?todo de "Async" ? executada assim que os dados ou a opera??o solicitada est? pronta. Isso geralmente ocorre imediatamente, mas pode haver um pequeno atraso antes de retornar.

O diagrama a seguir mostra o fluxo de execu??o de uma chamada para um m?todo de "Async" que l? os dados selecionados pelo usu?rio em um documento aberto no Word Online ou Excel Online baseados no servidor. No ponto em que a chamada "Async" ? feita, o thread de execu??o do JavaScript fica livre para executar qualquer processamento adicional do lado do cliente. (Embora nenhum seja mostrado no diagrama.) Quando o m?todo "Async" retorna, o retorno de chamada retoma a execu??o no thread e o suplemento pode acessar os dados, fazer algo com eles e exibir os resultados. O mesmo padr?o de execu??o ass?ncrona ocorre ao trabalhar com aplicativos host de clientes avan?ados do Office, como Word 2013 ou Excel 2013.

*Figura 1. Fluxo de execu??o da programa??o ass?ncrono*

![Fluxo de execu??o do encadeamento da programa??o ass?ncrono](../images/office15-app-async-prog-fig01.png)

O suporte a esse design ass?ncrono em clientes Web e avan?ados faz parte das metas de design "gravar plataforma cruzada j? executada" do modelo de desenvolvimento de Suplementos do Office. Por exemplo, voc? pode um suplemento do painel de tarefas ou conte?do com uma ?nica base de c?digo que ser? executada no Excel 2013 e Excel Online.

## <a name="writing-the-callback-function-for-an-async-method"></a>Gravar a fun??o de retorno de chamada para um m?todo "Async"


A fun??o de retorno de chamada que voc? transmite como o argumento _callback_ para um m?todo de "Async" deve declarar um ?nico par?metro que o tempo de execu??o do suplemento usar? para fornecer acesso a um objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) quando a fun??o de retorno de chamada for executada. Voc? pode gravar:


- Uma fun??o an?nima que deve ser gravada e transmitida diretamente embutida com a chamada para o m?todo "Async" como o par?metro _callback_ do m?todo "Async".
    
- Uma fun??o nomeada, transmitindo o nome da fun??o como o par?metro _callback_ de um m?todo "Async".
    
Uma fun??o an?nima ? ?til se voc? s? for usar seu c?digo uma vez ? porque ele n?o possui um nome, voc? n?o pode referenci?-la em outra parte do seu c?digo. Uma fun??o nomeada ? ?til se voc? quiser reutilizar a fun??o retorno de chamada para mais de um m?todo "Async".


### <a name="writing-an-anonymous-callback-function"></a>Gravar uma fun??o de retorno de chamada an?nima

A seguinte fun??o de retorno de chamada an?nima declara um ?nico par?metro chamado `result` que recupera os dados da propriedade [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) quando o retorno de chamada retornar.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

O exemplo a seguir mostra como transmitir essa fun??o de retorno de chamada an?nima de acordo com o contexto de um retorno de chamada completo do m?todo "Async" para o m?todo **Document.getSelectedDataAsync**.


- O primeiro argumento _coercionType_, `Office.CoercionType.Text`, especifica para retornar os dados selecionados como uma cadeia de texto.
    
- O segundo argumento _callback_ ? a fun??o an?nima transmitida de acordo com o m?todo. Quando a fun??o ? executada, ela usa o par?metro _result_ para acessar a propriedade **value** do objeto **AsyncResult** para exibir os dados selecionados pelo usu?rio no documento.
    



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

Voc? tamb?m pode usar o par?metro da sua fun??o de retorno de chamada para acessar outras propriedades do objeto **AsyncResult**. Use a propriedade [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) para determinar se a chamada teve ?xito ou falhou. Se sua chamada falhar, voc? pode usar a propriedade [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) para acessar um objeto [Error](https://dev.office.com/reference/add-ins/shared/error) para informa??es sobre o erro.

Para saber mais sobre como usar o m?todo **getSelectedDataAsync**, consulte [Ler e gravar dados na se??o sele??o ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>Gravar uma fun??o de retorno de chamada nomeada

Como alternativa, voc? pode escrever uma fun??o nomeada e transmitir o nome dela para o par?metro _callback_ de um m?todo "Async". Por exemplo, o exemplo anterior pode ser reescrito para transmitir uma fun??o chamada `writeDataCallback` como o par?metro _callback_ assim.


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>Diferen?as entre o que ? retornado para a propriedade AsyncResult.value


As propriedades **asyncContext**, **status** e **error** do objeto **AsyncResult** retornam os mesmos tipos de informa??es para a fun??o de retorno de chamada transmitida para todos os m?todos de "Async". No entanto, o que ? retornado para a propriedade **AsyncResult.value** varia de acordo com a funcionalidade do m?todo "Async".

Por exemplo, os m?todos **addHandlerAsync** (dos objetos [Binding](https://dev.office.com/reference/add-ins/shared/binding), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), [Document](https://dev.office.com/reference/add-ins/shared/document), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) e [Settings](https://dev.office.com/reference/add-ins/shared/settings)) s?o usados para adicionar fun??es de manipulador de eventos aos itens representados por esses objetos. Voc? pode acessar a propriedade **AsyncResult.value** a partir da fun??o de retorno de chamada transmitida para qualquer um dos m?todos **addHandlerAsync**, mas como nenhum dado ou objeto est? sendo acessado quando voc? adiciona um manipulador de eventos, a propriedade **value** sempre retornar? **undefined** se voc? tentar acess?-la.

Por outro lado, se voc? chamar o m?todo **Document.getSelectedDataAsync**, ele retornar? os dados que os usu?rios selecionaram no documento para a propriedade **AsyncResult.value** no retorno de chamada. Ou, se voc? chamar o m?todo [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync), ele retornar? uma matriz de todos os objetos **Binding** no documento. E se voc? chamar o m?todo [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync), ele retornar? um ?nico objeto **Binding**.

Para obter uma descri??o do que ? retornado para a propriedade **AsyncResult.value** para um m?todo "Async", consulte a se??o "Valor de retorno de chamada" do t?pico de refer?ncia do m?todo. Para obter um resumo de todos os objetos que oferecem m?todos "Async", consulte a tabela na parte inferior do t?pico do objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult).


## <a name="asynchronous-programming-patterns"></a>Padr?es de programa??o ass?ncrona


A API do JavaScript para o Office oferece suporte a dois tipos de padr?es de programa??o ass?ncrona:


- Usando retornos de chamada aninhados
    
- Usando o padr?o de promessas
    
A programa??o ass?ncrona com fun??es de retorno de chamada frequentemente exigem que voc? aninhe o resultado retornado de um retorno de chamada dentro de dois ou mais retornos de chamada. Se voc? precisar fazer isso, ? poss?vel usar retornos de chamada aninhados de todos os m?todos "Async" da API.

Usar retornos de chamada aninhados ? um padr?o de programa??o familiar para a maioria dos desenvolvedores de JavaScript, mas c?digos com retornos de chamada profundamente aninhados podem ser dif?ceis de ler e entender. Como alternativa para retornos de chamada aninhados, a API do JavaScript para o Office tamb?m oferece suporte a uma implementa??o do padr?o de promessas. No entanto, na vers?o atual da API do JavaScript para o Office, o padr?o de promessas s? funciona com o c?digo para [associa??o em planilhas do Excel e documentos do Word](bind-to-regions-in-a-document-or-spreadsheet.md).

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programa??o ass?ncrona usando fun??es aninhadas de retorno de chamada


Frequentemente, voc? precisa executar duas ou mais opera??es ass?ncronas para concluir uma tarefa. Para fazer isso, voc? pode aninhar uma chamada "Async" dentro de outra. 

O exemplo de c?digo a seguir aninha duas ou mais chamadas ass?ncronas. 


- Primeiro, o m?todo [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ? chamado para acessar uma associa??o no documento chamado "MyBinding". O objeto **AsyncResult** retornado para o par?metro `result` do retorno de chamada fornece acesso ao objeto de associa??o especificado da propriedade **AsyncResult.value**.
    
- Em seguida, o objeto de associa??o acessado do primeiro par?metro `result` ? usado para chamar o m?todo [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync).
    
- Por fim, o par?metro `result2` do retorno de chamada transmitido para o m?todo**Binding.getDataAsync** ? usado para exibir os dados na associa??o.
    



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

Esse padr?o de retorno de chamada aninhado b?sico pode ser usado para todos os m?todos ass?ncronos na API do JavaScript para Office.

As se??es a seguir mostram como usar fun??es an?nimas ou nomeadas para retornos de chamada aninhados em m?todos ass?ncronos.


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Usando fun??es an?nimas para retornos de chamada aninhados

No exemplo a seguir, duas fun??es an?nimas s?o declaradas embutidas e transmitidas para os m?todos **getByIdAsync** e **getDataAsync** como retornos de chamada aninhados. Como as fun??es s?o simples e embutidas, a inten??o da implementa??o fica imediatamente clara.


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


#### <a name="using-named-functions-for-nested-callbacks"></a>Usando fun??es nomeadas para retornos de chamada aninhados

Em implementa??es complexas, pode ser ?til usar fun??es nomeadas para facilitar a leitura, manuten??o e reutiliza??o do seu c?digo. No exemplo a seguir, as duas fun??es an?nimas do exemplo na se??o anterior foram reescritas como fun??es nomeadas `deleteAllData` e `showResult`. Essas fun??es nomeadas s?o ent?o transmitidas para os m?todos **getByIdAsync** e **deleteAllDataValuesAsync** como retornos de chamada por nome.


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>Programa??o ass?ncrona usando o padr?o de promessas para acessar dados em associa??es


Em vez de transmitir a fun??o de retorno de chamada e aguardar at? que a fun??o retorne antes da continua??o da execu??o, o padr?o de programa??o de promessas retorna imediatamente retorna um objeto de promessa que representa o resultado desejado. No entanto, ao contr?rio da verdadeira programa??o s?ncrona, nos bastidores o cumprimento do resultado prometido ?, na verdade, adiado at? que o ambiente de tempo de execu??o dos Suplementos do Office possa concluir a solicita??o. Um manipulador _onError_ ? fornecido para atender a situa??es em que a solicita??o n?o pode ser cumprida.

A API do JavaScript para Office fornece o m?todo [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) para dar suporte ao padr?o de promessas para funcionar com objetos de associa??o existentes. O objeto de promessa retornado para o m?todo **Office.select** oferece suporte somente aos quatro m?todos que voc? pode acessar diretamente do objeto [Binding](https://dev.office.com/reference/add-ins/shared/binding): [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync), [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) e [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync).

O padr?o de promessas para funcionar com associa??es tem este formato:

 **Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_

O par?metro _selectorExpression_ assume a forma `"bindings#bindingId"`, em que _bindingId_ ? o nome (**id**) de uma associa??o que voc? criou anteriormente no documento ou planilha (usando um dos m?todos "addFrom" da cole??o **Bindings**: **addFromNamedItemAsync**, **addFromPromptAsync** ou **addFromSelectionAsync**). Por exemplo, a express?o seletora `bindings#cities` especifica que voc? deseja acessar a associa??o com uma **id** de "cidades".

O par?metro _onError_ ? uma fun??o de manipula??o de erro que usa um ?nico par?metro do tipo **AsyncResult** que pode ser usado para acessar um objeto **Error**, se o m?todo **select** falhar ao acessar as associa??es especificadas. O exemplo a seguir mostra uma fun??o de manipulador de erro b?sica que pode ser transmitida para o par?metro _onError_.




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

Substitua o espa?o reservado _BindingObjectAsyncMethod_ por uma chamada a qualquer um dos quatro m?todos de objeto **Binding** com suporte pelo objeto de promessa: **getDataAsync**, **setDataAsync**, **addHandlerAsync** ou **removeHandlerAsync**. As chamadas para esses m?todos n?o oferecem suporte a promessas adicionais. Voc? deve cham?-los usando o [padr?o de fun??o de retorno de chamada aninhado](#AsyncProgramming_NestedCallbacks).

Depois que uma promessa de objeto **Binding** ? cumprida, ela pode ser reutilizada na chamada do m?todo encadeada como se fosse uma associa??o (o tempo de execu??o do suplemento n?o tentar? novamente cumprir a promessa de forma ass?ncrona). Se a promessa do objeto **Binding** n?o puder ser cumprida, o tempo de execu??o do suplemento tentar? novamente acessar o objeto de associa??o da pr?xima vez que um dos seus m?todos ass?ncronos for chamado.

O exemplo de c?digo a seguir usa o m?todo **select** para recuperar uma associa??o com a **id** "`cities`" da cole??o **Bindings** e, em seguida, chama o m?todo [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) para adicionar um manipulador de eventos ao evento [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) da associa??o.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> A promessa do objeto **Binding** retornada pelo m?todo **Office.select** oferece acesso a apenas um dos quatro m?todos do objeto **Binding**. Se precisar acessar qualquer um dos outros membros do objeto **Binding**, voc? dever? usar a propriedade **Document.bindings** e os m?todos **Bindings.getByIdAsync** ou **Bindings.getAllAsync** para recuperar o objeto **Binding**. Por exemplo, se voc? precisar acessar qualquer uma das propriedades do objeto **Binding** (as propriedades **document**, **id** ou **type**) ou caso precise acessar as propriedades dos objetos [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) ou [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding), voc? dever? usar os m?todos **getByIdAsync** ou **getAllAsync** para recuperar um objeto **Binding**.


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Transmitir par?metros opcionais para m?todos ass?ncronos


A sintaxe comum para todos os m?todos "Async" segue este padr?o:

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_ `);`

Todos os m?todos ass?ncronos d?o suporte par?metros opcionais, que s?o passados como um objeto JSON (JavaScript Object Notation) contendo um ou mais par?metros opcionais. O objeto JSON que cont?m os par?metros opcionais ? uma cole??o desordenada de pares de valores e chaves com o caractere ":" separando os valores e as chaves. Cada par do objeto ? separado por v?rgula e o conjunto completo de pares ? inclu?do entre chaves. A chave ? o nome do par?metro e o valor ? o valor a ser transmitido para esse par?metro.

Voc? pode criar o objeto JSON que cont?m par?metros opcionais embutidos ou criando um objeto `options` e transmitindo ele como o par?metro _options_.


### <a name="passing-optional-parameters-inline"></a>Transmitir par?metros opcionais embutidos

Por exemplo, a sintaxe para chamar o m?todo [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) com par?metros opcionais embutidos tem esta apar?ncia:

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext:' asyncContext},callback);

```

Neste formul?rio da sintaxe de chamada, os dois par?metros opcionais, _coercionType_ e _asyncContext_, s?o definidos como um objeto JSON embutido entre chaves.

O exemplo a seguir mostra como chamar o m?todo **Document.setSelectedDataAsync** especificando par?metros opcionais embutidos.


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
> ? poss?vel especificar par?metros opcionais em qualquer ordem no objeto JSON desde que seus nomes sejam especificados corretamente.


### <a name="passing-optional-parameters-in-an-options-object"></a>Transmitir par?metros opcionais em um objeto de op??es

Como alternativa, voc? pode criar um objeto nomeado `options` que especifica os par?metros opcionais separadamente da chamada do m?todo e, em seguida, transmitir o objeto `options` como o argumento _options_.

O exemplo a seguir mostra uma maneira de criar o objeto `options`, onde `parameter1`, `value1` e assim por diante, s?o espa?os reservados para os valores e nomes reais de par?metros.




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

Que ? semelhante ao exemplo a seguir quando usado para especificar os par?metros [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) e [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration).




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Esta ? outra maneira de criar o objeto `options`.




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Que ? semelhante ao exemplo a seguir quando usado para especificar os par?metros **ValueFormat** e **FilterType**.


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> Ao usar um dos m?todos para a cria??o do objeto `options`, ? poss?vel especificar par?metros opcionais em qualquer ordem, desde que os nomes deles sejam especificados corretamente.

O exemplo a seguir mostra como chamar o m?todo **Document.setSelectedDataAsync** especificando par?metros opcionais em um objeto `options`.




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


Em ambos os exemplos de par?metros opcionais, o par?metro _callback_ ? especificado como o ?ltimo par?metro (acompanhando os par?metros opcionais embutido ou seguindo o objeto do argumento _options_). Como alternativa, voc? pode especificar o par?metro _callback_ dentro o objeto JSON embutido ou no objeto `options`. No entanto, voc? pode transmitir o par?metro _callback_ em um s? local: no objeto _options_ (embutido ou criado externamente) ou como o ?ltimo par?metro, mas n?o ambos.


## <a name="see-also"></a>Veja tamb?m

- [No??es b?sicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md) 
- [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)
     
