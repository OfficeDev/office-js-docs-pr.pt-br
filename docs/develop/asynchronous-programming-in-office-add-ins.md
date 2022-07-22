---
title: Programação assíncrona em Suplementos do Office
description: Saiba como a biblioteca JavaScript do Office usa programação assíncrona em Suplementos do Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f2d8682488f41786d60c8fcec02b120f35e696ae
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958857"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programação assíncrona em Suplementos do Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Por que a API de Suplementos do Office usa a programação assíncrona? Como o JavaScript é uma linguagem de thread único, se o script invocar um processo síncrono demorado, todas as execuções subsequentes do script serão bloqueadas até que o processo seja concluído. Como determinadas operações em clientes Web do Office (mas clientes avançados também) podem bloquear a execução se forem executadas de forma síncrona, a maioria das APIs JavaScript do Office foi projetada para ser executada de forma assíncrona. Isso garante que os Suplementos do Office sejam responsivos e rápidos. Em geral, isso também requer que você escreva funções de retorno de chamada ao trabalhar com esses métodos assíncronos.

Os nomes de todos os métodos assíncronos na API terminam com "Async", `Document.getSelectedDataAsync`como , `Binding.getDataAsync`ou `Item.loadCustomPropertiesAsync` métodos. Quando um método "Async" é chamado, ele é executado imediatamente e qualquer execução subsequente do script poderá continuar. A função de retorno de chamada opcional que você passar para um método de "Async" é executada assim que os dados ou a operação solicitada está pronta. Isso geralmente ocorre imediatamente, mas pode haver um pequeno atraso antes de retornar.

O diagrama a seguir mostra o fluxo de execução de uma chamada para um método "Assíncrono" que lê os dados que o usuário selecionou em um documento aberto no Word ou Excel baseado em servidor. No ponto em que a chamada "Assíncrona" é feita, o thread de execução do JavaScript é livre para executar qualquer processamento adicional do lado do cliente (embora nenhum seja mostrado no diagrama). Quando o método "Async" retorna, o retorno de chamada retoma a execução no thread e o suplemento pode acessar dados, fazer algo com ele e exibir o resultado. O mesmo padrão de execução assíncrona é válido ao trabalhar com os aplicativos cliente avançados do Office, como o Word 2013 ou o Excel 2013.

*Figura 1. Fluxo de execução da programação assíncrona*

![Diagrama mostrando a interação de execução de comando ao longo do tempo com o usuário, a página do suplemento e o servidor de aplicativo Web que hospeda o suplemento.](../images/office-addins-asynchronous-programming-flow.png)

O suporte a este design assíncrono em clientes Web e avançados faz parte das metas de design "gravar plataforma cruzada já executada" do modelo de desenvolvimento de Suplementos do Office. Por exemplo, você pode criar um suplemento do painel de tarefas ou conteúdo com uma única base de código que será executada no Excel 2013 e Excel Online.

## <a name="write-the-callback-function-for-an-async-method"></a>Gravar a função de retorno de chamada para um método "Async"

A função de retorno de chamada que você  passa como o argumento de retorno de chamada para um método "Async" deve declarar um único parâmetro que o runtime do suplemento usará para fornecer acesso a um objeto [AsyncResult](/javascript/api/office/office.asyncresult) quando a função de retorno de chamada for executada. Você pode gravar:

- Uma função anônima que deve ser gravada e passada diretamente em linha com a chamada para o método "Async" como o  parâmetro de retorno de chamada do método "Async".

- Uma função nomeada, passando o nome dessa função como o parâmetro *de retorno* de chamada de um método "Async".

Uma função anônima é útil se você só for usar seu código uma vez – porque ele não possui um nome, você não pode referenciá-la em outra parte do seu código. Uma função nomeada é útil se você quiser reutilizar a função retorno de chamada para mais de um método "Async".

### <a name="write-an-anonymous-callback-function"></a>Escrever uma função de retorno de chamada anônima

A função de retorno de chamada anônima a seguir declara um único `result` parâmetro chamado que recupera dados da [propriedade AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) quando o retorno de chamada retorna.

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

O exemplo a seguir mostra como passar essa função de retorno de chamada anônima em linha no contexto de uma chamada completa do método "Async" para o `Document.getSelectedDataAsync` método.

- O primeiro *argumento coercionType* , `Office.CoercionType.Text`especifica para retornar os dados selecionados como uma cadeia de caracteres de texto.

- O segundo *argumento de retorno* de chamada é a função anônima passada em linha para o método. Quando a função é executada,  `value` `AsyncResult` ela usa o parâmetro de resultado para acessar a propriedade do objeto para exibir os dados selecionados pelo usuário no documento.

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

Você também pode usar o parâmetro da função de retorno de chamada para acessar outras propriedades do `AsyncResult` objeto. Use a propriedade [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) para determinar se a chamada teve êxito ou falhou. Se sua chamada falhar, você pode usar a propriedade [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) para acessar um objeto [Error](/javascript/api/office/office.error) para informações sobre o erro.

Para obter mais informações sobre como usar o `getSelectedDataAsync` método, consulte [Ler e gravar dados na seleção ativa em um documento ou planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

### <a name="write-a-named-callback-function"></a>Escrever uma função de retorno de chamada nomeada

Como alternativa, você pode escrever uma função nomeada e passar seu nome para o parâmetro *de retorno* de chamada de um método "Async". Por exemplo, o exemplo anterior pode ser reescrito para transmitir uma função chamada `writeDataCallback` como o parâmetro *callback* assim.

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

## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>Diferenças entre o que é retornado para a propriedade AsyncResult.value

As `asyncContext`propriedades `status`, e `error` do objeto `AsyncResult` retornam os mesmos tipos de informações para a função de retorno de chamada passada para todos os métodos "Async". No entanto, o que é retornado `AsyncResult.value` para a propriedade varia dependendo da funcionalidade do método "Async".

Por exemplo, `addHandlerAsync` os métodos (dos objetos [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) e [Settings](/javascript/api/office/office.settings) ) são usados para adicionar funções de manipulador de eventos aos itens representados por esses objetos. Você pode acessar a propriedade da função de retorno de chamada que passa para qualquer um dos métodos, mas como nenhum dado ou objeto está sendo acessado quando você adiciona um manipulador de eventos, `value` a propriedade sempre  retorna indefinida se você tentar acessar.`AsyncResult.value` `addHandlerAsync`

Por outro lado, se você chamar `Document.getSelectedDataAsync` o método, `AsyncResult.value` ele retornará os dados que o usuário selecionou no documento para a propriedade no retorno de chamada. Ou, se você chamar o [método Bindings.getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) , `Binding` ele retornará uma matriz de todos os objetos no documento. E, se você chamar o [método Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) , ele retornará um único `Binding` objeto.

Para obter uma descrição do que é `AsyncResult.value` retornado à propriedade de um método, consulte a `Async` seção "Valor de retorno de chamada" do tópico de referência desse método. Para obter um resumo de todos os objetos que `Async` fornecem métodos, consulte a tabela na parte inferior do tópico do objeto [AsyncResult](/javascript/api/office/office.asyncresult) .

## <a name="asynchronous-programming-patterns"></a>Padrões de programação assíncrona

A API JavaScript do Office dá suporte a dois tipos de padrões de programação assíncrona.

- Usando retornos de chamada aninhados
- Usando o padrão de promessas

A programação assíncrona com funções de retorno de chamada frequentemente exigem que você aninhe o resultado retornado de um retorno de chamada dentro de dois ou mais retornos de chamada. Se você precisar fazer isso, é possível usar retornos de chamada aninhados de todos os métodos "Async" da API.

Usar retornos de chamada aninhados é um padrão de programação familiar para a maioria dos desenvolvedores de JavaScript, mas códigos com retornos de chamada profundamente aninhados podem ser difíceis de ler e entender. Como alternativa aos retornos de chamada aninhados, a API JavaScript do Office também dá suporte a uma implementação do padrão de promessas.

> [!NOTE]
> Na versão atual da API JavaScript do *Office, o* suporte interno para o padrão de promessas funciona apenas com código para associações em [planilhas do Excel e documentos do Word](bind-to-regions-in-a-document-or-spreadsheet.md). No entanto, você pode encapsular outras funções que têm retornos de chamada dentro de sua própria função personalizada de retorno de promessa. Para obter mais informações, consulte [Encapsular APIs comuns em funções de retorno de promessa](#wrap-common-apis-in-promise-returning-functions).

### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programação assíncrona usando funções aninhadas de retorno de chamada

Frequentemente, você precisa executar duas ou mais operações assíncronas para concluir uma tarefa. Para fazer isso, você pode aninhar uma chamada "Async" dentro de outra.

O exemplo de código a seguir aninha duas ou mais chamadas assíncronas.

- Primeiro, o método [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) é chamado para acessar uma associação no documento chamado "MyBinding". O `AsyncResult` objeto retornado ao parâmetro `result` desse retorno de chamada fornece acesso ao objeto de associação especificado da `AsyncResult.value` propriedade.
- Em seguida, o objeto de associação acessado do primeiro `result` parâmetro é usado para chamar o [método Binding.getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) .
- Por fim, `result2` o parâmetro do retorno de chamada passado para o `Binding.getDataAsync` método é usado para exibir os dados na associação.

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

Esse padrão de retorno de chamada aninhado básico pode ser usado para todos os métodos assíncronos na API JavaScript do Office.

As seções a seguir mostram como usar funções anônimas ou nomeadas para retornos de chamada aninhados em métodos assíncronos.

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>Usar funções anônimas para retornos de chamada aninhados

No exemplo a seguir, duas funções anônimas são declaradas embutidas `getByIdAsync` `getDataAsync` e passadas para os métodos e como retornos de chamada aninhados. Como as funções são simples e embutidas, a intenção da implementação fica imediatamente clara.

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

#### <a name="use-named-functions-for-nested-callbacks"></a>Usar funções nomeadas para retornos de chamada aninhados

Em implementações complexas, pode ser útil usar funções nomeadas para facilitar a leitura, manutenção e reutilização do seu código. No exemplo a seguir, as duas funções anônimas do exemplo na seção anterior foram reescritas como funções nomeadas e `deleteAllData` `showResult`. Essas funções nomeadas são então passadas para os métodos `getByIdAsync` e `deleteAllDataValuesAsync` como retornos de chamada por nome.

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

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>Programação assíncrona usando o padrão de promessas para acessar dados em associações

Em vez de transmitir a função de retorno de chamada e aguardar até que a função retorne antes da continuação da execução, o padrão de programação de promessas retorna imediatamente retorna um objeto de promessa que representa o resultado desejado. No entanto, ao contrário da verdadeira programação síncrona, nos bastidores o cumprimento do resultado prometido é, na verdade, adiado até que o ambiente de tempo de execução dos Suplementos do Office possa concluir a solicitação. Um manipulador *onError* é fornecido para atender a situações em que a solicitação não pode ser cumprida.

A API JavaScript do Office fornece a [função Office.select](/javascript/api/office#Office_select_expression__callback_) para dar suporte ao padrão de promessas para trabalhar com objetos de associação existentes. O objeto promise `Office.select` retornado à função dá suporte apenas aos quatro métodos que você pode acessar diretamente do objeto [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)), [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)), [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) e [removeHandlerAsync](/javascript/api/office/office.binding#office-office-binding-removehandlerasync-member(1)).

O padrão de promessas para trabalhar com associações assume esse formato.

**Office.select(**_selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_

O *parâmetro selectorExpression* `"bindings#bindingId"`assume o formato , em que *bindingId* é o nome ( `id`) de uma associação que você criou anteriormente no documento ou planilha (usando um dos métodos "addFrom" `Bindings` da coleção: `addFromNamedItemAsync`, ou `addFromPromptAsync``addFromSelectionAsync`). Por exemplo, a expressão do seletor `bindings#cities` especifica que você deseja acessar a associação com uma **ID** de "cidades".

O *parâmetro onError* é uma função de tratamento de erro que usa um único parâmetro do tipo que pode ser usado para acessar um objeto, `select` se a função não conseguir acessar a `AsyncResult` `Error` associação especificada. O exemplo a seguir mostra uma função de manipulador de erro básica que pode ser transmitida para o parâmetro *onError*.

```js
function onError(result){
    const err = result.error;
    write(err.name + ": " + err.message);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Substitua *o espaço reservado BindingObjectAsyncMethod* por uma chamada para `Binding` qualquer um dos quatro métodos de objeto compatíveis com o objeto promise: `getDataAsync`, `setDataAsync`, , `addHandlerAsync`ou `removeHandlerAsync`. As chamadas para esses métodos não oferecem suporte a promessas adicionais. Você deve chamá-los usando o [padrão de função de retorno de chamada aninhado](#asynchronous-programming-using-nested-callback-functions).

`Binding` Depois que uma promessa de objeto é atendida, ela pode ser reutilizada na chamada de método encadeada como se fosse uma associação (o runtime do suplemento não tentará novamente de forma assíncrona cumprir a promessa). Se a `Binding` promessa de objeto não puder ser atendida, o runtime do suplemento tentará novamente acessar o objeto de associação na próxima vez que um de seus métodos assíncronos for invocado.

`select` O exemplo de código a seguir usa a `id` função para recuperar uma associação com o "`cities`" `Bindings` da coleção e, em seguida, chama o método [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) para adicionar um manipulador de eventos para o evento [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) da associação.

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> A `Binding` promessa de objeto retornada `Office.select` pela função fornece acesso apenas aos quatro métodos do `Binding` objeto. Se você precisar acessar qualquer um dos outros `Binding` membros do objeto, `Bindings.getAllAsync` `Document.bindings` `Bindings.getByIdAsync` deverá usar a propriedade e os métodos para recuperar o `Binding` objeto. Por exemplo, `Binding` se você precisar acessar qualquer uma das propriedades do objeto ( `document`as , `id`ou `type` propriedades) ou precisar acessar as propriedades dos objetos [MatrixBinding](/javascript/api/office/office.matrixbinding) ou [TableBinding](/javascript/api/office/office.tablebinding) , `getByIdAsync` `getAllAsync` `Binding` deverá usar os métodos ou os métodos para recuperar um objeto.

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>Passar parâmetros opcionais para métodos assíncronos

A sintaxe comum para todos os métodos "Async" segue esse padrão.

 *AsyncMethod* `(`*RequiredParameters*`, [`*OptionalParameters*`],`*CallbackFunction*`);`

Todos os métodos assíncronos dão suporte a parâmetros opcionais, que são passados como um objeto JavaScript que contém um ou mais parâmetros opcionais. O objeto que contém os parâmetros opcionais é uma coleção não ordenada de pares chave-valor com o caractere ":" separando a chave e o valor. Cada par do objeto é separado por vírgula e o conjunto completo de pares é incluído entre chaves. A chave é o nome do parâmetro e o valor é o valor a ser transmitido para esse parâmetro.

Você pode criar o objeto que contém parâmetros opcionais embutidos ou `options` criando um objeto e passando-o como o *parâmetro de* opções.

### <a name="pass-optional-parameters-inline"></a>Passar parâmetros opcionais embutidos

Por exemplo, a sintaxe para chamar o método [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) com parâmetros opcionais embutidos tem esta aparência:

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

Nessa forma da sintaxe de chamada, os dois parâmetros opcionais, *coercionType* e *asyncContext*, são definidos como um objeto JavaScript anônimo embutido entre chaves.

O exemplo a seguir mostra como chamar o método `Document.setSelectedDataAsync` especificando parâmetros opcionais embutidos.

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
> Você pode especificar parâmetros opcionais em qualquer ordem no objeto de parâmetro, desde que seus nomes sejam especificados corretamente.

### <a name="pass-optional-parameters-in-an-options-object"></a>Passar parâmetros opcionais em um objeto options

Como alternativa, você pode criar `options` um objeto chamado que especifica os parâmetros opcionais separadamente da chamada de método e, em seguida, `options` passar o objeto como o *argumento de* opções.

O exemplo a seguir mostra uma maneira de `options` criar o objeto, `parameter1`em que , `value1`e assim por diante, são espaços reservados para os valores e nomes de parâmetro reais.

```js
const options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};
```

Que é semelhante ao exemplo a seguir quando usado para especificar os parâmetros [ValueFormat](/javascript/api/office/office.valueformat) e [FilterType](/javascript/api/office/office.filtertype).

```js
const options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Aqui está outra maneira de criar o `options` objeto.

```js
const options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Que se parece com o exemplo a seguir quando usado para especificar os `ValueFormat` parâmetros `FilterType` e os parâmetros:

```js
const options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> Ao usar qualquer método de criação do `options` objeto, você pode especificar parâmetros opcionais em qualquer ordem, desde que seus nomes sejam especificados corretamente.

O exemplo a seguir mostra como chamar o método `Document.setSelectedDataAsync` especificando parâmetros opcionais em um `options` objeto.

```js
const options = {
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

Em ambos os exemplos de parâmetro opcionais, o parâmetro de retorno de chamada é especificado como o último parâmetro (seguindo os parâmetros opcionais embutidos ou seguindo o objeto *de* argumento options). Como alternativa, você pode especificar o parâmetro *de retorno* de chamada dentro do objeto JavaScript embutido ou no `options` objeto. No entanto, você pode passar o parâmetro *de* retorno de chamada em apenas um local: `options` no objeto (embutido ou criado externamente) ou como o último parâmetro, mas não em ambos.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Encapsular APIs comuns em funções de retorno de promessa

Os métodos comuns de API (e API do Outlook) não retornam [Promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Portanto, você não pode usar [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída. Se precisar de `await` comportamento, você poderá encapsular a chamada de método em uma Promessa criada explicitamente.

O padrão básico é criar um método *assíncrono* que retorna um objeto Promise imediatamente e resolve esse objeto Promise quando o método interno é concluído ou rejeita o objeto  se o método falhar. Apresentamos um exemplo simples a seguir.

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

Quando essa função precisa ser aguardada, ela pode `await` ser chamada com a palavra-chave ou passada para uma `then` função.

> [!NOTE]
> Essa técnica é especialmente útil quando você precisa chamar uma API `run` Comum dentro de uma chamada da função em um modelo de objeto específico do aplicativo. Para obter um exemplo da `getDocumentFilePath` função que está sendo usada dessa maneira, consulte o arquivo [Home.js exemplo Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).

A seguir está um exemplo usando TypeScript.

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

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
