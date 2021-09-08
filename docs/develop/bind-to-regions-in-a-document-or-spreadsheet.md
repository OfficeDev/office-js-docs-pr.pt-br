---
title: Associar a regiões em um documento ou em uma planilha
description: Saiba como usar a associação para garantir o acesso consistente a uma região ou elemento específico de um documento ou planilha por meio de um identificador.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 9db35168274b599b93a6688d1318103c48edee55
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936322"
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>Associe a regiões em um documento ou planilha

O acesso a dados baseados em associação permite que os complementos do painel de tarefas e conteúdo acessem consistentemente uma determinada região de um documento ou planilha por meio de um identificador. Primeiro, o suplemento precisa estabelecer a associação chamando um dos métodos que vinculam uma parte do documento a um identificador exclusivo: [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync]. Depois que a associação é estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na região vinculada do documento ou da planilha. A criação de vinculações fornece o seguinte valor ao seu complemento.

- Permite o acesso a estruturas comuns de dados em aplicativos compatíveis do Office, como: tabelas, intervalos ou texto (uma execução contígua de caracteres).
- Habilita operações de leitura/gravação sem exigir que o usuário realize uma seleção.
- Estabelece uma relação entre o suplemento e os dados presentes no documento. As associações estão presentes no documento e podem ser acessadas em um momento posterior.

A criação de uma associação também permite que você se inscreva em eventos de alteração de seleção e de dados que apresentem um escopo definido para essa região específica do documento ou da planilha. Isso significa que o suplemento só é notificado sobre alterações que ocorrem dentro da região associada, e não sobre alterações gerais que ocorrem em todo o documento ou planilha.

O objeto [Bindings] expõe um método [getAllAsync], que dá acesso ao conjunto de todas as associações estabelecidas no documento ou na planilha. Uma associação individual pode ser acessada por sua ID, usando o método Bindings.[getByIdAsync] ou [Office.select]. Você pode estabelecer novas associações e remover as existentes usando um dos seguintes métodos do objeto [Bindings]: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].

## <a name="binding-types"></a>Tipos de associação

Há três [tipos diferentes de Office.][ BindingType] que você especifica com o parâmetro _bindingType_ ao criar uma associação com os métodos [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync.]

1. **[Text Binding][TextBinding]**: associa a uma região do documento que pode ser representada como texto.

    No Word, a maioria das seleções contíguas são válidas, enquanto no Excel apenas as seleções de células únicas podem ser usadas para uma associação de texto. No Excel, só há suporte para texto sem formatação. No Word, há suporte para três formatos: texto sem formatação, HTML e Open XML do Office.

2. **[Matrix Binding][MatrixBinding]** - Vincula a uma região fixa de um documento que contém dados tabulares sem headers. Os dados em uma associação de matriz são gravados ou lidos como uma **Matriz** bidimensional , que em JavaScript é implementada como uma matriz de matrizes. Por exemplo, duas linhas de valores **string** em duas colunas podem ser gravadas ou lidas como `[['a', 'b'], ['c', 'd']]`, e uma única coluna de três linhas pode ser gravada ou lida como `[['a'], ['b'], ['c']]`.

    No Excel, qualquer seleção contígua de células pode ser usada para estabelecer uma associação de matriz. No Word, apenas as tabelas dão suporte à associação de matriz.

3. **[Table Binding][TableBinding]**: associa uma região de um documento que contém uma tabela com cabeçalhos. Os dados em uma associação de tabela são gravados ou lidos como um objeto [TableData](/javascript/api/office/office.tabledata). O objeto `TableData` expõe os dados por meio das propriedades `headers` e `rows`.

    Qualquer tabela do Excel ou Word pode ser a base para uma associação de tabela. Após estabelecer uma associação de tabelas, as linhas ou colunas novas que um usuário adicionar à tabela são automaticamente incluídas na associação. 

Depois que uma associação é criada usando um dos três métodos "addFrom" do objeto, você pode trabalhar com os dados e propriedades da associação usando os métodos do objeto `Bindings` correspondente: [MatrixBinding], [TableBinding]ou [TextBinding]. Esses três objetos herdam os métodos [getDataAsync] e [setDataAsync] do objeto `Binding`, o que permite interagir com os dados associados.

> [!NOTE]
> **Quando devo usar a matriz ou as associações de tabela?** Quando os dados tabulares com os quais você está trabalhando contiverem uma linha de totais, você deverá usar uma associação de matriz se o script do suplemento precisar acessar valores na linha de totais ou detectar que a seleção do usuário está na linha de totais. Se você estabelecer uma associação de tabela para os dados tabulares que contêm uma linha de totais, a propriedade [TableBinding.rowCount] e as propriedades `rowCount` and `startRow` do objeto [BindingSelectionChangedEventArgs] nos manipuladores de eventos não refletirão a linha de totais em seus valores. Para resolver essa limitação, você deve estabelecer uma associação de matriz para trabalhar com a linha de totais.

## <a name="add-a-binding-to-the-users-current-selection"></a>Adicionar uma associação à seleção atual do usuário

O exemplo a seguir mostra como adicionar uma associação de texto chamada `myBinding` à seleção atual em um documento usando o método [addFromSelectionAsync].

```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Neste exemplo, o tipo de associação especificado é texto. Isso significa que um [TextBinding] será criado para a seleção. Diferentes tipos de associação expõem dados e operações diferentes. [Office.BindingType] é uma enumeração de valores de tipos de associações disponíveis.

O segundo parâmetro opcional é um objeto que especifica a ID da nova associação que está sendo criada. Se uma ID não for especificada, uma será gerada automaticamente.

A função anônima transmitida para a função como o parâmetro _callback_ final é executada quando a criação da associação é concluída. A função é chamada com um único parâmetro, `asyncResult`, que fornece acesso a um objeto [AsyncResult] que fornece o status da chamada. A propriedade `AsyncResult.value` contém uma referência para um objeto [Binding] do tipo especificado para a associação recém-criada. Você pode usar esse objeto [Binding] para obter e definir os dados.

## <a name="add-a-binding-from-a-prompt"></a>Adicionar uma associação a partir de um prompt

O exemplo a seguir mostra como adicionar uma associação de texto chamada `myBinding` usando o método [addFromPromptAsync]. Este método permite ao usuário especificar o intervalo da associação usando o prompt de seleção de intervalo interno do aplicativo.

```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

Neste exemplo, o tipo de associação especificado é texto. Isso significa que um [TextBinding] será criado para a seleção que o usuário especificar no prompt.

O segundo parâmetro é um objeto que contém a ID da nova associação que está sendo criada. Se uma ID não for especificada, uma será gerada automaticamente.

A função anônima passada para a função como o terceiro parâmetro _de retorno_ de chamada é executada quando a criação da associação é concluída. Quando a função de retorno de chamada é executada, o objeto [AsyncResult] contém o status da chamada e a vinculação recém-criada.

A Figura 1 mostra o prompt de seleção do intervalo interno no Excel.

*Figura 1. Selecionar IU de Dados do Excel*

![Captura de tela mostrando a caixa de diálogo Selecionar Dados.](../images/agave-api-overview-excel-selection-ui.png)

## <a name="add-a-binding-to-a-named-item"></a>Adicionar uma associação a um item nomeado

O exemplo a seguir mostra como adicionar uma associação ao item nomeado existente como uma associação de "matriz" usando o `myRange` [método addFromNamedItemAsync] e atribui a associação como `id` "myMatrix".

```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

**Para Excel**, o parâmetro do método `itemName` [addFromNamedItemAsync] pode se referir a um intervalo nomeado existente, um intervalo especificado com o estilo de referência `A1` ou uma `("A1:A3")` tabela. Por padrão, a adição de uma tabela no Excel atribui o nome "Tabela1" à primeira tabela que você adicionar, "Tabela2" para a segunda tabela adicionada e assim por diante. Para atribuir um nome significativo para uma tabela na interface do usuário Excel, use a propriedade no table `Table Name` **Tools | Guia Design** da faixa de opções.

> [!NOTE]
> No Excel, ao especificar uma tabela como um item nomeado, você deve qualificar totalmente o nome para incluir o nome da planilha no nome da tabela neste formato:`"Sheet1!Table1"`

O exemplo a seguir cria uma associação Excel as três primeiras células da coluna A ( ), atribui a id e grava três nomes de cidade a `"A1:A3"` `"MyCities"` essa associação.

```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**Para o Word,** `itemName` o parâmetro do método [addFromNamedItemAsync] refere-se à propriedade `Title` de um controle de `Rich Text` conteúdo. Não é possível associar a controles de conteúdo diferentes de controles de conteúdo `Rich Text`.

Por padrão, um controle de conteúdo não tem `Title*` valor atribuído. Para atribuir um nome significativo na interface do usuário do Word, depois de inserir um controle de conteúdo **Rich Text** do grupo **Controles** na guia **Desenvolvedor** da faixa de opções, use o comando **Propriedades** no grupo **Controles** para exibir a caixa de diálogo **Propriedades de Controle de Conteúdo**. Em `Title` seguida, de definir a propriedade do controle de conteúdo como o nome que você deseja fazer referência do código.

O exemplo a seguir cria uma associação de texto no Word a um controle de conteúdo de texto rico chamado , atribui a `"FirstName"` **id** e exibe `"firstName"` essas informações.

```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a>Obter todas as associações

O exemplo a seguir mostra como obter todas as associações em um documento usando o método Bindings.[getAllAsync].

```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

A função anônima que é passada para a função como `callback` o parâmetro é executada quando a operação é concluída. A função é chamada com um único parâmetro, `asyncResult` , que contém uma matriz das associações no documento. A matriz é repetida para criar uma cadeia de caracteres contendo as IDs das vinculações. A cadeia de caracteres é, então, exibida em uma caixa de mensagem.

## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>Obter uma associação por ID usando o método getByIdAsync do objeto Bindings

O exemplo a seguir mostra como usar o método [getByIdAsync] para obter uma associação em um documento ao especificar sua ID. Este exemplo supõe que uma associação nomeada `'myBinding'` foi adicionada ao documento usando um dos métodos descritos anteriormente neste tópico.

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

No exemplo, o primeiro `id` parâmetro é a ID da associação a ser recuperada.

A função anônima que é passada para a função como o segundo parâmetro _de retorno_ de chamada é executada quando a operação é concluída. A função é chamada com um único parâmetro, _asyncResult_, que contém o status da chamada e as vinculações com a ID "myBinding".

## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>Obter uma associação pela ID usando o método selecionado do objeto Office

O exemplo a seguir mostra como usar o método [Office.select] para obter a promessa de um objeto [Binding] em um documento especificando sua ID em uma cadeia de caracteres do seletor. Em seguida, chama o método Binding.[getDataAsync] para obter os dados na associação especificada. Este exemplo supõe que uma associação denominada `'myBinding'` foi adicionada ao documento usando um dos métodos descritos anteriormente neste tópico.

```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> Se a promessa de método retornar com êxito um objeto Binding, esse objeto exporá apenas os quatro métodos a seguir `select` do objeto: [getDataAsync], [setDataAsync,] [addHandlerAsync]e [removeHandlerAsync]. [] Se a promessa não puder retornar um objeto Binding, o retorno de chamada poderá ser usado para acessar um `onError` objeto .error [asyncResult]para obter mais informações. Se você precisar chamar um membro do objeto Binding diferente dos quatro métodos expostos pela promessa de objeto [Binding] retornada pelo método, use o método `select` [getByIdAsync] usando a propriedade [Document.bindings] e Bindings.[ Método getByIdAsync] para recuperar o [objeto Binding.]

## <a name="release-a-binding-by-id"></a>Liberar uma associação pela ID

O exemplo a seguir mostra como usar o método [releaseByIdAsync] para liberar uma associação em um documento, especificando sua ID.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

No exemplo, o primeiro parâmetro `id` é a ID da associação a liberar.

A função anônima que é transmitida para a função como o segundo parâmetro é um retorno de chamada executado quando a operação é concluída. A função é chamada com um único parâmetro, [asyncResult], que contém o status da chamada.

## <a name="read-data-from-a-binding"></a>Ler os dados de uma associação

O exemplo a seguir mostra como usar o método [getDataAsync] para obter dados de uma associação existente.

```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

`myBinding` é uma variável que contém uma associação de texto existente no documento. Como alternativa, é possível usar [Office.select] para acessar a associação pela ID, e iniciar sua chamada para o método [getDataAsync], assim:

```js
Office.select("bindings#myBindingID").getDataAsync
```

A função anônima transmitida para a função é um retorno de chamada executado quando a operação é concluída. A propriedade [AsyncResult].value contém os dados em `myBinding`. O tipo do valor depende do tipo de associação. A associação neste exemplo é uma associação de texto. Portanto, o valor conterá uma cadeia de caracteres. Para obter mais exemplos de como trabalhar com as associações de tabela e matriz, confira o tópico do método [getDataAsync].

## <a name="write-data-to-a-binding"></a>Gravar dados em uma associação

O exemplo a seguir mostra como usar o método [setDataAsync] para definir os dados em uma associação existente.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

`myBinding` é uma variável que contém uma associação de texto existente no documento.

No exemplo, o primeiro parâmetro é o valor a ser definido em `myBinding` . Como esta é uma associação de texto, o valor é uma `string`. Diferentes tipos de associação aceitam diferentes tipos de dados.

A função anônima que é transmitida para a função é um retorno de chamada executado quando a operação é concluída. A função é chamada com um único parâmetro, `asyncResult` , que contém o status do resultado.

> [!NOTE]
> A partir da versão do Excel 2013 SP1 e da compilação correspondente do Excel Online, agora é possível [definir a formatação ao escrever e atualizar dados em tabelas associadas](../excel/excel-add-ins-tables.md).

## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>Detectar alterações nos dados ou a seleção em uma associação

O exemplo a seguir mostra como anexar um manipulador de eventos ao evento [DataChanged](/javascript/api/office/office.binding) de uma associação com uma id "MyBinding".

```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

`myBinding` é uma variável que contém uma associação de texto existente no documento.

O primeiro _parâmetro eventType_ do [método addHandlerAsync] especifica o nome do evento a ser inscrito. [Office.EventType] é uma enumeração de valores de tipos de eventos disponíveis. `Office.EventType.BindingDataChanged` avalia para a cadeia de caracteres "bindingDataChanged".

A função passada para a função como o segundo parâmetro de manipulador é um manipulador de eventos executado quando os dados na `dataChanged` associação são  alterados. A função é chamada com um único parâmetro, _eventArgs_, que contém uma referência para a vinculação. Essa associação pode ser usada para recuperar os dados atualizados.

Da mesma forma, é possível detectar quando um usuário altera a seleção em uma associação anexando um manipulador de eventos ao evento [SelectionChanged] de uma associação. Para fazer isso, especifique o parâmetro `eventType` do método [addHandlerAsync] como `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.

Você pode adicionar vários manipuladores de eventos para um determinado evento chamando o método [addHandlerAsync] novamente e transmitindo uma função do manipulador de eventos adicional para o parâmetro `handler`. Isso funcionará corretamente, contanto que o nome de cada função do manipulador de eventos seja exclusivo.

### <a name="remove-an-event-handler"></a>Remover um manipulador de eventos

Para remover um manipulador de eventos de um evento, chame o método [removeHandlerAsync] passando o tipo de evento como o primeiro parâmetro _eventType_ e o nome da função do manipulador de eventos a remover como o segundo parâmetro _handler_. Por exemplo, a função a seguir removerá a função de manipulador de eventos `dataChanged` adicionada no exemplo da seção anterior.

```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```

> [!IMPORTANT]
> Se o parâmetro _manipulador_ opcional for omitido quando o [método removeHandlerAsync] for chamado, todos os manipuladores de eventos para o especificado `eventType` serão removidos.

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [Programação assíncrona nos Suplementos do Office](asynchronous-programming-in-office-add-ins.md)
- [Leia e grave dados na seleção ativa, em um documento ou em uma planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Associação]:               /javascript/api/office/office.binding
[MatrixBinding]:         /javascript/api/office/office.matrixbinding
[TableBinding]:          /javascript/api/office/office.tablebinding
[TextBinding]:           /javascript/api/office/office.textbinding
[getDataAsync]:          /javascript/api/office/office.binding#getDataAsync_options__callback_
[setDataAsync]:          /javascript/api/office/office.binding#setDataAsync_data__options__callback_
[SelectionChanged]:      /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]:       /javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_
[removeHandlerAsync]:    /javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_

[Associações]:              /javascript/api/office/office.bindings
[getByIdAsync]:          /javascript/api/office/office.bindings#getByIdAsync_id__options__callback_
[getAllAsync]:           /javascript/api/office/office.bindings#getAllAsync_options__callback_
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_
[addFromPromptAsync]:    /javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_
[releaseByIdAsync]:      /javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_

[AsyncResult]:          /javascript/api/office/office.asyncresult
[Office.BindingType]:   /javascript/api/office/office.bindingtype
[Office.select]:        /javascript/api/office 
[Office.EventType]:     /javascript/api/office/office.eventtype 
[Document.bindings]:    /javascript/api/office/office.document

[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs
