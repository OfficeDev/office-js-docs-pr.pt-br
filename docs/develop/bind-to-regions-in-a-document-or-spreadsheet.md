---
title: Associar a regi?es em um documento ou em uma planilha
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd26aa12e5d6da145fb6a2a89daf937cf6e88f04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>Associar a regi?es em um documento ou em uma planilha

O acesso a dados baseado em associa??o permite que os suplementos de conte?do e de pain?is de tarefas acessem determinada regi?o de um documento ou planilha por meio de um identificador. Primeiro, o suplemento precisa estabelecer a associa??o. Para isso, ele chama um dos m?todos que associa uma parte do documento a um identificador exclusivo: [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync]. Depois que a associa??o ? estabelecida, o suplemento pode usar o identificador fornecido para acessar os dados contidos na regi?o associada do documento ou da planilha. A cria??o de associa??es proporciona o seguinte valor para o seu suplemento:


- Permite o acesso a estruturas comuns de dados em aplicativos compat?veis do Office, como: tabelas, intervalos ou texto (uma execu??o cont?gua de caracteres).
    
- Habilita opera??es de leitura/grava??o sem exigir que o usu?rio realize uma sele??o.
    
- Estabelece uma rela??o entre o suplemento e os dados presentes no documento. As associa??es est?o presentes no documento e podem ser acessadas em um momento posterior.
    
A cria??o de uma associa??o tamb?m permite que voc? se inscreva em eventos de altera??o de sele??o e de dados que apresentem um escopo definido para essa regi?o espec?fica do documento ou da planilha. Isso significa que o suplemento s? ? notificado sobre altera??es que ocorrem dentro da regi?o associada, e n?o sobre altera??es gerais que ocorrem em todo o documento ou planilha.

O objeto [Bindings] exp?e um m?todo [getAllAsync], que d? acesso ao conjunto de todas as associa??es estabelecidas no documento ou na planilha. Uma associa??o individual pode ser acessada por sua ID, usando o m?todo Bindings.[getByIdAsync] ou [Office.select]. Voc? pode estabelecer novas associa??es e remover as existentes usando um dos seguintes m?todos do objeto [Bindings]: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].


## <a name="binding-types"></a>Tipos de associa??o

H? [tr?s tipos diferentes de associa??es][Office.BindingType] que podem ser especificadas com o par?metro _bindingType_ ao criar uma associa??o com os m?todos [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync]:

1. **[Text Binding][TextBinding]**: associa a uma regi?o do documento que pode ser representada como texto.

    No Word, a maioria das sele??es cont?guas s?o v?lidas, enquanto no Excel apenas as sele??es de c?lulas ?nicas podem ser usadas para uma associa??o de texto. No Excel, s? h? suporte para texto sem formata??o. No Word, h? suporte para tr?s formatos: texto sem formata??o, HTML e Open XML do Office.

2. **[Matrix Binding][MatrixBinding]**: associa uma regi?o fixa de um documento que cont?m dados tabulares sem cabe?alhos. Os dados em uma associa??o de matriz s?o gravados ou lidos como uma **Array** bidimensional, que ? implementada no JavaScript como uma matriz de matrizes. Por exemplo, duas linhas de valores da  **cadeia de caracteres** em duas colunas podem ser gravadas ou lidas como ` [['a', 'b'], ['c', 'd']]` e uma ?nica coluna de tr?s linhas pode ser gravada ou lida como  `[['a'], ['b'], ['c']]`.

    No Excel, qualquer sele??o cont?gua de c?lulas pode ser usada para estabelecer uma associa??o de matriz. No Word, apenas as tabelas d?o suporte ? associa??o de matriz.

3. **[Table Binding][TableBinding]**: associa uma regi?o de um documento que cont?m uma tabela com cabe?alhos. Os dados em uma associa??o de tabela s?o gravados ou lidos como um objeto [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). O objeto `TableData` exp?e os dados por meio das propriedades `headers` e `rows`.

    Qualquer tabela do Excel ou Word pode ser a base para uma associa??o de tabela. Ap?s estabelecer uma associa??o de tabelas, as linhas ou colunas novas que um usu?rio adicionar ? tabela s?o automaticamente inclu?das na associa??o.

Depois que uma associa??o ? criada usando um dos tr?s m?todos "addFrom" do objeto `Bindings` ? poss?vel trabalhar com dados e as propriedades da associa??o usando os m?todos do objeto correspondente: [MatrixBinding], [TableBinding] ou [TextBinding]. Esses tr?s objetos herdam os m?todos  [getDataAsync] e [setDataAsync] do objeto `Binding`, o que permite interagir com os dados associados.

> [!NOTE]
> **Quando devo usar a matriz ou as associa??es de tabela?** Quando os dados tabulares com os quais voc? est? trabalhando contiverem uma linha de totais, voc? dever? usar uma associa??o de matriz se o script do suplemento precisar acessar valores na linha de totais ou detectar que a sele??o do usu?rio est? na linha de totais. Se voc? estabelecer uma associa??o de tabela para os dados tabulares que cont?m uma linha de totais, a propriedade [TableBinding.rowCount] e as propriedades `rowCount` and `startRow` do objeto [BindingSelectionChangedEventArgs] nos manipuladores de eventos n?o refletir?o a linha de totais em seus valores. Para resolver essa limita??o, voc? deve estabelecer uma associa??o de matriz para trabalhar com a linha de totais.

## <a name="add-a-binding-to-the-users-current-selection"></a>Adicionar uma associa??o ? sele??o atual do usu?rio

O exemplo a seguir mostra como adicionar uma associa??o de texto chamada `myBinding` ? sele??o atual em um documento usando o m?todo [addFromSelectionAsync].


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

Neste exemplo, o tipo de associa??o especificado ? texto. Isso significa que um [TextBinding] ser? criado para a sele??o. Diferentes tipos de associa??o exp?em dados e opera??es diferentes. [Office.BindingType] ? uma enumera??o de valores de tipos de associa??es dispon?veis.

O segundo par?metro opcional ? um objeto que especifica a ID da nova associa??o que est? sendo criada. Se uma ID n?o for especificada, uma ser? gerada automaticamente.

A fun??o an?nima transmitida para a fun??o como o par?metro _callback_ final ? executada quando a cria??o da associa??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que fornece acesso a um objeto [AsyncResult] que fornece o status da chamada. A propriedade `AsyncResult.value` cont?m uma refer?ncia para um objeto [Binding] do tipo especificado para a associa??o rec?m-criada. Voc? pode usar esse objeto [Binding] para obter e definir os dados.

## <a name="add-a-binding-from-a-prompt"></a>Adicionar uma associa??o a partir de um prompt

O exemplo a seguir mostra como adicionar uma associa??o de texto chamada `myBinding` usando o m?todo [addFromPromptAsync]. Este m?todo permite ao usu?rio especificar o intervalo da associa??o usando o prompt de sele??o de intervalo interno do aplicativo.


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

Neste exemplo, o tipo de associa??o especificado ? texto. Isso significa que um [TextBinding] ser? criado para a sele??o que o usu?rio especificar no prompt.

O segundo par?metro ? um objeto que cont?m a ID da nova associa??o que est? sendo criada. Se uma ID n?o for especificada, uma ser? gerada automaticamente.

A fun??o an?nima transmitida para a fun??o como o terceiro par?metro _callback_ ? executada quando a cria??o da associa??o ? conclu?da. Quando a fun??o de retorno de chamada ? executada, o objeto [AsyncResult] cont?m o status da chamada e a associa??o rec?m-criada.

A Figura 1 mostra o prompt de sele??o do intervalo interno no Excel.


*Figura 1. Selecionar IU de Dados do Excel*

![Selecionar IU de Dados do Excel](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a>Adicionar uma associa??o a um item nomeado


O exemplo a seguir mostra como adicionar uma associa??o ao item nomeado `myRange` existente como uma associa??o de "matriz" usando o m?todo [addFromNamedItemAsync] e atribui a `id` da associa??o como "myMatrix".


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

**Para Excel**, o par?metro `itemName` do m?todo [addFromNamedItemAsync] pode se referir a um intervalo nomeado existente, a um intervalo especificado com o estilo de refer?ncia `A1` `("A1:A3")` ou a uma tabela. Por padr?o, adicionar uma tabela no Excel atribui o nome "Tabela1" ? primeira tabela adicionada, "Tabela2" ? segunda tabela adicionada e assim por diante. Para atribuir um nome significativo para uma tabela na IU do Excel, use a propriedade **Table Name** na guia **Ferramentas da Tabela | Design** da faixa de op??es.


> [!NOTE]
> No Excel, ao especificar uma tabela como um item nomeado, ? preciso qualificar totalmente o nome ao incluir o nome da planilha no nome da tabela neste formato: `"Sheet1!Table1"`.  `"Sheet1!Table1"`

O exemplo a seguir cria uma associa??o no Excel para as tr?s primeiras c?lulas na coluna A (`"A1:A3"`), atribui a ID `"MyCities"` e, em seguida, grava tr?s nomes de cidades ? associa??o.


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

**Para o Word**, o par?metro `itemName` do m?todo [addFromNamedItemAsync] refere-se ? propriedade `Title` de um controle de conte?do `Rich Text`. (N?o ? poss?vel associar controles de conte?do diferentes do controle de conte?do `Rich Text`.)

Por padr?o, um controle de conte?do n?o tem um valor `Title*` atribu?do. Para atribuir um nome significativo na IU do Word, ap?s inserir um controle de conte?do **Rich Text** do grupo **Controles** na guia **Desenvolvedor** da faixa de op??es, use o comando **Propriedades** no grupo **Controles** para exibir a caixa de di?logo **Propriedades de Controle do Conte?do**. Em seguida, defina a propriedade **Title** do controle de conte?do para o nome que voc? deseja referenciar a partir de seu c?digo.

O exemplo a seguir cria uma associa??o de texto no Word para um controle de conte?do de rich text denominado `"FirstName"`, atribui a **id**`"firstName"` e, em seguida, exibe essas informa??es.


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

## <a name="get-all-bindings"></a>Obter todas as associa??es


O exemplo a seguir mostra como obter todas as associa??es em um documento usando o m?todo Bindings.[getAllAsync].


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

A fun??o an?nima transmitida para a fun??o como o par?metro `callback` ? executada quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que cont?m uma matriz das associa??es no documento. A matriz ? repetida para compilar uma cadeia de caracteres contendo as IDs das associa??es. A cadeia de caracteres ?, ent?o, exibida em uma caixa de mensagem.


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>Obter uma associa??o por ID usando o m?todo getByIdAsync do objeto Bindings


O exemplo a seguir mostra como usar o m?todo [getByIdAsync] para obter uma associa??o em um documento ao especificar sua ID. Este exemplo sup?e que uma associa??o nomeada `'myBinding'` foi adicionada ao documento usando um dos m?todos descritos anteriormente neste t?pico.


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

No exemplo, o primeiro par?metro `id` ? a ID da associa??o a recuperar.

A fun??o an?nima que ? transmitida para a fun??o como o segundo par?metro _callback_ ? executada quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, _asyncResult_, que cont?m o status da chamada e a associa??o com a ID "myBinding".


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>Obter uma associa??o pela ID usando o m?todo selecionado do objeto Office


O exemplo a seguir mostra como usar o m?todo [Office.select] para obter a promessa de um objeto [Binding] em um documento especificando sua ID em uma cadeia de caracteres do seletor. Em seguida, chama o m?todo Binding.[getDataAsync] para obter os dados na associa??o especificada. Este exemplo sup?e que uma associa??o denominada `'myBinding'` foi adicionada ao documento usando um dos m?todos descritos anteriormente neste t?pico.


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
> Se a promessa do m?todo `select` retornar um objeto [Binding] com ?xito, esse objeto ir? expor somente os seguintes quatro m?todos do objeto: [getDataAsync], [setDataAsync], [addHandlerAsync], e [removeHandlerAsync]. Se a promessa n?o puder retornar um objeto Binding, o retorno de chamada `onError` pode ser usado para acessar um objeto [asyncResult].error para mais informa??es. Se for preciso chamar um membro do objeto Binding diferente dos quatro m?todos expostos pela promessa do objeto Binding retornada pelo m?todo `select`, use o m?todo [getByIdAsync] utilizando a propriedade [Document.bindings] e o m?todo Bindings.[getByIdAsync] para recuperar o objeto Binding**.

## <a name="release-a-binding-by-id"></a>Liberar uma associa??o pela ID


O exemplo a seguir mostra como usar o m?todo [releaseByIdAsync] para liberar uma associa??o em um documento, especificando sua ID.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

No exemplo, o primeiro par?metro `id` ? a ID da associa??o a liberar.

A fun??o an?nima que ? transmitida para a fun??o como o segundo par?metro ? um retorno de chamada executado quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, [asyncResult], que cont?m o status da chamada.


## <a name="read-data-from-a-binding"></a>Ler os dados de uma associa??o


O exemplo a seguir mostra como usar o m?todo [getDataAsync] para obter dados de uma associa??o existente.


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

 `myBinding` ? uma vari?vel que cont?m uma associa??o de texto existente no documento. Como alternativa, ? poss?vel usar [Office.select] para acessar a associa??o pela ID, e iniciar sua chamada para o m?todo [getDataAsync], assim: 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


A fun??o an?nima transmitida para a fun??o ? um retorno de chamada executado quando a opera??o ? conclu?da. A propriedade [AsyncResult].value cont?m os dados em `myBinding`. O tipo do valor depende do tipo de associa??o. A associa??o neste exemplo ? uma associa??o de texto. Portanto, o valor conter? uma cadeia de caracteres. Para obter mais exemplos de como trabalhar com as associa??es de tabela e matriz, confira o t?pico do m?todo [getDataAsync].


## <a name="write-data-to-a-binding"></a>Gravar dados em uma associa??o

O exemplo a seguir mostra como usar o m?todo [setDataAsync] para definir os dados em uma associa??o existente.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` ? uma vari?vel que cont?m uma associa??o de texto existente no documento.

No exemplo, o primeiro par?metro ? o valor a definir em `myBinding`. Como esta ? uma associa??o de texto, o valor ? uma `string`. Diferentes tipos de associa??o aceitam diferentes tipos de dados.

A fun??o an?nima que ? transmitida para a fun??o ? um retorno de chamada executado quando a opera??o ? conclu?da. A fun??o ? chamada com um ?nico par?metro, `asyncResult`, que cont?m o status do resultado.

> [!NOTE]
> A partir da vers?o do Excel 2013 SP1 e da compila??o correspondente do Excel Online, agora ? poss?vel [definir a formata??o ao escrever e atualizar dados em tabelas de vincula??o](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>Detectar altera??es nos dados ou a sele??o em uma associa??o


O exemplo a seguir mostra como anexar um manipulador de eventos ao evento [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) de uma associa??o com uma id "MyBinding".


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

? uma vari?vel que cont?m uma associa??o de texto existente no documento.`myBinding`

O primeiro par?metro `eventType` do m?todo [addHandlerAsync] especifica o nome do evento no qual se inscrever. [Office.EventType] ? uma enumera??o dos valores do tipo de evento dispon?veis. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"`.

A fun??o `dataChanged` que ? transmitida para a fun??o como o segundo par?metro _handler_ ? um manipulador de eventos executado quando os dados na associa??o s?o alterados. A fun??o ? chamada com um ?nico par?metro, _eventArgs_, que cont?m uma refer?ncia para a associa??o. Essa associa??o pode ser usada para recuperar os dados atualizados.

Da mesma forma, ? poss?vel detectar quando um usu?rio altera a sele??o em uma associa??o anexando um manipulador de eventos ao evento [SelectionChanged] de uma associa??o. Para fazer isso, especifique o par?metro `eventType` do m?todo [addHandlerAsync] como `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.

Voc? pode adicionar v?rios manipuladores de eventos para um determinado evento chamando o m?todo [addHandlerAsync] novamente e transmitindo uma fun??o do manipulador de eventos adicional para o par?metro `handler`. Isso funcionar? corretamente, contanto que o nome de cada fun??o do manipulador de eventos seja exclusivo.


### <a name="remove-an-event-handler"></a>Remover um manipulador de eventos


Para remover um manipulador de eventos de um evento, chame o m?todo [removeHandlerAsync] passando o tipo de evento como o primeiro par?metro _eventType_ e o nome da fun??o do manipulador de eventos a remover como o segundo par?metro _handler_. Por exemplo, a fun??o a seguir remover? a fun??o de manipulador de eventos `dataChanged` adicionada no exemplo da se??o anterior.


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> Se o par?metro opcional _handler_ for omitido ao chamar o m?todo [removeHandlerAsync], todos os manipuladores de eventos do `eventType` especificado ser?o removidos.


## <a name="see-also"></a>Veja tamb?m

- [No??es b?sicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md) 
- [Programa??o ass?ncrona nos Suplementos do Office](asynchronous-programming-in-office-add-ins.md)
- [Leia e grave dados na sele??o ativa, em um documento ou em uma planilha](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Associa??o]:               https://dev.office.com/reference/add-ins/shared/binding
[MatrixBinding]:         https://dev.office.com/reference/add-ins/shared/binding.matrixbinding
[TableBinding]:          https://dev.office.com/reference/add-ins/shared/binding.tablebinding
[TextBinding]:           https://dev.office.com/reference/add-ins/shared/binding.textbinding
[getDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.getdataasync
[setDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.setdataasync
[SelectionChanged]:      https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent
[addHandlerAsync]:       https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync
[removeHandlerAsync]:    https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync

[Associa??es]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[getByIdAsync]:          https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync 
[getAllAsync]:           https://dev.office.com/reference/add-ins/shared/bindings.getallasync
[addFromNamedItemAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync
[addFromSelectionAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync
[addFromPromptAsync]:    https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync
[releaseByIdAsync]:      https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync

[AsyncResult]:          https://dev.office.com/reference/add-ins/shared/asyncresult
[Office.BindingType]:   https://dev.office.com/reference/add-ins/shared/bindingtype-enumeration
[Office.select]:        https://dev.office.com/reference/add-ins/shared/office.select 
[Office.EventType]:     https://dev.office.com/reference/add-ins/shared/eventtype-enumeration 
[Document.bindings]:    https://dev.office.com/reference/add-ins/shared/document.bindings


[TableBinding.rowCount]: https://dev.office.com/reference/add-ins/shared/binding.tablebinding.rowcount
[BindingSelectionChangedEventArgs]: https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedeventargs
