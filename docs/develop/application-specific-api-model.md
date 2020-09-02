---
title: Usando o modelo de API específico do aplicativo
description: Saiba mais sobre o modelo de API baseado em promessa para os suplementos do Excel, OneNote e Word.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: cabd1ea0076b672a1dbda3079a767b0e8a1a62b7
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326279"
---
# <a name="using-the-application-specific-api-model"></a>Usando o modelo de API específico do aplicativo

Este artigo descreve como usar o modelo de API para criar suplementos no Excel, no Word e no OneNote. Ele apresenta os principais conceitos fundamentais para o uso das APIs baseadas em promessa.

> [!NOTE]
> Não há suporte para esse modelo nos clientes do Office 2013. Use o [modelo de API comum](office-javascript-api-object-model.md) para trabalhar com essas versões do Office. Para ver as notas de disponibilidade completa da plataforma, confira [disponibilidade de aplicativos e plataformas do cliente Office para suplementos do Office](../overview/office-add-in-availability.md).

> [!TIP]
> Os exemplos nesta página usam as APIs JavaScript do Excel, mas os conceitos também se aplicam ao OneNote, Visio e APIs JavaScript do Word.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Natureza assíncrona das APIs baseadas em promessa

Os suplementos do Office são sites que aparecem dentro de um contêiner de navegadores em aplicativos do Office, como o Excel. Esse contêiner é incorporado no aplicativo do Office em plataformas baseadas em área de trabalho, como o Office no Windows, e é executado dentro de um iFrame HTML no Office na Web. Devido a considerações de desempenho, as APIs do Office.js não podem interagir de forma síncrona com os aplicativos do Office em todas as plataformas. Portanto, a `sync()` chamada de API no Office.js retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) resolvida quando o aplicativo do Office conclui as ações de leitura ou gravação solicitadas. Além disso, você pode enfileirar várias ações, como definir propriedades ou invocar métodos, e executá-las como um lote de comandos com uma única chamada para `sync()` , em vez de enviar uma solicitação separada para cada ação. As seções a seguir descrevem como fazer isso usando as `run()` `sync()` APIs e.

## <a name="run-function"></a>função *. Run

`Excel.run`, `Word.run` e `OneNote.run` Execute uma função que especifica as ações a serem executadas em relação ao Excel, Word e OneNote. `*.run` cria automaticamente um contexto de solicitação que você pode usar para interagir com objetos do Office. Quando `*.run` é concluído, uma promessa é resolvida e todos os objetos que foram alocados no tempo de execução são automaticamente liberados.

O exemplo a seguir mostra como usar o `Excel.run` . O mesmo padrão também é usado com o Word e o OneNote.

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a>Contexto de solicitação

O aplicativo do Office e seu suplemento são executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos exigem um `RequestContext` objeto para conectar seu suplemento a objetos no Office, como planilhas, intervalos, parágrafos e tabelas. Esse `RequestContext` objeto é fornecido como um argumento ao chamar `*.run` .

## <a name="proxy-objects"></a>Objetos proxy

Os objetos JavaScript do Office que você declara e usa com as APIs baseadas em promessa são objetos de proxy. Todos os métodos invocados, ou as propriedades definidas ou carregadas em objetos proxy são simplesmente adicionados a uma fila de comandos pendentes. Quando você chama o `sync()` método no contexto de solicitação (por exemplo, `context.sync()` ), os comandos enfileirados são expedidos para o aplicativo do Office e executados. Essas APIs são essencialmente centradas em lote. Você pode enfileirar quantas alterações desejar no contexto da solicitação e, em seguida, chamar o `sync()` método para executar o lote de comandos enfileirados.

Por exemplo, o trecho de código a seguir declara o objeto JavaScript [Excel. Range](/javascript/api/excel/excel.range) local, `selectedRange` para fazer referência a um intervalo selecionado na pasta de trabalho do Excel e, em seguida, define algumas propriedades nesse objeto. O `selectedRange` objeto é um objeto proxy, portanto, as propriedades que são definidas e o método invocado nesse objeto não serão refletidas no documento do Excel até que seu suplemento chame `context.sync()` .

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Dica de desempenho: minimizar o número de objetos de proxy criados

Evite criar repetidamente o mesmo objeto proxy. Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a>sync()

Chamar o `sync()` método no contexto de solicitação sincroniza o estado entre objetos de proxy e objetos no documento do Office. O `sync()` método executa todos os comandos que estão na fila no contexto de solicitação e recupera valores para todas as propriedades que devem ser carregadas nos objetos de proxy. O `sync()` método é executado de forma assíncrona e retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o `sync()` método é concluído.

O exemplo a seguir mostra uma função em lotes que define um objeto de proxy JavaScript local ( `selectedRange` ), carrega uma propriedade desse objeto e, em seguida, usa o padrão de promessas do JavaScript a ser chamado `context.sync()` para sincronizar o estado entre objetos proxy e objetos no documento do Excel.

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

No exemplo anterior, `selectedRange` é definido e sua propriedade `address` é carregada quando `context.sync()` é chamado.

Como `sync()` é uma operação assíncrona, você sempre deve retornar o `Promise` objeto para garantir que a `sync()` operação seja concluída antes de o script continuar a ser executado. Se você estiver usando o TypeScript ou ES6 + JavaScript, você `await` poderá `context.sync()` chamar em vez de retornar a promessa.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Dica de desempenho: minimizar o número de chamadas de sincronização

Na API do JavaScript do Excel, `sync()` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online na Web. Para otimizar o desempenho, minimize o número de chamadas para `sync()`, enfileirando o maior número possível de alterações antes de chamá-lo. Para obter mais informações sobre como otimizar `sync()` o desempenho do, consulte [Evite usar o método Context. Sync em loops](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Antes de poder ler as propriedades de um objeto proxy, você deve carregar explicitamente as propriedades para preencher o objeto proxy com dados do documento do Office e, em seguida, chamar `context.sync()` . Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e, em seguida, quiser ler a propriedade do intervalo selecionado `address` , você precisará carregar a `address` propriedade antes de poder lê-la. Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o `load()` método no objeto e especifique as propriedades a serem carregadas. O exemplo a seguir mostra a `Range.address` propriedade que está sendo carregada `myRange` .

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> Se você estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o `load()` método. O `load()` método só é necessário quando você deseja ler propriedades em um objeto proxy.

Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto de solicitação, sendo executadas na próxima vez que você chamar o método `sync()`. É possível enfileirar quantas chamadas de `load()` forem necessárias no contexto de solicitação.

#### <a name="scalar-and-navigation-properties"></a>Propriedades escalares e de navegação

Há duas categorias de propriedades: **escalar** e de **navegação**. As propriedades escalares são tipos atribuíveis, como cadeias de caracteres, inteiros e estruturas JSON. As propriedades de navegação são objetos somente leitura e coleções de objetos que têm seus campos atribuídos, em vez de atribuir diretamente a propriedade. Por exemplo, `name` e `position` os membros do objeto [Excel. Worksheet](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são propriedades de navegação.

O suplemento pode usar propriedades de navegação como um caminho para carregar Propriedades escalares específicas. O código a seguir enfileira um `load` comando para o nome da fonte usada por um `Excel.Range` objeto, sem carregar nenhuma outra informação.

```js
someRange.load("format/font/name")
```

Você também pode definir as propriedades escalares de uma propriedade de navegação atravessando o caminho. Por exemplo, você pode definir o tamanho da fonte de um `Excel.Range` usando `someRange.format.font.size = 10;` . Você não precisa carregar a propriedade antes de defini-la.

Observe que algumas das propriedades em um objeto podem ter o mesmo nome de outro objeto. Por exemplo, `format` é uma propriedade sob o `Excel.Range` objeto, mas `format` também é um objeto. Portanto, se você fizer uma chamada como `range.load("format")` , isso equivale a `range.format.load()` (uma instrução vazia indesejável `load()` ). Para evitar isso, o código só deve carregar os "nós folha" em uma árvore de objetos.

#### <a name="calling-load-without-parameters-not-recommended"></a>Chamar `load` sem parâmetros (não recomendado)

Se você chamar o `load()` método em um objeto (ou coleção) sem especificar nenhum parâmetro, todas as propriedades escalares do objeto ou dos objetos da coleção serão carregadas. O carregamento de dados desnecessários tornará o suplemento lento. Você sempre deve especificar explicitamente as propriedades a serem carregadas.

> [!IMPORTANT]
> A quantidade de dados retornados por uma declaração `load` sem parâmetros pode exceder os limites de tamanho do serviço. Para reduzir os riscos a suplementos mais antigos, algumas propriedades não são retornadas por `load` sem a solicitação explícita. As seguintes propriedades são excluídas dessas operações de carregamento:
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Os métodos nas APIs baseadas em promessa que retornam tipos primitivos têm um padrão semelhante ao `load` / `sync` paradigma. Por exemplo, `Excel.TableCollection.getCount` obtém o número de tabelas da coleção. `getCount` Retorna um `ClientResult<number>` , significando que a `value` propriedade no retornado [`ClientResult`](/javascript/api/office/officeextension.clientresult) é um número. Seu script não pode acessar esse valor até que `context.sync()` seja chamado.

O código a seguir obtém o número total de tabelas em uma pasta de trabalho do Excel e registra esse número no console.

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a>set()

A definição de propriedades em um objeto com propriedades de navegação aninhadas pode ser uma tarefa complicada. Como alternativa à definição de propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o `object.set()` método que está disponível em objetos nas APIs JavaScript baseadas em promessa. Com esse método, é possível definir várias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que são estruturadas, como as propriedades do objeto no qual o método é chamado.

O exemplo de código a seguir define várias propriedades do formato de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto `Range`. Este exemplo supõe que há dados no intervalo **B2:E2**.

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods-and-properties"></a>Métodos e propriedades do &#42;OrNullObject

Alguns métodos e propriedades de assessor geram uma exceção quando o objeto desejado não existe. Por exemplo, se você tentar obter uma planilha do Excel especificando um nome de planilha que não esteja na pasta de trabalho, o `getItem()` método gera uma `ItemNotFound` exceção. As bibliotecas específicas do aplicativo fornecem uma maneira de seu código testar a existência de entidades de documento sem exigir código de tratamento de exceção. Isso é feito usando as `*OrNullObject` variações de métodos e propriedades. Essas variações retornam um objeto cuja `isNullObject` propriedade é definida como `true` , se o item especificado não existir, em vez de gerar uma exceção.

Por exemplo, você pode chamar o `getItemOrNullObject()` método em uma coleção como **planilhas** para recuperar um item da coleção. O `getItemOrNullObject()` método retorna o item especificado se ele existir; caso contrário, retorna um objeto cuja `isNullObject` propriedade está definida como `true` . Seu código pode então avaliar essa propriedade para determinar se o objeto existe.

> [!NOTE]
> As `*OrNullObject` variações nunca retornam o valor de JavaScript `null` . Eles retornam objetos de proxy do Office comuns. Se a entidade que o objeto representa não existir, a `isNullObject` Propriedade do objeto será definida como `true` . Não teste o objeto retornado para nulidade ou falsity. Ele nunca é `null` , `false` ou `undefined` .

O exemplo de código a seguir tenta recuperar uma planilha do Excel chamada "data" usando o `getItemOrNullObject()` método. Se uma planilha com esse nome não existir, será criada uma nova planilha. Observe que o código não carrega a `isNullObject` propriedade. O Office carrega automaticamente essa propriedade quando `context.sync` é chamado, portanto, você não precisa carregá-la explicitamente com algo como `datasheet.load('isNullObject')` .

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a>Confira também

* [Modelo de objeto comum de API JavaScript para Office](office-javascript-api-object-model.md)
* [Problemas comuns de codificação e comportamentos inesperados da plataforma](common-coding-issues.md).
* [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
