---
title: Usando o modelo de API específica do aplicativo
description: Saiba mais sobre o modelo de API baseada em promessas para suplementos do Excel, do OneNote e do Word.
ms.date: 09/23/2022
ms.localizationpriority: medium
ms.openlocfilehash: d24b435318e1f462cd05ba25dbdd7f9a6018715f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810173"
---
# <a name="application-specific-api-model"></a>Modelo de API específico do aplicativo

Este artigo descreve como usar o modelo de API para criar suplementos no Excel, Word, PowerPoint e OneNote. Ele introduz os conceitos fundamentais do uso de APIs baseadas em promessas.

> [!NOTE]
> Esse modelo não tem suporte para clientes do Office 2013 nem do Outlook. Use o [modelo de API Comum](office-javascript-api-object-model.md) para trabalhar com essas versões do Office. Para notas completas sobre disponibilidade de plataforma, confira [Disponibilidade de plataforma e de Aplicativo cliente do Office para Suplementos do Office](/javascript/api/requirement-sets).

> [!TIP]
> Os exemplos nesta página usam as APIs JavaScript do Excel, mas os conceitos também se aplicam às APIs do OneNote, PowerPoint, Visio e Word JavaScript.

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Caráter assíncrono das APIs baseadas em promessas

Os Suplementos do Office são sites que aparecem dentro de um contêiner de navegador em aplicativos do Office, como o Excel. Esse contêiner é incorporado no aplicativo do Office em plataformas baseadas na área de trabalho, como o Office no Windows, e é executado em um iFrame HTML no Office na Web. Devido a considerações de desempenho, as APIs do Office.js não podem interagir de forma sincronizada com os aplicativos do Office em todas as plataformas. Desse modo, a chamada à API `sync()` no Office.js retorna uma [Promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) que é resolvida quando o aplicativo do Excel conclui as ações solicitadas de leitura ou de gravação. Além disso, você pode enfileirar várias ações, como configurar propriedades ou invocar métodos, e executá-las como um lote de comandos com uma única chamada a `sync()`, em vez de enviar uma solicitação separada para cada ação. As seções a seguir descrevem como fazer isso usando as APIs `run()` e `sync()`.

## <a name="run-function"></a>Função *.run

`Excel.run`, `OneNote.run`, `PowerPoint.run`e `Word.run` execute uma função que especifica as ações a serem executadas no Excel, Word e OneNote. `*.run` cria automaticamente um contexto de solicitação que pode ser usado para interagir com objetos do Excel. Ao concluir `*.run`, uma promessa será resolvida e todos os objetos que foram alocados em tempo de execução serão lançados automaticamente.

O exemplo a seguir mostra como usar `Excel.run`. O mesmo padrão também é usado com o OneNote, o PowerPoint e o Word.

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

O aplicativo do Office e seu suplemento são executados em processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos exigem um objeto `RequestContext` para conectar o suplemento a objetos no Office, como planilhas, intervalos, gráficos e tabelas. Esse objeto `RequestContext` é fornecido como um argumento ao chamar `*.run`.

## <a name="proxy-objects"></a>Objetos proxy

Os objetos JavaScript do Office, que você declara e usa com as APIs baseadas em promessa, são objetos proxy. Todos os métodos invocados, ou as propriedades definidas ou carregadas em objetos proxy são simplesmente adicionados a uma fila de comandos pendentes. Ao chamar o método `sync()` no contexto de solicitação (por exemplo, `context.sync()`), os comandos enfileirados são expedidos para o aplicativo do Office e executados. Essas APIs são fundamentalmente centradas em lotes. Enfileire quantas alterações desejar no contexto de solicitação e, em seguida, chame o método `sync()` para executar o lote de comandos enfileirados.

Por exemplo, o trecho de código a seguir declara o objeto JavaScript [Excel.Range](/javascript/api/excel/excel.range) local, `selectedRange`, para fazer referência a um intervalo selecionado na pasta de trabalho do Excel e, em seguida, define algumas propriedades nesse objeto. O objeto `selectedRange` é um objeto proxy, de modo que as propriedades definidas e o método invocado nesse objeto não serão refletidos no documento do Excel até que seu suplemento chame `context.sync()`.

```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>Minimizar o número de objetos proxy criados

Evite criar repetidamente o mesmo objeto proxy. Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
const range = worksheet.getRange("A1");
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

Chamar o método `sync()` no contexto de solicitação sincroniza o estado entre objetos proxy e objetos no documento do Office. O método `sync()` executa todos os comandos que são enfileirados no contexto de solicitação e recupera valores para qualquer propriedade que deva ser carregada nos objetos proxy. O método `sync()` é executado de modo assíncrono e retorna uma [Promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o método `sync()` é concluído.

O exemplo a seguir mostra uma função de lote que define um objeto proxy JavaScript local (`selectedRange`), carrega uma propriedade desse objeto e, em seguida, usa o padrão de promessas do JavaScript para chamar `context.sync()`, a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.

```js
await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    await context.sync();
    console.log('The selected range is: ' + selectedRange.address);
});
```

No exemplo anterior, `selectedRange` é definido e sua propriedade `address` é carregada quando `context.sync()` é chamado.

Como `sync()` é uma operação assíncrona, você sempre deve retornar o objeto `Promise` para garantir que a operação de `sync()` seja concluída antes que o script continue a ser executado. Se você estiver usando TypeScript ou ES6+ JavaScript, poderá `await` a chamada `context.sync()` em vez de retornar a promessa.

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>Dica de desempenho: minimizar o número de chamadas síncronas

Na API do JavaScript do Excel, `sync()` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online na Web. Para otimizar o desempenho, minimize o número de chamadas para `sync()`, enfileirando o maior número possível de alterações antes de chamá-lo. Para mais informações sobre como otimizar o desempenho com `sync()`, confira [Evitar o uso do método contexto.sync em loops](../concepts/correlated-objects-pattern.md).

### <a name="load"></a>load()

Antes de poder ler as propriedades de um objeto proxy, será necessário carregar explicitamente as propriedades para preencher o objeto proxy com dados do documento do Office e, em seguida, chamar `context.sync()`. Por exemplo, se você criar um objeto proxy para referenciar um intervalo selecionado e, em seguida, quiser ler a propriedade `address` do intervalo selecionado, carregue a propriedade `address` antes de poder lê-la. Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o método `load()` no objeto e especifique as propriedades a serem carregadas. O exemplo a seguir mostra a propriedade `Range.address` sendo carregada para `myRange`.

```js
await Excel.run(async (context) => {
    const sheetName = 'Sheet1';
    const rangeAddress = 'A1:B2';
    const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');
    await context.sync();
      
    console.log (myRange.address);   // ok
    //console.log (myRange.values);  // not ok as it was not loaded

    console.log('done');
});
```

> [!NOTE]
> Se estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, não é necessário chamar o método `load()`. O método `load()` só é necessário quando você deseja ler propriedades em um objeto proxy.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

#### <a name="scalar-and-navigation-properties"></a>Propriedades escalares e de navegação

Há duas categorias de propriedades: **escalar** e de **navegação**. As propriedades escalares são tipos atribuíveis, como cadeias de caracteres, inteiros e estruturas JSON. As propriedades de navegação são objetos somente leitura e coleções de objetos que têm seus campos atribuídos, em vez de atribuir diretamente a propriedade. Por exemplo, os membros `name` e `position` no objeto [Excel.Worksheet](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são propriedades de navegação.

O suplemento pode usar propriedades de navegação como um caminho para carregar propriedades escalares específicas. O código a seguir enfileira um comando `load` para o nome da fonte usada por um objeto `Excel.Range`, sem carregar nenhuma outra informação.

```js
someRange.load("format/font/name")
```

Também é possível definir propriedades escalares de uma propriedade de navegação percorrendo o caminho. Por exemplo, é possível definir o tamanho da fonte de um `Excel.Range` usando `someRange.format.font.size = 10;`. Não é necessário carregar a propriedade antes de configurá-la.

Esteja ciente de que algumas das propriedades em um objeto podem ter o mesmo nome que outro objeto. Por exemplo, `format` é uma propriedade no objeto `Excel.Range`, mas `format` também é um objeto. Portanto, se você fizer uma chamada como `range.load("format")`, isso equivale a `range.format.load()` (uma instrução vazia e `load()` indevida). Para evitar isso, o código deve carregar apenas "nós folha" na árvore de objetos.

#### <a name="calling-load-without-parameters-not-recommended"></a>Chamando `load` sem parâmetros (não recomendado)

Se você chamar o método `load()` em um objeto (ou coleção) sem especificar nenhum parâmetro, todas as propriedades escalares do objeto ou dos objetos da coleção serão carregadas. Carregar dados não necessários desacelerá o seu suplemento. Sempre especifique explicitamente quais propriedades devem ser carregadas.

> [!IMPORTANT]
> A quantidade de dados retornados por uma declaração `load` sem parâmetros pode exceder os limites de tamanho do serviço. Para reduzir os riscos a suplementos mais antigos, algumas propriedades não são retornadas por `load` sem a solicitação explícita. As propriedades a seguir são excluídas dessas operações de carga.
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

Os métodos nas APIs baseadas em promessas que retornam tipos primitivos têm um padrão semelhante ao paradigma `load`/`sync`. Por exemplo, `Excel.TableCollection.getCount` obtém o número de tabelas da coleção. `getCount` retorna um `ClientResult<number>`, o que significa que a propriedade `value` em [`ClientResult`](/javascript/api/office/officeextension.clientresult) retornado é um número. Seu script não pode acessar esse valor até que `context.sync()` seja chamado.

O script a seguir obtém o número total de tabelas na pasta de trabalho do Excel e registra esse número no console.

```js
const tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
await context.sync();

// Trying to log the value before calling sync would throw an error.
console.log (tableCount.value);
```

### <a name="set"></a>set()

A definição de propriedades em um objeto com propriedades de navegação aninhadas pode ser uma tarefa complicada. Como uma alternativa para definir propriedades individuais usando caminhos de navegação, conforme descrito acima, use o método `object.set()` disponível em todos os objetos nas APIs JavaScript baseadas em promessas. Com esse método, é possível definir várias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que são estruturadas, como as propriedades do objeto no qual o método é chamado.

The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E2");
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

    await context.sync();
});
```

### <a name="some-properties-cannot-be-set-directly"></a>Algumas propriedades não podem ser definidas diretamente.

Algumas propriedades não podem ser definidas, apesar de serem graváveis. Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto. Isso porque essa propriedade pai depende das subpropriedades com relações lógicas específicas. Essas propriedades pai devem ser definidas usando notação literal de objeto para definir o objeto inteiro, em vez de definir subpropriedades individuais do objeto. Um exemplo disso é encontrado na página [PageLayout](/javascript/api/excel/excel.pagelayout). A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui.

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

No exemplo anterior, ***não*** seria possível atribuir um valor a `zoom` diretamente: `sheet.pageLayout.zoom.scale = 200;`. Essa instrução lança um erro porque `zoom` não foi carregado. Mesmo que `zoom` fosse carregado, o conjunto de escalas não seria efetivado. Todas as operações de contexto ocorrem em `zoom`, atualizando o objeto proxy no suplemento e sobrescrevendo os valores definidos localmente.

Esse comportamento difere das [propriedades navegacionais](application-specific-api-model.md#scalar-and-navigation-properties) como [Range.format](/javascript/api/excel/excel.range#excel-excel-range-format-member). As propriedades de `format` podem ser definidas usando a navegação de objeto, conforme mostrado aqui.

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente, verificando seu modificador somente leitura. Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente. Propriedades graváveis, como `PageLayout.zoom`, devem ser definidas com um objeto nesse nível. Em resumo:

- Propriedade somente leitura: as subpropriedades podem ser definidas por meio da navegação.
- Propriedade gravável: As subpropriedades não podem ser definidas por meio da navegação (devem ser definidas como parte da atribuição do objeto pai inicial).

## <a name="42ornullobject-methods-and-properties"></a>Métodos e propriedades &#42;OrNullObject

Alguns métodos e propriedades do acessador lançam uma exceção quando o objeto desejado não existe. Por exemplo, ao tentar obter uma planilha do Excel especificando um nome de planilha que não esteja na pasta de trabalho, o método `getItem()` lança uma exceção `ItemNotFound`. As bibliotecas específicas por aplicativo fornecem uma maneira do código testar a existência de entidades de documentos sem exigir código de tratamento de exceções. Isso é realizado usando as variações `*OrNullObject` de métodos e propriedades. Essas variações retornam um objeto cuja propriedade `isNullObject` está definida como `true`, se o item especificado não existir, em vez de lançar uma exceção.

Por exemplo, você pode chamar o método `getItemOrNullObject()` em uma coleção, como **Planilhas**, para recuperar um item da coleção. O método `getItemOrNullObject()` retornará o item especificado se ele existir; caso contrário, ele retornará um objeto cuja propriedade `isNullObject` estiver definida como `true`. O código pode então avaliar essa propriedade para determinar se o objeto existe.

> [!NOTE]
> As variações `*OrNullObject` nunca retornam o valor de JavaScript `null`. Elas retornam objetos proxy comuns do Office. Se a entidade que o objeto representa não existir, então a propriedade `isNullObject` do objeto será definida como `true`. Não teste o objeto retornado para nulidade ou falsidade. Ele nunca é `null`, `false`ou `undefined`.

O exemplo de código a seguir tenta recuperar uma planilha do Excel chamada "Dados", usando o método `getItemOrNullObject()`. Se uma planilha com esse nome não existir, uma nova planilha será criada. Observe que o código não carrega a propriedade `isNullObject`. O Office carrega automaticamente essa propriedade quando `context.sync` for chamada, então não é necessário carregá-la explicitamente com algo como `dataSheet.load('isNullObject')`.

```js
await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
    
    await context.sync();
    
    if (dataSheet.isNullObject) {
        dataSheet = context.workbook.worksheets.add("Data");
    }
    
    // Set `dataSheet` to be the second worksheet in the workbook.
    dataSheet.position = 1;
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto comum de API JavaScript](office-javascript-api-object-model.md)
- [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
