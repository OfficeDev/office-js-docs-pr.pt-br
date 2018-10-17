---
title: Conceitos fundamentais de programação com a API JavaScript do Excel
description: Usar a API JavaScript do Excel para criar suplementos para o Excel.
ms.date: 10/03/2018
ms.openlocfilehash: f93ec7b5e34f90f2d61f29d861b7e0c19f66f6e3
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505983"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Conceitos fundamentais de programação com a API JavaScript do Excel
 
Este artigo descreve como usar a [API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) para criar suplementos para o Excel 2016 ou posterior. Ele apresenta os principais conceitos fundamentais para o uso da API e fornece orientação para a realização de tarefas específicas, como fazer a leitura ou gravação em um intervalo grande, atualizar todas as células de um intervalo e muito mais.

## <a name="asynchronous-nature-of-excel-apis"></a>Natureza assíncrona das APIs do Excel

Os suplementos do Excel baseados na Web são executados dentro de um contêiner de navegador incorporado ao aplicativo do Office em plataformas baseadas em área de trabalho, como o Office para o Windows, e é executado dentro de um iFrame HTML no Office Online. Não é viável habilitar a API Office.js para interagir de maneira síncrona com o host do Excel em todas as plataformas com suporte devido a considerações de desempenho. Portanto, a chamada **sync()** de API no Office.js retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) que é resolvida quando o aplicativo do Excel conclui as ações de leitura ou gravação solicitadas. Além disso, você pode adicionar várias ações, como definir propriedades ou métodos, a uma fila e executá-las como um lote de comandos com uma única chamada **sync()**, ao invés de enviar uma solicitação separada para cada ação. As seções a seguir descrevem como realizar essa tarefa usando as APIs **Excel.run()** e **sync()**.
 
## <a name="excelrun"></a>Excel.run
 
**Excel.Run** executa uma função onde você pode especificar as ações a serem executadas em relação ao modelo de objeto do Excel. **Excel.Run** cria automaticamente um contexto de solicitação que você pode usar para interagir com objetos do Excel. Quando **Excel.run** é concluída, uma promessa é resolvida e todos os objetos alocados durante o tempo de execução são automaticamente liberados.
 
O exemplo a seguir mostra como usar **Excel.run**. A instrução catch captura e registra os erros que ocorrem em **Excel.run**.
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="request-context"></a>Contexto de solicitação
 
O Excel e o seu suplemento são executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos do Excel exigem um objeto **RequestContext** para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gráficos e tabelas.
 
## <a name="proxy-objects"></a>Objetos proxy
 
Os objetos do Excel JavaScript que você declara e usa em um suplemento são objetos proxy. Qualquer método invocado ou propriedade definida ou carregada por você nos objetos proxy simplesmente são adicionadas a uma fila de comandos pendentes. Quando você chama o método **sync()** no contexto da solicitação (por exemplo, `context.sync()`), os comandos na fila são enviados para o Excel e executados. A API JavaScript do Excel é fundamentalmente centrada em lotes. Você colocar quantas alterações desejar na fila no contexto da solicitação e, então, chamar o método **sync()** para executar o lote de comandos.
 
Por exemplo, o trecho de código a seguir declara o objeto JavaScript **selectedRange** local para fazer referência a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto. O objeto **selectedRange** é um objeto proxy, portanto, as propriedades definidas e o método invocado em nele não refletem no documento do Excel até que seu suplemento chame **context.sync()**.
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>sync()
 
Chamar o método **sync()** no contexto da solicitação sincroniza o estado entre os objetos proxy e os objetos no documento do Excel. O método **sync()** executa os comandos na fila no contexto da solicitação e recupera os valores para todas as propriedades que devem ser carregadas nos objetos proxy. O método **sync()** é executado de forma assíncrona e retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o método **sync()** é concluído.
 
O exemplo a seguir mostra uma função de lote que define um objeto proxy JavaScript local (**selectedRange**), carrega uma propriedade desse objeto e, em seguida, usa o padrão JavaScript Promises para chamar **context.sync()** a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
No exemplo anterior, **selectedRange** está definido e sua propriedade **address** é carregada quando **context.sync()** é chamado.
 
Como **sync()** é uma operação assíncrona que retorna uma promessa, você sempre deve **retornar** a promessa (em JavaScript). Isso garante que a operação **sync()** seja concluída antes do script continuar a ser executado. Para obter mais informações sobre como otimizar o desempenho com **sync()**, consulte [Otimização do desempenho da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).
 
### <a name="load"></a>load()
 
Antes de poder ler as propriedades de um objeto proxy, você deve carregar explicitamente as propriedades para preencher o objeto com dados de um documento do Excel e, em seguida, chamar **context.sync()**. Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e quiser ler a propriedade **address** desse intervalo, você primeiro precisará carregar a propriedade **address**. Para solicitar que uma propriedade de um objeto proxy seja carregada, chame o método **load()** no objeto e especifique as propriedades que devem ser carregadas. 

> [!NOTE]
> Se você estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o método **load()**. O método **load()** só é necessário quando você deseja ler as propriedades em um objeto proxy.
 
Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto da solicitação, sendo executadas na próxima vez que você chamar o método **sync()** . É possível colocar quantas chamadas de **load()** forem necessárias na fila no contexto da solicitação.
 
No exemplo a seguir, somente propriedades específicas do intervalo são carregadas.
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
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
 
No exemplo anterior, como `format/font` não é especificado na chamada a **myRange.load()**, a propriedade `format.font.color` não pode ser lida.

Para otimizar o desempenho, você deve especificar explicitamente as propriedades e relações a serem carregadas ao usar o método **load()** em um objeto, conforme abordado em [Otimizações de desempenho da API JavaScript do Excel](performance.md). Para obter mais informações sobre o método **Load** , consulte [Conceitos de programação avançados com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md).

## <a name="null-or-blank-property-values"></a>Valores de propriedade null ou blank
 
### <a name="null-input-in-2-d-array"></a>entrada nula em uma matriz 2D
 
No Excel, um intervalo é representado por uma matriz 2-D, onde a primeira dimensão é formada por linhas e a segunda por colunas. Para definir valores, formatos de número ou fórmulas para células específicas dentro de um intervalo, especifique os valores, formatos de número ou fórmulas para essas células na matriz 2D e especifique `null` para todas as outras células.
 
Por exemplo, para atualizar o formato de número de apenas uma célula dentro de um intervalo de atualização e manter o formato existente para todas as outras células, especifique o novo formato para a célula que deseja atualizar e especifique `null` para todas as outras células. Os trechos de código a seguir definem um novo formato de número para a quarta célula do intervalo e mantém o formato inalterado para as três primeiras células no intervalo.
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>entrada nula para uma propriedade
 
`null` não é uma entrada válida para uma única propriedade. Por exemplo, o snippet de código a seguir não é válido, pois a propriedade **values** do intervalo não pode ser definida como `null`.
 
```js
range.values = null;
```
 
Da mesma forma, o snippet de código a seguir não é válido, pois `null` não é um valor válido para a propriedade **color**.
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>Valores nulos para propriedades na resposta
 
Propriedades de formatação, como `size` e `color` contêm valores `null` na resposta quando há valores diferentes no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar a sua propriedade `format.font.color`:
 
* Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especifica essa cor.
* Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.
 
### <a name="blank-input-for-a-property"></a>Entrada em branco para uma propriedade
 
Quando você especifica um valor em branco para uma propriedade (isto é, duas aspas sem espaço `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:
 
* Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.
 
* Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.
 
* Se você especificar um valor em branco para as propriedades `formula` e `formulaLocale`, os valores de fórmula serão apagados.
 
### <a name="blank-property-values-in-the-response"></a>Valores de propriedade em branco na resposta
 
Para operações de leitura, um valor de uma propriedade em branco na resposta (ou seja, duas aspas sem espaço `''`) indica que a célula não contém nenhum dado ou valor. No primeiro exemplo a seguir, a primeira e a última célula no intervalo não contêm nenhum dado. No segundo exemplo, as duas primeiras células no intervalo não contém uma fórmula.
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>Ler ou gravar em um intervalo não associado
 
### <a name="read-an-unbounded-range"></a>Ler um intervalo não associado
 
Um endereço de intervalo não associado é um endereço de intervalo que especifica colunas ou linhas inteiras. Por exemplo:
 
* Endereços de intervalo compostos por colunas inteiras:<ul><li>`C:C`</li><li>`A:F`</li></ul>
* Endereços de intervalo compostos por linhas inteiras:<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
Quando a API faz uma solicitação para recuperar um intervalo não associado (por exemplo, `getRange('C:C')`), a resposta contém valores `null` para propriedades de nível de célula, tais como `values`, `text`, `numberFormat`, e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, contêm valores válidos para o intervalo não associado.
 
### <a name="write-to-an-unbounded-range"></a>Gravar em um intervalo não associado
 
Você não pode definir propriedades em nível de célula, como `values`, `numberFormat`, e `formula`, em um intervalo não associado, pois a solicitação de entrada é muito grande. Por exemplo, o snippet de código a seguir não é válido pois tenta especificar `values` para um intervalo não associado. A API retornará um erro se você tentar definir propriedades de nível de célula para um intervalo não associado.
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>Ler ou gravar em um intervalo grande
 
Se um intervalo contiver um grande número de células, valores, formatos de número e/ou fórmulas, pode ser que não seja possível executar operações de API nesse intervalo. A API sempre fará a melhor tentativa para executar a operação solicitada em um intervalo (ou seja, recuperar ou gravar os dados especificados), mas a tentativa de executar operações de leitura ou gravação em um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos. Para evitar esses erros, recomendamos que você execute operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação em um intervalo grande.
 
## <a name="update-all-cells-in-a-range"></a>Atualizar todas as células em um intervalo
 
Para aplicar a mesma atualização a todas as células em um intervalo, (por exemplo, popular todas as células com o mesmo valor, definir o mesmo formato de número ou popular todas as células com a mesma fórmula), defina a propriedade correspondente no objeto **range** com o valor (único) desejado.
 
O exemplo a seguir obtém um intervalo que contém 20 células e, em seguida, define o formato de número e popula todas as células do intervalo com o valor **11/3/2015**.
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a>Mensagens de erro
 
Quando ocorre um erro de API, a API retorna um objeto de **erro** que contém um código e uma mensagem. A tabela a seguir define uma lista de erros que a API pode retornar.
 
|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |O argumento é inválido, ausente ou tem um formato incorreto.|
|InvalidRequest  |Não é possível processar a solicitação.|
|InvalidReference|Essa referência não é válida para a operação atual.|
|InvalidBinding  |Essa associação de objetos não é mais válida devido a atualizações anteriores.|
|InvalidSelection|A seleção atual é inválida para esta operação.|
|Unauthenticated |Informações de autenticação necessárias estão ausentes ou inválidas.|
|AccessDenied |Você não pode realizar a operação solicitada.|
|ItemNotFound |O recurso solicitado não existe.|
|ActivityLimitReached|O limite de atividades foi alcançado.|
|GeneralException|Ocorreu um erro interno ao processar a solicitação.|
|NotImplemented  |O recurso solicitado não foi implementado.|
|ServiceNotAvailable|O serviço não está disponível.|
|Conflict              |A solicitação não pôde ser processada devido a um conflito.|
|ItemAlreadyExists|O recurso que está sendo criado já existe.|
|UnsupportedOperation|Não há suporte para a operação.|
|RequestAborted|A solicitação foi anulada durante o tempo de execução.|
|ApiNotAvailable|A API solicitada não está disponível.|
|InsertDeleteConflict|A operação de exclusão ou inserção resultou em um conflito.|
|InvalidOperation|A operação é inválida no objeto.|
 
## <a name="see-also"></a>Veja também
 
* [Introdução aos suplementos do Excel](excel-add-ins-get-started-overview.md)
* [Exemplos de códigos de suplementos do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [Conceitos de programação avançados com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md)
* [Otimização de desempenho da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [Referência da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
