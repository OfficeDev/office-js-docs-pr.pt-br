---
title: Principais conceitos da API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1582268a3bdac2b7fe63c4b0a48cf1a19f85bd31
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-core-concepts"></a>Principais conceitos da API JavaScript do Excel
 
Este artigo descreve como usar a [API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) para desenvolver suplementos para o Excel 2016. Ele apresenta os conceitos b?sicos que s?o fundamentais para usar a API e fornece orienta??es para executar tarefas espec?ficas, como leitura ou grava??o em um intervalo grande, atualiza??o de todas as c?lulas do intervalo e muito mais.

## <a name="asynchronous-nature-of-excel-apis"></a>Natureza ass?ncrona das APIs do Excel

Os suplementos do Excel baseados na Web s?o executados dentro de um cont?iner de navegador que ? inserido no aplicativo do Office em plataformas baseadas em desktop, como Office para Windows, e executado dentro de um iFrame HTML no Office Online. N?o ? poss?vel habilitar a API Office.js para interagir de modo s?ncrono com o host do Excel em todas as plataformas suportadas devido ?s considera??es de desempenho. Desse modo, a chamada ? API **sync()** na Office.js retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) que ? resolvida quando o aplicativo Excel conclui as a??es solicitadas de leitura ou grava??o. Al?m disso, voc? pode enfileirar v?rias a??es, como configurar propriedades ou invocar m?todos, e execut?-las como um lote de comandos com uma ?nica chamada a **sync()**, em vez de enviar uma solicita??o separada para cada a??o. As se??es a seguir descrevem como fazer isso usando as APIs **Excel.run()** e **sync()**.
 
## <a name="excelrun"></a>Excel.run
 
A **Excel.run** executa uma fun??o em que voc? especifica as a??es a serem executadas no modelo de objeto do Excel. A **Excel.run** cria automaticamente um contexto de solicita??o que pode ser usado para sua intera??o com os objetos do Excel. Quando a **Excel.run** ? conclu?da, uma promessa ? resolvida e todos os objetos que foram alocados em tempo de execu??o s?o lan?ados automaticamente.
 
O exemplo a seguir mostra como usar a **Excel.run**. A instru??o catch captura e grava em log os erros que ocorrem na **Excel.run**.
 
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

## <a name="request-context"></a>Contexto de solicita??o
 
O Excel e seu suplemento s?o executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execu??o, os suplementos do Excel exigem um objeto **RequestContext** para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gr?ficos e tabelas.
 
## <a name="proxy-objects"></a>Objetos proxy
 
Os objetos JavaScript do Excel que voc? declara e usa em um suplemento s?o objetos proxy. Todos os m?todos invocados, ou as propriedades definidas ou carregadas em objetos proxy s?o simplesmente adicionados a uma fila de comandos pendentes. Quando voc? chama o m?todo **sync()** no contexto de solicita??o (por exemplo, `context.sync()`), os comandos enfileirados s?o expedidos para o Excel e executados. A API JavaScript do Excel ? basicamente centrada em lote. Voc? pode enfileirar quantas altera??es desejar no contexto de solicita??o e depois chamar o m?todo **sync()** para executar o lote de comandos enfileirados.
 
Por exemplo, o trecho de c?digo a seguir declara o objeto JavaScript local **selectedRange** para fazer refer?ncia a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto. O objeto **selectedRange** ? um objeto proxy, de modo que as propriedades que s?o definidas e o m?todo que ? invocado nesse objeto n?o ser?o refletidos no documento do Excel at? que o suplemento chame **context.sync()**.
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>sync()
 
Chamar o m?todo **sync()** no contexto de solicita??o sincroniza o estado entre objetos proxy e objetos no documento do Excel. O m?todo **sync()** executa todos os comandos que s?o enfileirados no contexto de solicita??o e recupera valores para qualquer propriedade que deva ser carregada nos objetos proxy. O m?todo **sync()** ? executado de modo ass?ncrono e retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), que ? resolvida quando o m?todo **sync()** ? conclu?do.
 
O exemplo a seguir mostra uma fun??o de lote que define um objeto proxy JavaScript local (**selectedRange**), carrega uma propriedade desse objeto e, em seguida, usa o padr?o Promessas do JavaScript para chamar **context.sync()** a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.
 
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
 
No exemplo anterior, **selectedRange** est? definido e sua propriedade **address** ? carregada quando **context.sync()** ? chamado.
 
Como **sync()** ? uma opera??o ass?ncrona que retorna uma promessa, voc? sempre deve **retornar** a promessa (no JavaScript). Isso garante que a opera??o **sync()** seja conclu?da antes que o script continue sendo executado. Para obter mais informa??es sobre como otimizar o desempenho com **sync()**, confira [Otimiza??o de desempenho da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/performance.md).
 
### <a name="load"></a>load()
 
Para que voc? possa ler as propriedades de um objeto proxy, ? preciso carregar explicitamente as propriedades para popular o objeto proxy com dados do documento do Excel e chamar **context.sync()**. Por exemplo, se voc? criar um objeto proxy para fazer refer?ncia a um intervalo selecionado e, em seguida, quiser ler a propriedade **address** do intervalo selecionado, ser? preciso carregar a propriedade **address** para que seja poss?vel l?-la. Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o m?todo **load()** no objeto e especifique as propriedades a serem carregadas. 

> [!NOTE]
> Se estiver apenas chamando m?todos ou definindo propriedades em um objeto proxy, voc? n?o precisa chamar o m?todo **load()**. O m?todo **load()** s? ? necess?rio quando voc? deseja ler propriedades em um objeto proxy.
 
Assim como as solicita??es para definir propriedades ou invocar m?todos em objetos proxy, as solicita??es para carregar propriedades em objetos proxy s?o adicionadas ? fila de comandos pendentes no contexto de solicita??o, sendo executadas na pr?xima vez que voc? chamar o m?todo **sync()**. ? poss?vel enfileirar quantas chamadas de **load()** forem necess?rias no contexto de solicita??o.
 
No exemplo a seguir, somente propriedades espec?ficas do intervalo s?o carregadas.
 
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
 
No exemplo anterior, como `format/font` n?o ? especificado na chamada a **myRange.load()**, a propriedade `format.font.color` n?o pode ser lida.

Para otimizar o desempenho, voc? deve especificar explicitamente as propriedades e os relacionamentos a serem carregados ao usar o m?todo **load()** em um objeto, conforme [Otimiza??es de desempenho da API JavaScript do Excel](performance.md). Para saber mais sobre o m?todo **load()**, confira os [conceitos avan?ados da API JavaScript do Excel](excel-add-ins-advanced-concepts.md).

## <a name="null-or-blank-property-values"></a>Valores de propriedade nula ou em branco
 
### <a name="null-input-in-2-d-array"></a>entrada nula em uma matriz 2D
 
No Excel, um intervalo ? representado por uma matriz 2D, onde a primeira dimens?o ? linhas e a segunda dimens?o ? colunas. Para definir valores, o formato do n?mero ou a f?rmula apenas para c?lulas espec?ficas em um intervalo, especifique os valores, o formato do n?mero ou a f?rmula para essas c?lulas na matriz 2D, bem como `null` para todas as outras c?lulas na matriz 2D.
 
Por exemplo, para atualizar o formato do n?mero apenas para uma c?lula em um intervalo e manter o formato de n?mero existente para todas as outras c?lulas no intervalo, especifique o novo formato de n?mero para a c?lula a ser atualizada e `null` para todas as outras c?lulas. O trecho de c?digo a seguir define um novo formato de n?mero para a quarta c?lula no intervalo e n?o altera o formato de n?mero para as primeiras tr?s c?lulas no intervalo.
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>entrada nula para uma propriedade
 
`null` n?o ? uma entrada v?lida para uma propriedade ?nica. Por exemplo, o trecho de c?digo a seguir n?o ? v?lido, pois a propriedade **values** do intervalo n?o pode ser definida como `null`.
 
```js
range.values = null;
```
 
Da mesma forma, o trecho de c?digo a seguir n?o ? v?lido, pois `null` n?o ? um valor v?lido para a propriedade **color**.
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>Valores da propriedade nula na resposta
 
A formata??o de propriedades como `size` e `color` conter? valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se voc? recuperar um intervalo e carregar sua propriedade `format.font.color`:
 
* Se todas as c?lulas no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificar? essa cor.
* Se houver v?rias cores de fonte dentro do intervalo, `range.format.font.color` ser? `null`.
 
### <a name="blank-input-for-a-property"></a>Entrada em branco para uma propriedade
 
Quando voc? especificar um valor em branco para uma propriedade (isto ?, duas aspas sem espa?o entre elas `''`), ele ser? interpretado como uma instru??o para limpar ou redefinir a propriedade. Por exemplo:
 
* Se voc? especificar um valor em branco para a propriedade `values` de um intervalo, o conte?do do intervalo ser? apagado.
 
* Se voc? especificar um valor em branco para a propriedade `numberFormat`, o formato de n?mero ser? redefinido para `General`.
 
* Se voc? especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de f?rmula ser?o apagados.
 
### <a name="blank-property-values-in-the-response"></a>Valores da propriedade em branco na resposta
 
Para opera??es de leitura, um valor de propriedade em branco na resposta (isto ?, duas aspas sem espa?o entre elas `''`) indica que a c?lula n?o cont?m dados nem valor. No primeiro exemplo abaixo, a primeira e a ?ltima c?lula no intervalo n?o cont?m dados. No segundo exemplo, as primeiras duas c?lulas no intervalo n?o cont?m uma f?rmula.
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>Ler ou gravar em um intervalo n?o limitado
 
### <a name="read-an-unbounded-range"></a>Ler um intervalo n?o limitado
 
Um endere?o de intervalo n?o limitado ? um endere?o de intervalo que especifica colunas ou linhas inteiras. Por exemplo:
 
* Endere?os de intervalo composto por colunas inteiras:<ul><li>`C:C`</li><li>`A:F`</li></ul>
* Endere?os de intervalo composto por linhas inteiras:<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
Quando uma API faz uma solicita??o para recuperar um intervalo n?o limitado (por exemplo, `getRange('C:C')`), a resposta conter? valores `null` para as propriedades no n?vel de c?lula, como `values`, `text`, `numberFormat` e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, conter?o valores v?lidos para o intervalo n?o limitado.
 
### <a name="write-to-an-unbounded-range"></a>Gravar em um intervalo n?o limitado
 
N?o ? poss?vel definir propriedades no n?vel de c?lula, como `values`, `numberFormat` e `formula`, no intervalo n?o limitado, pois a solicita??o de entrada ? muito grande. Por exemplo, o trecho de c?digo a seguir n?o ? v?lida porque ele tenta especificar `values` para um intervalo n?o limitado. A API retornar? um erro se voc? tentar definir as propriedades no n?vel de c?lula para um intervalo n?o limitado.
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>Ler ou gravar em um intervalo grande
 
Se um intervalo contiver um grande n?mero de c?lulas, valores, formatos de n?mero e/ou f?rmulas, talvez n?o seja poss?vel executar opera??es de API nesse intervalo. A API sempre far? a melhor tentativa de executar a opera??o solicitada em um intervalo (isto ?, para recuperar ou gravar os dados especificados), mas tentar executar opera??es de leitura ou grava??o para um intervalo grande pode resultar em um erro de API devido ? utiliza??o excessiva de recursos. Para evitar tais erros, ? recomend?vel executar opera??es de leitura ou grava??o separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma ?nica opera??o de leitura ou grava??o em um intervalo grande.
 
## <a name="update-all-cells-in-a-range"></a>Atualizar todas as c?lulas em um intervalo
 
Para aplicar a mesma atualiza??o a todas as c?lulas em um intervalo, (por exemplo, para popular todas as c?lulas com o mesmo valor, definir o mesmo formato de n?mero ou popular todas as c?lulas com a mesma f?rmula), defina a propriedade correspondente no objeto **range** para o valor (?nico) desejado.
 
O exemplo a seguir obt?m um intervalo que cont?m 20 c?lulas e, em seguida, define o formato de n?mero e popula todas as c?lulas do intervalo com o valor **11/3/2015**.
 
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
 
Quando ocorrer um erro de API, a API retornar? um objeto **error** que cont?m um c?digo e uma mensagem. A tabela a seguir define uma lista de erros que a API pode retornar.
 
|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |O argumento ? inv?lido, est? ausente ou tem um formato incorreto.|
|InvalidRequest  |N?o ? poss?vel processar a solicita??o.|
|InvalidReference|Esta refer?ncia n?o ? v?lida para a opera??o atual.|
|InvalidBinding  |Esta associa??o de objetos n?o ? mais v?lida devido ?s atualiza??es anteriores.|
|InvalidSelection|A sele??o atual ? inv?lida para esta opera??o.|
|Unauthenticated |Informa??es de autentica??o necess?rias est?o ausentes ou inv?lidas.|
|AccessDenied |Voc? n?o pode realizar a opera??o solicitada.|
|ItemNotFound |O recurso solicitado n?o existe.|
|ActivityLimitReached|O limite de atividades foi alcan?ado.|
|GeneralException|Ocorreu um erro interno ao processar a solicita??o.|
|NotImplemented  |O recurso solicitado n?o foi implementado.|
|ServiceNotAvailable|O servi?o n?o est? dispon?vel.|
|Conflito              |A solicita??o n?o p?de ser processada devido a um conflito.|
|ItemAlreadyExists|O recurso que est? sendo criado j? existe.|
|UnsupportedOperation|N?o h? suporte para a opera??o que est? sendo tentada.|
|RequestAborted|A solicita??o foi anulada durante o tempo de execu??o.|
|ApiNotAvailable|A API solicitada n?o est? dispon?vel.|
|InsertDeleteConflict|A tentativa de opera??o de exclus?o ou inser??o resultou em um conflito.|
|InvalidOperation|A tentativa de opera??o ? inv?lida no objeto.|
 
## <a name="see-also"></a>Veja tamb?m
 
* [Introdu??o aos suplementos do Excel](excel-add-ins-get-started-overview.md)
* [Exemplos de c?digo de suplementos do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Otimiza??o de desempenho da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/performance.md)
* [Refer?ncia da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
