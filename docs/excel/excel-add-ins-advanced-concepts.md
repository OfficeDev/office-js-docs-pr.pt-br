---
title: Conceitos avan?ados da API JavaScript do Excel
description: ''
ms.date: 1/18/2018
ms.openlocfilehash: 89db69e124475c882448a2105837787ce2c84753
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-advanced-concepts"></a>Conceitos avan?ados da API JavaScript do Excel

Este artigo foi criado com base nas informa??es em [conceitos principais da API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avan?ados que s?o essenciais para criar suplementos complexos para o Excel 2016. 

## <a name="officejs-apis-for-excel"></a>APIs Office.js para Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:

* **API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) fornece objetos fortemente tipados que voc? pode usar para acessar planilhas, intervalos, tabelas, gr?ficos e muito mais. 

* **APIs comuns**: introduzidas com o Office 2013, as APIs comuns (tamb?m conhecidas como a [API Compartilhada](https://dev.office.com/reference/add-ins/javascript-api-for-office)) podem ser usadas para acessar recursos como interface de usu?rio, caixas de di?logo e configura??es de cliente, que s?o comuns entre v?rios tipos de aplicativos host, como Word, Excel ou PowerPoint.

Enquanto voc? provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, voc? tamb?m usar? objetos na API Compartilhada. Por exemplo:

- [Contexto](https://dev.office.com/reference/add-ins/shared/context): o objeto **Context** representa o ambiente de tempo de execu??o do suplemento e oferece acesso aos principais objetos da API. Ele consiste em detalhes da configura??o da pasta de trabalho, como `contentLanguage` e `officeTheme`, al?m de fornecer informa??es sobre o ambiente de tempo de execu??o do suplemento, como `host` e `platform`. Al?m disso, ele fornece o m?todo `requirements.isSetSupported()`, que voc? pode usar para verificar se o conjunto de requisitos especificado ? suportado pelo aplicativo Excel onde o suplemento est? sendo executado. 

- [Document](https://dev.office.com/reference/add-ins/shared/document): O objeto **Document** fornece o m?todo `getFileAsync()`, que voc? pode usar para baixar o arquivo Excel onde o suplemento est? em execu??o. 

## <a name="requirement-sets"></a>Conjuntos de requisitos

Os conjuntos de requisitos s?o grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verifica??o de tempo de execu??o ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office d? suporte ?s APIs necess?rias ao suplemento. Para identificar os conjuntos de requisitos espec?ficos que est?o dispon?veis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificando o suporte ao conjunto de requisitos no tempo de execu??o

O exemplo de c?digo a seguir mostra como determinar se o aplicativo host, onde o suplemento est? em execu??o, d? suporte ao conjunto de requisitos da API especificado.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definindo o suporte ao conjunto de requisitos no manifesto

Voc? pode usar o [elemento Requirements](https://dev.office.com/reference/add-ins/manifest/requirements) no manifesto do suplemento para especificar os conjuntos de requisitos m?nimos e/ou os m?todos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o host do Office n?o der suporte aos conjuntos de requisitos ou aos m?todos de API que s?o especificados no elemento **Requirements** do manifesto, o suplemento n?o ser? executado nesse host ou plataforma e n?o ser? exibido na lista de suplementos que s?o mostrados em **Meus Suplementos**. 

O exemplo de c?digo a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos host do Office que d?o suporte ao conjunto de requisitos ExcelApi, vers?o 1.3 ou superior.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel para Windows, Excel Online e Excel para iPad, ? recomend?vel verificar o suporte a requisitos no tempo de execu??o, em vez de definir o suporte ao conjunto de requisitos no manifesto.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Conjuntos de requisitos para a API comum Office.js

Para saber mais sobre conjuntos de requisitos de API comum, confira [Conjuntos de requisitos de API comum do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).

## <a name="loading-the-properties-of-an-object"></a>Carregando as propriedades de um objeto

Chamar o m?todo `load()` em um objeto JavaScript do Excel orienta a API a carregar o objeto na mem?ria do JavaScript quando o m?todo `sync()` ? executado. O m?todo `load()` aceita uma cadeia de caracteres que cont?m nomes de propriedades delimitados por v?rgulas a serem carregados ou um objeto que especifica propriedades a serem carregadas, op??es de pagina??o, etc. 

> [!NOTE]
> Se voc? chamar o m?todo `load()` em um objeto (ou uma cole??o) sem especificar qualquer par?metro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos na cole??o) ser?o carregadas. Para reduzir a quantidade de transfer?ncia de dados entre o aplicativo host e o suplemento do Excel, voc? deve evitar chamar o m?todo `load()` sem especificar explicitamente quais propriedades carregar.

### <a name="method-details"></a>Detalhes do m?todo

#### <a name="loadparam-object"></a>load(param: object)

Preenche o objeto proxy criado na camada JavaScript com os valores da propriedade e do objeto especificados pelos par?metros.

#### <a name="syntax"></a>Sintaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Par?metros

|**Par?metro**|**Tipo**|**Descri??o**|
|:------------|:-------|:----------|
|`param`|objeto|Opcional. Aceita nomes de par?metro e de rela??o como uma matriz ou cadeia de caracteres delimitada por v?rgulas. Um objeto tamb?m pode ser passado para definir as propriedades de navega??o e sele??o (conforme mostrado no exemplo abaixo).|

#### <a name="returns"></a>Retorna

nulo

#### <a name="example"></a>Exemplo

O exemplo de c?digo a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo. Observe que o objeto de origem deve ser carregado primeiro para que seus valores de propriedade possam ser acessados e gravados no intervalo de destino. Este exemplo pressup?e que h? dados nos dois intervalos (**B2:E2** e **B7:E7**) e que os dois intervalos s?o inicialmente formatados de modo diferente.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange); 
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a>Carregar propriedades de op??o

Como uma alternativa para passar uma cadeia de caracteres delimitada por v?rgulas ou uma matriz ao chamar o m?todo `load()`, voc? pode passar um objeto que cont?m as propriedades a seguir. 

|**Propriedade**|**Tipo**|**Descri??o**|
|:-----------|:-------|:----------|
|`select`|objeto|Inclui uma lista delimitada por v?rgula ou uma matriz de nomes de par?metro/rela??o. Opcional.|
|`expand`|objeto|Inclui uma lista delimitada por v?rgula ou uma matriz de nomes de rela??o. Opcional.|
|`top`|int| Especifica o n?mero m?ximo de itens da cole??o que podem ser inclu?dos no resultado. Opcional. Voc? s? pode usar essa op??o quando usar a op??o de nota??o de objeto.|
|`skip`|int|Determina o n?mero de itens da cole??o que devem ser ignorados e n?o inclu?dos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados ser? iniciado depois de ignorar o n?mero de itens especificado. Opcional. Voc? s? pode usar esta op??o ao usar a op??o de nota??o de objeto.|

O exemplo de c?digo a seguir carrega uma cole??o de planilhas selecionando a propriedade `name` e o `address` do intervalo usado para cada planilha na cole??o. Ele tamb?m especifica que apenas as cinco planilhas principais na cole??o devem ser carregadas. Voc? poderia processar o pr?ximo conjunto de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributo. 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propriedades escalares e de navega??o 

Na documenta??o de refer?ncia da API JavaScript do Excel, voc? pode notar que os membros do objeto s?o agrupados em duas categorias: **propriedades** e **rela??es**. Uma propriedade de um objeto ? um membro escalar como uma cadeia de caracteres, um n?mero inteiro ou um valor booliano, enquanto uma rela??o de um objeto (tamb?m conhecida como uma propriedade de navega??o) ? um membro que ? ou um objeto ou uma cole??o de objetos. Por exemplo, os membros `name` e `position` no objeto [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) s?o propriedades escalares, enquanto `protection` e `tables` s?o rela??es (propriedades de navega??o). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propriedades escalares e propriedades de navega??o com `object.load()`

Chamar o m?todo `object.load()` sem par?metros especificados carregar? todas as propriedades escalares do objeto; as propriedades de navega??o do objeto n?o ser?o carregadas. Al?m disso, as propriedades de navega??o n?o podem ser carregadas diretamente. Em vez disso, voc? deve usar o m?todo `load()` para fazer refer?ncia ?s propriedades escalares individuais na propriedade de navega??o desejada. Por exemplo, para carregar o nome da fonte de um intervalo, voc? deve especificar as propriedades de navega??o **format** e **font** como o caminho para a propriedade **name**:

```js
someRange.load("format/font/name")
```

> [!NOTE]
> com a API JavaScript do Excel, ? poss?vel definir propriedades escalares de uma propriedade de navega??o percorrendo o caminho. Por exemplo, ? poss?vel definir o tamanho da fonte de um intervalo usando `someRange.format.font.size = 10;`. N?o ? necess?rio carregar a propriedade antes de configur?-la. 

## <a name="setting-properties-of-an-object"></a>Definindo propriedades de um objeto

A defini??o de propriedades em um objeto com propriedades de navega??o aninhadas pode ser uma tarefa complicada. Como uma alternativa para definir propriedades individuais usando caminhos de navega??o, conforme descrito acima, voc? pode usar o m?todo `object.set()` que est? dispon?vel em todos os objetos na API JavaScript do Excel. Com esse m?todo, ? poss?vel definir v?rias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que s?o estruturadas, como as propriedades do objeto no qual o m?todo ? chamado.

> [!NOTE]
> O m?todo `set()` ? implementado apenas para objetos nas APIs JavaScript do Office espec?ficas de host, como a API JavaScript do Excel. As APIs comuns (compartilhadas) n?o d?o suporte a esse m?todo. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

As propriedades do objeto em que o m?todo ? chamado s?o definidas com os mesmos valores das propriedades correspondentes do objeto transmitido. Se o par?metro `properties` for um objeto JavaScript, as propriedades do objeto transmitido que correspondem ? propriedade de somente leitura no objeto em que o m?todo ? chamado ser?o ignoradas ou causar?o uma exce??o, dependendo do par?metro `options`.

#### <a name="syntax"></a>Sintaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Par?metros

|**Par?metro**|**Tipo**|**Descri??o**|
|:------------|:--------|:----------|
|`properties`|objeto|Um objeto do mesmo tipo de objeto do Office.js no qual o m?todo ? chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o m?todo ? chamado.|
|`options`|objeto|Opcional. S? pode ser transmitido quando o primeiro par?metro ? um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padr?o ? `true`: indicar um erro se o objeto JavaScript transmitido incluir propriedades de somente leitura.)|

#### <a name="returns"></a>Retorna

nulo    

#### <a name="example"></a>Exemplo

O exemplo de c?digo a seguir define v?rias propriedades do formato de um intervalo chamando o m?todo `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo sup?e que h? dados no intervalo **B2:E2**.

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
## <a name="42ornullobject-methods"></a>M?todos *OrNullObject

Muitos m?todos da API JavaScript do Excel retornar?o uma exce??o quando a condi??o da API n?o for atendida. Por exemplo, se voc? tentar obter uma planilha especificando um nome de planilha que n?o existe na pasta de trabalho, o m?todo `getItem()` retornar? uma exce??o `ItemNotFound`. 

Em vez de implementar a l?gica complexa de tratamento de exce??o para cen?rios como este, voc? pode usar a variante do m?todo `*OrNullObject` que est? dispon?vel para v?rios m?todos na API JavaScript do Excel. Um m?todo `*OrNullObject` retornar? um objeto nulo (n?o o `null` do JavaScript), em vez de emitir uma exce??o se o item especificado n?o existir. Por exemplo, voc? pode chamar o m?todo `getItemOrNullObject()` em uma cole??o, como **Worksheets**, para tentar recuperar um item da cole??o. O m?todo `getItemOrNullObject()` retornar? o item especificado se ele existir; caso contr?rio, ele retornar? um objeto nulo. O objeto nulo que ? retornado cont?m a propriedade booliana `isNullObject`, que voc? pode avaliar para determinar se o objeto existe.

O exemplo de c?digo a seguir tenta recuperar uma planilha chamada "Data" usando o m?todo `getItemOrNullObject()`. Se o m?todo retornar um objeto nulo, uma nova folha precisar? ser criada para que as a??es possam ser tomadas na folha.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data"); 

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) { 
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a>Veja tamb?m
 
* [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
* [Exemplos de c?digo de suplementos do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Otimiza??o de desempenho da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/performance.md)
* [Refer?ncia da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
