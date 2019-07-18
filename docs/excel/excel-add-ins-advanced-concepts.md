---
title: Conceitos avançados de programação com a API JavaScript do Excel
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 4439ecf494a5d619e0d57604170c771e07b2e2b6
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771495"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Conceitos avançados de programação com a API JavaScript do Excel

Este artigo se baseia nas informações contidas em [conceitos fundamentais de programação API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avançados que são essenciais para a criação de suplementos complexos para o Excel 2016 ou posterior.

## <a name="officejs-apis-for-excel"></a>APIs Office.js para Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:

* **API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais. 

* **APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.

Enquanto você provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, você também usará objetos na API comum. Por exemplo:

- [Contexto](/javascript/api/office/office.context): o objeto **Context** representa o ambiente de tempo de execução do suplemento e oferece acesso aos principais objetos da API. Ele consiste em detalhes da configuração da pasta de trabalho, como `contentLanguage` e `officeTheme`, além de fornecer informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado é suportado pelo aplicativo Excel onde o suplemento está sendo executado. 

- [Document](/javascript/api/office/office.document): O objeto **Document** fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo Excel onde o suplemento está em execução. 

## <a name="requirement-sets"></a>Conjuntos de requisitos

Os conjuntos de requisitos são grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office dá suporte às APIs necessárias ao suplemento. Para identificar os conjuntos de requisitos específicos que estão disponíveis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificando o suporte ao conjunto de requisitos no tempo de execução

O exemplo de código a seguir mostra como determinar se o aplicativo host, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3') === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definindo o suporte ao conjunto de requisitos no manifesto

Você pode usar o [elemento Requirements](/office/dev/add-ins/reference/manifest/requirements) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o host do Office não der suporte aos conjuntos de requisitos ou aos métodos de API que são especificados no elemento **Requirements** do manifesto, o suplemento não será executado nesse host ou plataforma e não será exibido na lista de suplementos que são mostrados em **Meus Suplementos**. 

O exemplo de código a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos host do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel Online, Windows e iPad, é recomendável verificar o suporte a requisitos no tempo de execução, em vez de definir o suporte ao conjunto de requisitos no manifesto.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Conjuntos de requisitos para a API comum Office.js

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).

## <a name="loading-the-properties-of-an-object"></a>Carregando as propriedades de um objeto

Chamar o método `load()` em um objeto JavaScript do Excel orienta a API a carregar o objeto na memória do JavaScript quando o método `sync()` é executado. O método `load()` aceita uma cadeia de caracteres que contém nomes de propriedades delimitados por vírgulas a serem carregados ou um objeto que especifica propriedades a serem carregadas, opções de paginação, etc. 

> [!NOTE]
> Se você chamar o método `load()` em um objeto (ou uma coleção) sem especificar qualquer parâmetro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos na coleção) serão carregadas. Para reduzir a quantidade de transferência de dados entre o aplicativo host e o suplemento do Excel, você deve evitar chamar o método `load()` sem especificar explicitamente quais propriedades carregar.

### <a name="method-details"></a>Detalhes do método

#### <a name="loadparam-object"></a>load(param: object)

Preenche o objeto proxy criado na camada JavaScript com os valores da propriedade e do objeto especificados pelos parâmetros.

#### <a name="syntax"></a>Sintaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Parâmetros

|**Parâmetro**|**Tipo**|**Descrição**|
|:------------|:-------|:----------|
|`param`|objeto|Opcional. Aceita nomes de parâmetro e de relação como uma matriz ou cadeia de caracteres delimitada por vírgulas. Um objeto também pode ser passado para definir as propriedades de navegação e seleção (conforme mostrado no exemplo abaixo).|

#### <a name="returns"></a>Retorna

nulo

#### <a name="example"></a>Exemplo

O exemplo de código a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo. Observe que o objeto de origem deve ser carregado primeiro para que seus valores de propriedade possam ser acessados e gravados no intervalo de destino. Este exemplo pressupõe que há dados nos dois intervalos (**B2:E2** e **B7:E7**) e que os dois intervalos são inicialmente formatados de modo diferente.

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

### <a name="load-option-properties"></a>Carregar propriedades de opção

Como uma alternativa para passar uma cadeia de caracteres delimitada por vírgulas ou uma matriz ao chamar o método `load()`, você pode passar um objeto que contém as propriedades a seguir. 

|**Propriedade**|**Tipo**|**Descrição**|
|:-----------|:-------|:----------|
|`select`|objeto|Inclui uma lista delimitada por vírgula ou uma matriz de nomes de parâmetro/relação. Opcional.|
|`expand`|objeto|Inclui uma lista delimitada por vírgula ou uma matriz de nomes de relação. Opcional.|
|`top`|int| Especifica o número máximo de itens da coleção que podem ser incluídos no resultado. Opcional. Você só pode usar essa opção quando usar a opção de notação de objeto.|
|`skip`|int|Determina o número de itens da coleção que devem ser ignorados e não incluídos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados será iniciado depois de ignorar o número de itens especificado. Opcional. Você só pode usar esta opção ao usar a opção de notação de objeto.|

O exemplo de código a seguir carrega uma coleção de planilhas selecionando a `name`propriedade e o `address`do intervalo usado para cada planilha na coleção. Ele também especifica que apenas as cinco planilhas principais na coleção devem ser carregadas. Você poderia processar o próximo conjunto de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributo.

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propriedades escalares e de navegação

Na documentação de referência da API JavaScript do Excel, você pode notar que os membros do objeto são agrupados em duas categorias: **propriedades** e **relações**. Uma propriedade de um objeto é um membro escalar como uma cadeia de caracteres, um número inteiro ou um valor booliano, enquanto uma relação de um objeto (também conhecida como uma propriedade de navegação) é um membro que é ou um objeto ou uma coleção de objetos. Por exemplo, os membros `name` e `position` no objeto [Worksheet](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são relações (propriedades de navegação). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propriedades escalares e propriedades de navegação com `object.load()`

Chamar o método `object.load()` sem parâmetros especificados carregará todas as propriedades escalares do objeto; as propriedades de navegação do objeto não serão carregadas. Além disso, as propriedades de navegação não podem ser carregadas diretamente. Em vez disso, você deve usar o método `load()` para fazer referência às propriedades escalares individuais na propriedade de navegação desejada. Por exemplo, para carregar o nome da fonte de um intervalo, você deve especificar as propriedades de navegação **format** e **font** como o caminho para a propriedade **name**:

```js
someRange.load("format/font/name")
```

> [!NOTE]
> com a API JavaScript do Excel, é possível definir propriedades escalares de uma propriedade de navegação percorrendo o caminho. Por exemplo, é possível definir o tamanho da fonte de um intervalo usando `someRange.format.font.size = 10;`. Não é necessário carregar a propriedade antes de configurá-la. 

## <a name="setting-properties-of-an-object"></a>Definindo propriedades de um objeto

A definição de propriedades em um objeto com propriedades de navegação aninhadas pode ser uma tarefa complicada. Como uma alternativa para definir propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o método `object.set()` que está disponível em todos os objetos na API JavaScript do Excel. Com esse método, é possível definir várias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que são estruturadas, como as propriedades do objeto no qual o método é chamado.

> [!NOTE]
> O método `set()` é implementado apenas para objetos nas APIs JavaScript do Office específicas de host, como a API JavaScript do Excel. As APIs comuns (compartilhadas) não dão suporte a esse método. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

As propriedades do objeto em que o método é chamado são definidas com os mesmos valores das propriedades correspondentes do objeto transmitido. Se o parâmetro `properties` for um objeto JavaScript, as propriedades do objeto transmitido que correspondem à propriedade de somente leitura no objeto em que o método é chamado serão ignoradas ou causarão uma exceção, dependendo do parâmetro `options`.

#### <a name="syntax"></a>Sintaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Parâmetros

|**Parâmetro**|**Tipo**|**Descrição**|
|:------------|:--------|:----------|
|`properties`|objeto|Um objeto do mesmo tipo de objeto do Office.js no qual o método é chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o método é chamado.|
|`options`|objeto|Opcional. Só pode ser transmitido quando o primeiro parâmetro é um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padrão é `true`: indicar um erro se o objeto JavaScript transmitido incluir propriedades de somente leitura.)|

#### <a name="returns"></a>Retorna

nulo

#### <a name="example"></a>Exemplo

O exemplo de código a seguir define várias propriedades do formato de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo supõe que há dados no intervalo **B2:E2**.

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

## <a name="42ornullobject-methods"></a>Métodos &#42;OrNullObject

Muitos métodos da API JavaScript do Excel retornarão uma exceção quando a condição da API não for atendida. Por exemplo, se você tentar obter uma planilha especificando um nome de planilha que não existe na pasta de trabalho, o método `getItem()` retornará uma exceção `ItemNotFound`. 

Em vez de implementar a lógica complexa de tratamento de exceção para cenários como este, você pode usar a variante do método `*OrNullObject` que está disponível para vários métodos na API JavaScript do Excel. Um método `*OrNullObject` retornará um objeto nulo (não o `null` do JavaScript), em vez de emitir uma exceção se o item especificado não existir. Por exemplo, você pode chamar o método `getItemOrNullObject()` em uma coleção, como **Worksheets**, para tentar recuperar um item da coleção. O método `getItemOrNullObject()` retornará o item especificado se ele existir; caso contrário, ele retornará um objeto nulo. O objeto nulo que é retornado contém a propriedade booliana `isNullObject`, que você pode avaliar para determinar se o objeto existe.

O exemplo de código a seguir tenta recuperar uma planilha chamada "Data" usando o método `getItemOrNullObject()`. Se o método retornar um objeto nulo, uma nova folha precisará ser criada para que as ações possam ser tomadas na folha.

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

## <a name="see-also"></a>Confira também

* [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
* [Exemplos de código de suplementos do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Otimização de desempenho da API JavaScript do Excel](performance.md)
* [Referência da API JavaScript do Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
