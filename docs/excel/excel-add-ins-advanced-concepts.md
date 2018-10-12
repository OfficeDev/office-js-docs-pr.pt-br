---
title: Conceitos avançados de programação com a API JavaScript do Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 190eb65e45ce246009b6d85d378571bd2f451e0b
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459249"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Conceitos avançados de programação com a API JavaScript do Excel

Este artigo foi criado com base nas informações em [Conceitos fundamentais da API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avançados que são essenciais para criar suplementos complexos para o Excel 2016 ou posterior.

## <a name="officejs-apis-for-excel"></a>APIs Office.js para Excel

Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:

* **API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais. 

* **APIs comuns**: introduzidas com o Office 2013, as APIs comuns (também conhecidas como [API Compartilhada](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)) podem ser usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos host, como Word, Excel ou PowerPoint.

Enquanto você provavelmente usará a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos direcionados ao Excel 2016 ou posterior , você também usará objetos da API Compartilhada. Por exemplo:

- [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): o objeto **Context** representa o ambiente em tempo de execução do suplemento e fornece acesso aos principais objetos da API. Ele consiste em detalhes de configuração de pasta de trabalho como `contentLanguage` e `officeTheme` e também fornece informações sobre ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado é suportado pelo aplicativo Excel em que o suplemento está sendo executado. 

- [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js): o objeto **Document** fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo Excel onde o suplemento está sendo executado. 

## <a name="requirement-sets"></a>Conjuntos de requisitos

Conjuntos de requisitos são grupos nomeados de membros da API. Um suplemento do Office pode executar uma verificação em tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office suporta as APIs que o suplemento precisa. Para identificar os conjuntos de requisitos específico que estão disponíveis em cada plataforma compatível, confira [Conjuntos de requisitos da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificar compatibilidade com o conjunto de requisitos em tempo de execução

O exemplo de código a seguir mostra como determinar se o aplicativo host, onde o suplemento está sendo executado é compatível com o conjunto de requisitos especificado pela API .

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definir a compatibilidade com o conjunto de requisitos no manifesto

Você pode usar o [elemento Requirements](https://docs.microsoft.com/javascript/office/manifest/requirements?view=office-js) no manifesto do suplemento para especificar o conjunto mínimo de requisitos e/ou os métodos de API que o seu suplemento necessita para ser ativado. Se o host do Office ou outra plataforma não for compatível com o  conjunto de requisitos ou métodos da API especificados no elemento **Requirements** do manifesto, o suplemento não será executado nessa plataforma ou host e não será exibido na lista **Meus Suplementos**. 

O exemplo de código a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica que o suplemento deve ser carregado em todos os aplicativos host do Office compatíveis com o  conjunto de requisitos ExcelApi, versão 1.3 ou superior.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel para Windows, Excel Online e Excel para iPad, é recomendável verificar a compatibilidade com os requisitos em tempo de execução, em vez de defini-lo no manifesto.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Conjuntos de requisitos para a API comum Office.js

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js).

## <a name="loading-the-properties-of-an-object"></a>Carregar as propriedades de um objeto

Chamar o método `load()` em um objeto JavaScript do Excel instrui a API para carregar o objeto na memória do JavaScript quando o método `sync()` for executado. O método `load()` aceita uma sequência de caracteres que contém os nomes das propriedades que devem ser carregadas delimitadas por vírgulas ou um objeto que especifica as propriedades a carregar, opções de paginação, etc. 

> [!NOTE]
> Se você chamar o método `load()` em um objeto (ou coleção) sem especificar nenhum parâmetro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos da coleção) serão carregadas. Para reduzir o volume de dados transferidos entre o Excel e o suplemento, você deve evitar chamadas ao método `load()`  sem especificar explicitamente quais propriedades para carregar.

### <a name="method-details"></a>Detalhes do método

#### <a name="loadparam-object"></a>load(param: object)

Preenche o objeto de proxy criado na camada JavaScript com os valores de propriedades e objetos especificados pelos parâmetros.

#### <a name="syntax"></a>Sintaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Parâmetros

|**Parâmetro**|**Tipo**|**Descrição**|
|:------------|:-------|:----------|
|`param`|objeto|Opcional. Aceita nomes de parâmetro e relacionamento como uma sequência de caracteres delimitadas por vírgulas ou uma matriz. Um objeto também pode ser passado para definir as propriedades de seleção e navegação (conforme mostrado no exemplo a seguir).|

#### <a name="returns"></a>Retorna

nulo

#### <a name="example"></a>Exemplo

O exemplo de código a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo. Observe que o objeto de origem deve ser carregado primeiro para que os valores de suas propriedades possam ser acessados e gravados no intervalo de destino. Este exemplo pressupõe que haja dados os dois intervalos (**B2:E2** e **B7:E7**) e os dois intervalos tenham inicialmente uma formatação diferente.

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

### <a name="load-option-properties"></a>Opções de propriedades de carga

Como alternativa para passar uma sequência de caracteres delimitada por vírgulas ou uma matriz ao chamar o método `load()`, você pode passar um objeto que contém as propriedades a seguir. 

|**Propriedade**|**Tipo**|**Descrição**|
|:-----------|:-------|:----------|
|`select`|objeto|Inclui uma lista delimitada por vírgulas ou uma matriz de nomes de parâmetro/relacionamentos. Opcional.|
|`expand`|objeto|Inclui uma lista delimitada por vírgulas ou uma matriz de nomes de relacionamentos. Opcional.|
|`top`|int| Especifica o número máximo de itens da coleção que podem ser incluídos no resultado. Opcional. Você só pode usar essa opção quando usar a opção de notação de objeto.|
|`skip`|int|Determina o número de itens da coleção que devem ser ignorados e não incluídos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados será iniciado depois de ignorar o número de itens especificado. Opcional. Você só pode usar esta opção ao usar a opção de notação de objeto.|

O exemplo de código a seguir carrega uma coleção planilhas selecionando a propriedade `name` e o intervalo `address` do intervalo utilizado cada planilha na coleção. Ela também especifica que apenas as cinco primeiras planilhas na coleção devem ser carregadas. Você poderia processar o conjunto seguinte de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributos. 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propriedades escalares e de navegação 

Na documentação de referência de API JavaScript do Excel, você pode observar que os objetos membros são agrupados em duas categorias: **propriedades** e **relacionamentos**. Uma propriedade de um objeto é um membro escalar como uma sequência de caracteres, um número inteiro ou um valor booleano, enquanto o relacionamento de um objeto (também conhecido como propriedade de navegação) é um membro que é um objeto ou a coleção de objetos. Por exemplo, os membros `name` e `position` do objeto [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) são propriedades escalares, enquanto `protection` e `tables` são relacionamentos (propriedades de navegação). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propriedades escalares e propriedades de navegação com `object.load()`

Chamar o método `object.load()` sem especificar parâmetros carregará todas as propriedades escalares do objeto; as propriedades de navegação do objeto não serão carregadas. Além disso, não é possível carregar as propriedades de navegação diretamente. Em vez disso, você deve usar o método  `load()` para referenciar propriedades escalares individualmente dentro das propriedade de navegação desejadas. Por exemplo, para carregar o nome da fonte de um intervalo, você deve especificar as propriedades de navegação **format** e **font** como caminho para a propriedade **name** :

```js
someRange.load("format/font/name")
```

> [!NOTE]
> Com a API JavaScript do Excel, você pode definir propriedades escalares de uma propriedade de navegação percorrendo o caminho. Por exemplo, você pode definir o tamanho da fonte para um intervalo usando `someRange.format.font.size = 10;`. Você não precisará carregar a propriedade antes de defini-la. 

## <a name="setting-properties-of-an-object"></a>Definir propriedades de um objeto

Definir propriedades em um objeto com propriedades de navegação aninhadas pode ser complicado. Como alternativa à definição de propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o método `object.set()` que está disponível em todos os objetos na API JavaScript do Excel. Com este método, você pode definir várias propriedades de um objeto ao mesmo tempo, passando um outro objeto do mesmo tipo da Office.js ou um objeto JavaScript com propriedades que são estruturadas como as propriedades do objeto no qual o método é chamado.

> [!NOTE]
> O método `set()` é implementado apenas para objetos dentro das APIs JavaScript do Office específicas do host, como a API JavaScript do Excel. As APIs (compartilhadas) comuns não suportam esse método. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

As propriedades do objeto em que o método é chamado são definidas com os mesmos valores das propriedades correspondentes do objeto passado. Se o parâmetro `properties` for um objeto JavaScript, as propriedades do objeto passado que correspondem à propriedades somente para leitura no objeto em que o método é chamado serão ignoradas ou causarão uma exceção, dependendo do parâmetro `options`.

#### <a name="syntax"></a>Sintaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Parâmetros

|**Parâmetro**|**Tipo**|**Descrição**|
|:------------|:--------|:----------|
|`properties`|objeto|Um objeto do mesmo tipo de objeto da Office.js no qual o método é chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o método é chamado.|
|`options`|objeto|Opcional. Só pode ser passado quando o primeiro parâmetro é um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padrão é `true`: indicar um erro se o objeto JavaScript passado incluir propriedades com acesso de somente leitura.)|

#### <a name="returns"></a>Retorna

nulo    

#### <a name="example"></a>Exemplo

O exemplo de código a seguir define várias propriedades de formatação de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo supõe que há dados no intervalo **B2:E2**.

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
## <a name="42ornullobject-methods"></a>Métodos *OrNullObject

Muitos métodos da API JavaScript do Excel retornam uma exceção quando a condição da API não é atendida. Por exemplo, se você tentar obter uma planilha, especificando um nome de planilha que não existe na pasta de trabalho, o método  `getItem()` retornará uma exceção de `ItemNotFound`. 

Em vez de implementar uma lógica complexa de tratamento de exceções em cenários como esse, você pode usar a variante do método `*OrNullObject`  que está disponível para vários métodos na API JavaScript do Excel. Um método `*OrNullObject` retornará um objeto nulo (não o JavaScript `null`) em vez de gerar uma exceção se o item especificado não existir. Por exemplo, você pode chamar o método `getItemOrNullObject()` em um conjunto como **planilhas** para tentar recuperar um item da coleção. O método  `getItemOrNullObject()` retorna o item especificado se ele existir. Caso contrário, retorna um objeto nulo. O objeto nulo retornado contém a propriedade booleana `isNullObject` que pode ser avaliada para determinar se o objeto existe.

O exemplo de código a seguir tenta recuperar uma planilha chamada "Data" usando o método `getItemOrNullObject()` . Se o método retornar um objeto nulo, uma nova planilha deve ser criada para que as ações sejam executadas.

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
* [Exemplos de códigos de suplementos do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Otimização de desempenho da API JavaScript do Excel](performance.md)
* [Referência da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
