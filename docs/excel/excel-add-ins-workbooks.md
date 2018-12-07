---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: 1cfde9bfdf306e35f47595f936679d9fa6e1814e
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/30/2018
ms.locfileid: "27002334"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Trabalhar com pastas de trabalho usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos que o objeto **Workbook** suporta, confira [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook). Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).

O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel. Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados. O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais. Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas. O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.

## <a name="get-the-active-cell-or-selected-range"></a>Obter a célula ativa ou o intervalo selecionado

O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`. `getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range). O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

O método `getSelectedRange()` retorna o intervalo único selecionado atualmente. Se houver vários intervalos selecionados, será gerado um erro InvalidSelection. O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a>Criar uma pasta de trabalho

O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente. O objeto do Excel tem o método `createWorkbook` para esta finalidade. Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel. O suplemento permanece aberto e em execução com a pasta de trabalho anterior.

```js
Excel.createWorkbook();
```

O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente. O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional. A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.

Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada com Base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo. 

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var mybase64 = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(mybase64);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a>Proteger a estrutura da pasta de trabalho

O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho. A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`. O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho. 

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

O método `protect` aceita um parâmetro opcional de cadeia de caracteres. Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.

A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada. Para saber mais, confira a seção **Proteção de dados**do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).

> [!NOTE] 
> Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).

## <a name="access-document-properties"></a>Acessar propriedades do documentos

Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados. O exemplo a seguir mostra como definir a propriedade **author**.

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

Você também pode definir propriedades personalizadas. O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário. O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a>Acessar configurações do documentos

As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas. A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo. O exemplo a seguir mostra como criar e acessar uma configuração.

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="control-calculation-behavior"></a>Controlar o comportamento do cálculo

### <a name="set-calculation-mode"></a>Configurar o modo de cálculo

Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada. O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo. O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`. Esta propriedade pode ser configurada com os seguintes valores:

 - `automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.
 - `automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.
 - `manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.
 
### <a name="set-calculation-type"></a>Configurar o tipo de cálculo

O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato. `Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado. Os seguintes valores podem ser especificados:

 - `full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.
 - `fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.
 - `recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.
 
> [!NOTE] 
> Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Suspender os cálculos temporariamente

A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado. Isso é feito pelo `suspendApiCalculationUntilNextSync()`. Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md)
- [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md)