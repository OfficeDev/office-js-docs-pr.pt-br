---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: Exemplos de código que mostram como executar tarefas comuns com pastas de trabalho ou recursos de nível de aplicativo usando a API JavaScript do Excel.
ms.date: 05/06/2020
localization_priority: Normal
ms.openlocfilehash: 4fec6a217a2764eaf664463943ca384b3a2d847b
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170762"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Trabalhar com pastas de trabalho usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos `Workbook` quais o objeto oferece suporte, consulte [objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook). Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).

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

Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada em base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a>Inserir uma cópia de uma pasta de trabalho para a seção atual (visualização)

> [!NOTE]
> O método `WorksheetCollection.addFromBase64` no momento só está disponível na visualização pública e somente no Office no Windows e Mac. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho. Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento. Uma pasta de trabalho [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) tem o método `addFromBase64` para inserir cópias de planilhas da pasta de trabalho de destino nela mesma. O outro arquivo da pasta de trabalho é passado como em cadeia de caracteres codificado em base 64, como a chamada `Excel.createWorkbook`.

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

O exemplo a seguir mostra planilhas da pasta de trabalho que estão sendo inseridas em uma pasta de trabalho atual, logo após a planilha ativa. Observe que `null` é passado para o parâmetro `sheetNamesToInsert?: string[]`. Isso significa que todas as planilhas são inseridas.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a>Protege a estrutura da pasta de trabalho

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

Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados. O exemplo a seguir mostra como definir a `author` propriedade.

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

## <a name="access-application-culture-settings"></a>Configurações de cultura de aplicativo do Access

Uma pasta de trabalho tem configurações de idioma e cultura que afetam o modo como determinados dados são exibidos. Essas configurações podem ajudar a localizar dados quando os usuários do seu suplemento estiverem compartilhando pastas de trabalho em diferentes idiomas e culturas. O suplemento pode usar a análise de cadeia de caracteres para localizar o formato de números, datas e horas com base nas configurações de cultura do sistema para que cada usuário veja os dados em seu próprio formato de cultura.

`Application.cultureInfo`define as configurações de cultura do sistema como um objeto [CultureInfo](/javascript/api/excel/excel.cultureinfo) . Contém configurações como o separador decimal numérico ou o formato de data.

Algumas configurações de cultura podem ser [alteradas por meio da interface do usuário do Excel](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e). As configurações do sistema são preservadas `CultureInfo` no objeto. As alterações locais são mantidas como propriedades no nível do [aplicativo](/javascript/api/excel/excel.application), `Application.decimalSeparator`como.

O exemplo a seguir altera o caractere separador decimal de uma cadeia de caracteres numérica de um ', ' para o caractere usado pelas configurações do sistema.

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a>Adicionar dados XML personalizados à pasta de trabalho

O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho. Esses dados persistem na pasta de trabalho, independentemente do suplemento.

Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart). Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente. Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.

Os exemplos a seguir mostram como usar partes XML personalizadas. O primeiro bloco de códigos demonstra como inserir dados XML no documento. Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura. O segundo bloco mostra como acessar esse XML mais tarde. A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho. Os dados XML são então impressos no console.

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> `CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.

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

## <a name="save-the-workbook"></a>Salvar a pasta de trabalho

`Workbook.save` salva a pasta de trabalho para armazenamento persistente. O método `save` usa um parâmetro simples e opcional `saveBehavior` que pode ter um dos seguintes valores:

- `Excel.SaveBehavior.save` (padrão): o arquivo será salvo sem solicitar que o usuário especifique o nome do arquivo e local de salvamento. Se o arquivo não tiver sido salvo anteriormente, ele será salvo no local padrão. Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local.
- `Excel.SaveBehavior.prompt`: se o arquivo ainda não foi salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local de salvamento. Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local sem que o usuário seja solicitado.

> [!CAUTION]
> Se o usuário for solicitado a salvar e, em vez disso, cancelar a operação, `save` gera uma exceção.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>Fechar a pasta de trabalho

`Workbook.close` fecha a pasta de trabalho, além de suplementos que estão associados com a pasta de trabalho (o aplicativo Excel permanece aberto). O método `close` usa um parâmetro simples e opcional `closeBehavior` que pode ter um dos seguintes valores:

- `Excel.CloseBehavior.save` (padrão): o arquivo será salvo antes de fechar. Se o arquivo não tiver sido salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local para salvá-lo.
- `Excel.CloseBehavior.skipSave`: o arquivo é fechado imediatamente, sem ser salvo. Quaisquer alterações não salvas serão perdidas.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md)
- [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md)
