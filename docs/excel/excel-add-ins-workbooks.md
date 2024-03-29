---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: Saiba como executar tarefas comuns com as guias de trabalho ou recursos no nível do aplicativo usando Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f003c59ab3fcd029d16bde2ca95cd3a4fdbd15b9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745466"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Trabalhar com pastas de trabalho usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel. Para ver a lista completa de propriedades e `Workbook` métodos que o objeto oferece suporte, consulte [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook). Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).

O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel. Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados. O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais. Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas. O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.

## <a name="get-the-active-cell-or-selected-range"></a>Obter a célula ativa ou o intervalo selecionado

O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`. `getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range). O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.

```js
await Excel.run(async (context) => {
    let activeCell = context.workbook.getActiveCell();
    activeCell.load("address");
    await context.sync();

    console.log("The active cell is " + activeCell.address);
});
```

O método `getSelectedRange()` retorna o intervalo único selecionado atualmente. Se houver vários intervalos selecionados, será gerado um erro InvalidSelection. O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    await context.sync();
});
```

## <a name="create-a-workbook"></a>Criar uma pasta de trabalho

O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente. O objeto do Excel tem o método `createWorkbook` para esta finalidade. Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel. O suplemento permanece aberto e em execução com a pasta de trabalho anterior.

```js
Excel.createWorkbook();
```

O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente. O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional. A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.

Você pode obter a pasta de trabalho atual do seu complemento como uma cadeia de caracteres codificada com base64 usando o [fatiamento de arquivo](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)). A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
let myFile = document.getElementById("file");
let reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        let startIndex = reader.result.toString().indexOf("base64,");
        let externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    });
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a>Inserir uma cópia de uma pasta de trabalho para a seção atual

O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho. Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento. Uma [pasta de](/javascript/api/excel/excel.workbook) trabalho tem o `insertWorksheetsFromBase64` método para inserir cópias das planilhas da pasta de trabalho de destino em si. O arquivo da outra pasta de trabalho é passado como uma cadeia de caracteres codificada com base64, assim como a `Excel.createWorkbook` chamada.

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> O `insertWorksheetsFromBase64` método é suportado para Excel no Windows, Mac e na Web. Não há suporte para iOS. Além disso, Excel na Web esse método não dá suporte a planilhas de origem com elementos PivotTable, Chart, Comment ou Slicer. Se esses objetos estão presentes, o `insertWorksheetsFromBase64` método retorna o `UnsupportedFeature` erro em Excel na Web.

O exemplo de código a seguir mostra como inserir planilhas de outra pasta de trabalho na pasta de trabalho atual. [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) Este exemplo de código primeiro processa um arquivo de pasta de trabalho com um objeto e extrai uma cadeia de caracteres codificada com base64 e insere essa cadeia de caracteres codificada com base64 na pasta de trabalho atual. As novas planilhas são inseridas após a planilha chamada **Sheet1**. Observe que `[]` é passado como o parâmetro para a [propriedade InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member) . Isso significa que todas as planilhas da pasta de trabalho de destino são inseridas na pasta de trabalho atual.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
let myFile = document.getElementById("file");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        let startIndex = reader.result.toString().indexOf("base64,");
        let externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        let workbook = context.workbook;
            
        // Set up the insert options. 
        let options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a>Protege a estrutura da pasta de trabalho

O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho. A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`. O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.

```js
await Excel.run(async (context) => {
    let workbook = context.workbook;
    workbook.load("protection/protected");
    await context.sync();

    if (!workbook.protection.protected) {
        workbook.protection.protect();
    }
});
```

O método `protect` aceita um parâmetro opcional de cadeia de caracteres. Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.

A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada. Para saber mais, confira a seção **Proteção de dados** do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).

> [!NOTE]
> Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.microsoft.com/office/7e365a4d-3e89-4616-84ca-1931257c1517).

## <a name="access-document-properties"></a>Acessar propriedades do documentos

Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.microsoft.com/office/21d604c2-481e-4379-8e54-1dd4622c6b75). A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados. O exemplo a seguir mostra como definir a `author` propriedade.

```js
await Excel.run(async (context) => {
    let docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    await context.sync();
});
```

### <a name="custom-properties"></a>Propriedades personalizadas

Você também pode definir propriedades personalizadas. O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário. O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.

```js
await Excel.run(async (context) => {
    let customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    await context.sync();
});

[...]

await Excel.run(async (context) => {
    let customDocProperties = context.workbook.properties.custom;
    let customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key, value"]);
    await context.sync();

    console.log("Custom key  : " + customProperty.key); // "Introduction"
    console.log("Custom value : " + customProperty.value); // "Hello"
});
```

#### <a name="worksheet-level-custom-properties"></a>Propriedades personalizadas no nível da planilha

As propriedades personalizadas também podem ser definidas no nível da planilha. Elas são semelhantes às propriedades personalizadas no nível do documento, exceto que a mesma chave pode ser repetida em planilhas diferentes. O exemplo a seguir mostra como criar uma propriedade personalizada chamada **WorksheetGroup** com o valor "Alfa" na planilha atual e, em seguida, recuperá-la.

```js
await Excel.run(async (context) => {
    // Add the custom property.
    let customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    await context.sync();
});

[...]

await Excel.run(async (context) => {
    // Load the keys and values of all custom properties in the current worksheet.
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    let customWorksheetProperties = worksheet.customProperties;
    let customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    await context.sync();

    // Log the WorksheetGroup custom property to the console.
    console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
    console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
});
```

## <a name="access-document-settings"></a>Acessar configurações do documentos

As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas. A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo. O exemplo a seguir mostra como criar e acessar uma configuração.

```js
await Excel.run(async (context) => {
    let settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    let needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    await context.sync();
    console.log("Workbook needs review : " + needsReview.value);
});
```

## <a name="access-application-culture-settings"></a>Configurações de cultura de aplicativos do Access

Uma workbook tem configurações de idioma e cultura que afetam a forma como determinados dados são exibidos. Essas configurações podem ajudar a localização de dados quando os usuários do seu complemento estão compartilhando as guias de trabalho em diferentes idiomas e culturas. Seu complemento pode usar a análise de cadeia de caracteres para localizar o formato de números, datas e horas com base nas configurações de cultura do sistema para que cada usuário veja dados no formato de sua própria cultura.

`Application.cultureInfo` define as configurações de cultura do sistema como um [objeto CultureInfo](/javascript/api/excel/excel.cultureinfo) . Isso contém configurações como o separador decimal numérico ou o formato de data.

Algumas configurações de cultura podem ser [alteradas por meio da interface Excel interface do usuário](https://support.microsoft.com/office/c093b545-71cb-4903-b205-aebb9837bd1e). As configurações do sistema são preservadas no `CultureInfo` objeto. Quaisquer alterações locais são mantidas como [propriedades no](/javascript/api/excel/excel.application) nível do aplicativo, como `Application.decimalSeparator`.

O exemplo a seguir altera o caractere separador decimal de uma cadeia numérica de um ',' para o caractere usado pelas configurações do sistema.

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let decimalSource = sheet.getRange("B2");

    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");
    await context.sync();

    let systemDecimalSeparator =
        context.application.cultureInfo.numberFormat.numberDecimalSeparator;
    let oldDecimalString = decimalSource.values[0][0];

    // This assumes the input column is standardized to use "," as the decimal separator.
    let newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

    let resultRange = sheet.getRange("C2");
    resultRange.values = [[newDecimalString]];
    resultRange.format.autofitColumns();
    await context.sync();
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a>Adicionar dados XML personalizados à pasta de trabalho

O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho. Esses dados persistem na pasta de trabalho, independentemente do suplemento.

Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart). Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente. Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.

Os exemplos a seguir mostram como usar partes XML personalizadas. O primeiro bloco de códigos demonstra como inserir dados XML no documento. Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura. O segundo bloco mostra como acessar esse XML mais tarde. A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho. Os dados XML são então impressos no console.

```js
await Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    let originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    let customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");
    await context.sync();

    // Store the XML part's ID in a setting
    let settings = context.workbook.settings;
    settings.add("ContosoReviewXmlPartId", customXmlPart.id);
});
```

```js
await Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    let settings = context.workbook.settings;
    let xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    await context.sync();

    if (xmlPartIDSetting.value) {
        let customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
        let xmlBlob = customXmlPart.getXml();

        await context.sync();

        // Add spaces to make it more human-readable in the console.
        let readableXML = xmlBlob.value.replace(/></g, "> <");
        console.log(readableXML);
    }
});
```

> [!NOTE]
> `CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.

## <a name="control-calculation-behavior"></a>Controlar o comportamento do cálculo

### <a name="set-calculation-mode"></a>Configurar o modo de cálculo

Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada. O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo. O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`. Ele pode ser definido para os seguintes valores.

- `automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.
- `automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.
- `manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.

### <a name="set-calculation-type"></a>Configurar o tipo de cálculo

O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato. `Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado. Os valores a seguir podem ser especificados.

- `full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.
- `fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.
- `recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.

> [!NOTE]
> Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Suspender os cálculos temporariamente

A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado. Isso é feito pelo `suspendApiCalculationUntilNextSync()`. Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation"></a>Detectar a ativação de uma agenda de trabalho

O seu complemento pode detectar quando uma workbook é ativada. Uma workbook fica *inativa* quando o usuário alterna o foco para outra workbook, para outro aplicativo ou (em Excel na Web) para outra guia do navegador da Web. Uma workbook *é ativada quando* o usuário retorna o foco para a workbook. A ativação da workbook pode disparar funções de retorno de chamada no seu complemento, como atualizar dados da agenda de trabalho.

Para detectar quando uma caixa de trabalho é ativada, [registre](excel-add-ins-events.md#register-an-event-handler) um manipulador de eventos para o [evento onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) de uma workbook. Os manipuladores de eventos do evento `onActivated` recebem um [objeto WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) quando o evento é acionado.

> [!IMPORTANT]
> O `onActivated` evento não detecta quando uma workbook é aberta. Esse evento só detecta quando um usuário alterna o foco de volta para uma workbook já aberta.

O exemplo de código a seguir mostra como registrar o manipulador `onActivated` de eventos e configurar uma função de retorno de chamada.

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve the workbook.
        let workbook = context.workbook;
    
        // Register the workbook activated event handler.
        workbook.onActivated.add(workbookActivated);
        await context.sync();
    });
}

async function workbookActivated(event) {
    await Excel.run(async (context) => {
        // Retrieve the workbook and load the name.
        let workbook = context.workbook;
        workbook.load("name");        
        await context.sync();

        // Callback function for when the workbook is activated.
        console.log(`The workbook ${workbook.name} was activated.`);
    });
}
```

## <a name="save-the-workbook"></a>Salvar a pasta de trabalho

`Workbook.save` salva a pasta de trabalho para armazenamento persistente. O `save` método tem um único parâmetro opcional `saveBehavior` que pode ser um dos seguintes valores.

- `Excel.SaveBehavior.save` (padrão): o arquivo será salvo sem solicitar que o usuário especifique o nome do arquivo e local de salvamento. Se o arquivo não tiver sido salvo anteriormente, ele será salvo no local padrão. Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local.
- `Excel.SaveBehavior.prompt`: se o arquivo ainda não foi salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local de salvamento. Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local sem que o usuário seja solicitado.

> [!CAUTION]
> Se o usuário for solicitado a salvar e, em vez disso, cancelar a operação, `save` gera uma exceção.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>Fechar a pasta de trabalho

`Workbook.close` fecha a pasta de trabalho, além de suplementos que estão associados com a pasta de trabalho (o aplicativo Excel permanece aberto). O `close` método tem um único parâmetro opcional `closeBehavior` que pode ser um dos seguintes valores.

- `Excel.CloseBehavior.save` (padrão): o arquivo será salvo antes de fechar. Se o arquivo não tiver sido salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local para salvá-lo.
- `Excel.CloseBehavior.skipSave`: o arquivo é fechado imediatamente, sem ser salvo. Quaisquer alterações não salvas serão perdidas.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md)
