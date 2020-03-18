---
title: Trabalhar com formas usando a API JavaScript do Excel
description: Saiba como o Excel define formas como qualquer objeto que se encontra na camada de desenho do Excel.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 7522bf440389e983efc3ec696375694e5539c442
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717114"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Trabalhar com formas usando a API JavaScript do Excel

O Excel define formas como qualquer objeto que esteja na camada de desenho do Excel. Isso significa que algo fora de uma célula é uma forma. Este artigo descreve como usar formas geométricas, linhas e imagens em conjunto com as APIs [Shape](/javascript/api/excel/excel.shape) e [ShapeCollection](/javascript/api/excel/excel.shapecollection) . Os [gráficos](/javascript/api/excel/excel.chart) são abordados em seu próprio artigo, [trabalhar com gráficos usando a API JavaScript do Excel](excel-add-ins-charts.md).

A imagem a seguir mostra formas que formam um termômetro.
![Imagem de um termômetro criado como uma forma do Excel](../images/excel-shapes.png)

## <a name="create-shapes"></a>Criar formas

As formas são criadas e armazenadas na coleção Shape de uma planilha (`Worksheet.shapes`). `ShapeCollection`tem vários `.add*` métodos para essa finalidade. Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção. São as `name` Propriedades e `id` , respectivamente. `name`pode ser definido pelo suplemento para fácil recuperação com o `ShapeCollection.getItem(name)` método.

Os seguintes tipos de formas são adicionados usando o método associado:

| Forma | Add Method | Assinatura |
|-------|------------|-----------|
| Forma geométrica | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Imagem (JPEG ou PNG) | [AddImage](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| Linha | [addLine](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| Caixa de Texto | [addTextBox](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Formas geométricas

Uma forma geométrica é criada `ShapeCollection.addGeometricShape`com o. Esse método utiliza uma enumeração [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) como um argumento.

O exemplo de código a seguir cria um retângulo de 150x150 chamado **"Square"** que é posicionado 100 pixels a partir da parte superior e esquerda dos lados da planilha.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="images"></a>Imagens

Imagens JPEG, PNG e SVG podem ser inseridas em uma planilha como formas. O `ShapeCollection.addImage` método usa uma cadeia de caracteres codificada em base64 como um argumento. É uma imagem JPEG ou PNG no formato de cadeia de caracteres. `ShapeCollection.addSvg`o também usa uma cadeia de caracteres, embora esse argumento seja XML que define o gráfico.

O exemplo de código a seguir mostra um arquivo de imagem sendo carregado por um [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) como uma cadeia de caracteres. A cadeia de caracteres tem os metadados "base64", removidos antes da forma ser criada.

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Linhas

Uma linha é criada com `ShapeCollection.addLine`o. Esse método precisa das margens esquerda e superior dos pontos inicial e final da linha. Ele também utiliza um [ConnectorType](/javascript/api/excel/excel.connectortype) enum para especificar como a linha se deformará entre pontos de extremidade. O exemplo de código a seguir cria uma linha reta na planilha.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

As linhas podem ser conectadas a outros objetos Shape. Os `connectBeginShape` métodos `connectEndShape` e anexam o início e o fim de uma linha a formas nos pontos de conexão especificados. Os locais desses pontos variam por forma, mas o `Shape.connectionSiteCount` pode ser usado para garantir que seu suplemento não se conecte a um ponto que está fora do limite. Uma linha é desconectada de formas anexadas `disconnectBeginShape` usando `disconnectEndShape` os métodos e.

O exemplo de código a seguir conecta a linha **"myline"** a duas formas chamadas **"LeftShape"** e **"RightShape"**.

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-and-resize-shapes"></a>Mover e redimensionar formas

As formas ficam na parte superior da planilha. O `left` posicionamento é definido pela propriedade e `top` . Eles atuam como margens das bordas da planilha, com [0, 0] sendo o canto superior esquerdo. Eles podem ser definidos diretamente ou ajustados a partir da sua posição atual `incrementLeft` com `incrementTop` os métodos e. O quanto uma forma é girada a partir da posição padrão também é estabelecida dessa maneira, com a `rotation` propriedade sendo o valor absoluto e o `incrementRotation` método ajustando a rotação existente.

A profundidade de uma forma em relação a outras formas é definida `zorderPosition` pela propriedade. Isso é definido usando o `setZOrder` método, que usa um [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder`ajusta a ordenação da forma atual em relação às outras formas.

O suplemento tem algumas opções para alterar a altura e a largura das formas. A definição da `height` propriedade `width` or altera a dimensão especificada sem alterar a outra dimensão. O `scaleHeight` e `scaleWidth` ajustam as respectivas dimensões da forma em relação ao tamanho atual ou original (com base no valor do [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)fornecido). Um parâmetro [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) opcional especifica de onde a forma é dimensionada (canto superior esquerdo, médio ou canto inferior direito). Se a `lockAspectRatio` propriedade for **true**, os métodos Scale mantêm a taxa de proporção atual da forma ajustando também a outra dimensão.

> [!NOTE]
> As alterações diretas `height` nas `width` Propriedades e só afetam essa propriedade, independentemente do `lockAspectRatio` valor da propriedade.

O exemplo de código a seguir mostra uma forma que está sendo dimensionada para 1,25 vezes seu tamanho original e girado 30 graus.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="text-in-shapes"></a>Texto em formas

Formas geométricas podem conter texto. As formas têm `textFrame` uma propriedade do tipo [TextFrame](/javascript/api/excel/excel.textframe). O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto). `TextFrame.textRange`é um objeto [TextRange](/javascript/api/excel/excel.textrange) com o conteúdo de texto e as configurações de fonte.

O exemplo de código a seguir cria uma forma geométrica chamada "Wave" com o texto "Shape Text". Também ajusta a forma e as cores do texto, bem como define o alinhamento horizontal do texto para o centro.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    return context.sync();
}).catch(errorHandlerFunction);
```

O `addTextBox` método de `ShapeCollection` criar um `GeometricShape` tipo `Rectangle` com um texto de fundo branco e preto. É o mesmo que o que é criado pelo botão **caixa de texto** do Excel na guia **Inserir** . `addTextBox` Obtém um argumento de cadeia de caracteres para definir o `TextRange`texto do.

O exemplo de código a seguir mostra a criação de uma caixa de texto com o texto "Hello!".

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="shape-groups"></a>Grupos de formas

As formas podem ser agrupadas juntas. Isso permite que um usuário o trate como uma única entidade de posicionamento, dimensionamento e outras tarefas relacionadas. Um grupo de [formas](/javascript/api/excel/excel.shapegroup) é um tipo `Shape`de, para que o suplemento trate o grupo como uma única forma.

O exemplo de código a seguir mostra três formas que estão sendo agrupadas. O exemplo de código subsequente mostra que o grupo de formas sendo movido para a direita 50 pixels.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var square = shapes.getItem("Square");
    var pentagon = shapes.getItem("Pentagon");
    var octagon = shapes.getItem("Octagon");

    var shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    return context.sync();
}).catch(errorHandlerFunction);

// This sample moves the previously created shape group to the right by 50 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shapeGroup = sheet.shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> As formas individuais dentro do grupo são referenciadas `ShapeGroup.shapes` por meio da propriedade, que é do tipo [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection). Eles não podem mais ser acessados por meio da coleção Shape da planilha após serem agrupados. Por exemplo, se a sua planilha tivesse três formas e todos foram agrupadas, o método da `shapes.getCount` planilha retornará uma contagem de 1.

## <a name="export-shapes-as-images"></a>Exportar formas como imagens

Qualquer `Shape` objeto pode ser convertido em uma imagem. [Shape. getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) retorna Cadeia de caracteres codificada em base64. O formato da imagem é especificado como um enum [PictureFormat](/javascript/api/excel/excel.pictureformat) passado para `getAsImage`.

```js
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shape = sheet.shapes.getItem("Image");
    var stringResult = shape.getAsImage(Excel.PictureFormat.png);

    return context.sync().then(function () {
        console.log(stringResult.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

## <a name="delete-shapes"></a>Excluir formas

As formas são removidas da planilha `Shape` com o `delete` método do objeto. Nenhum outro metadado é necessário.

O exemplo de código a seguir exclui todas as formas de **myworksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Trabalhar com gráficos usando a API JavaScript do Excel](excel-add-ins-charts.md)
