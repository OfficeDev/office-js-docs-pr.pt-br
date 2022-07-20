---
title: Trabalhar com formas usando a API JavaScript do Excel
description: Saiba como o Excel define formas como qualquer objeto que fique na camada de desenho do Excel.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 507ae05b570e7eef4f3bf5560ca47c1bfbd40f9f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889594"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Trabalhar com formas usando a API JavaScript do Excel

O Excel define formas como qualquer objeto que fique na camada de desenho do Excel. Isso significa que qualquer coisa fora de uma célula é uma forma. Este artigo descreve como usar formas geométricas, linhas e imagens em conjunto com as APIs [Shape](/javascript/api/excel/excel.shape) e [ShapeCollection](/javascript/api/excel/excel.shapecollection) . [Os](/javascript/api/excel/excel.chart) gráficos são abordados em seu próprio artigo, [Trabalhar com gráficos usando a API JavaScript do Excel](excel-add-ins-charts.md).

A imagem a seguir mostra formas que formam um termômetro.
![Imagem de um termômetro feito como uma forma do Excel.](../images/excel-shapes.png)

## <a name="create-shapes"></a>Criar formas

As formas são criadas por meio e armazenadas na coleção de formas de uma planilha (`Worksheet.shapes`). `ShapeCollection` tem vários `.add*` métodos para essa finalidade. Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção. Essas são as propriedades `name` e as `id` propriedades, respectivamente. `name` pode ser definido pelo suplemento para recuperação fácil com o `ShapeCollection.getItem(name)` método.

Os tipos de formas a seguir são adicionados usando o método associado.

| Shape | Add Method | Assinatura |
|-------|------------|-----------|
| Forma Geométrica | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Imagem (JPEG ou PNG) | [Addimage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [Addline](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Caixa de Texto | [Addtextbox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Formas geométricas

Uma forma geométrica é criada com `ShapeCollection.addGeometricShape`. Esse método usa uma [enumeração GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) como um argumento.

O exemplo de código a seguir cria um retângulo de 150 x 150 pixels chamado **"Quadrado"** posicionado a 100 pixels dos lados superior e esquerdo da planilha.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### <a name="images"></a>Imagens

As imagens JPEG, PNG e SVG podem ser inseridas em uma planilha como formas. O `ShapeCollection.addImage` método usa uma cadeia de caracteres codificada em base64 como um argumento. Essa é uma imagem JPEG ou PNG no formato de cadeia de caracteres. `ShapeCollection.addSvg` também usa uma cadeia de caracteres, embora esse argumento seja XML que define o gráfico.

O exemplo de código a seguir mostra um arquivo de imagem sendo carregado por [um FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) como uma cadeia de caracteres. A cadeia de caracteres tem os metadados "base64", removidos antes que a forma seja criada.

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Linhas

Uma linha é criada com `ShapeCollection.addLine`. Esse método precisa das margens esquerda e superior dos pontos inicial e final da linha. Ele também usa uma [enumeração ConnectorType](/javascript/api/excel/excel.connectortype) para especificar como a linha contorts entre pontos de extremidade. O exemplo de código a seguir cria uma linha reta na planilha.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

As linhas podem ser conectadas a outros objetos Shape. O `connectBeginShape` e `connectEndShape` os métodos anexam o início e o término de uma linha às formas nos pontos de conexão especificados. Os locais desses pontos variam de acordo com a forma, `Shape.connectionSiteCount` mas eles podem ser usados para garantir que seu suplemento não se conecte a um ponto fora dos limites. Uma linha é desconectada de qualquer forma anexada usando os `disconnectBeginShape` métodos `disconnectEndShape` e os métodos.

O exemplo de código a seguir conecta a **linha "MyLine"** a duas formas chamadas **"LeftShape"** e **"RightShape"**.

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Mover e redimensionar formas

As formas ficam sobre a planilha. Seu posicionamento é definido pela propriedade `left` e pela `top` propriedade. Elas atuam como margens das respectivas bordas da planilha, sendo [0, 0] o canto superior esquerdo. Eles podem ser definidos diretamente ou ajustados de sua posição atual com os `incrementLeft` métodos `incrementTop` e os métodos. O quanto uma forma é girada da posição padrão também é estabelecida dessa maneira, `rotation` `incrementRotation` com a propriedade sendo a quantidade absoluta e o método ajustando a rotação existente.

A profundidade de uma forma em relação a outras formas é definida pela `zorderPosition` propriedade. Isso é definido usando o método `setZOrder` , que usa [um ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` ajusta a ordenação da forma atual em relação às outras formas.

Seu suplemento tem algumas opções para alterar a altura e a largura das formas. Definir a ou `height` a propriedade `width` altera a dimensão especificada sem alterar a outra dimensão. O `scaleHeight` e `scaleWidth` ajuste as respectivas dimensões da forma em relação ao tamanho atual ou original (com base no valor do [ShapeScaleType fornecido](/javascript/api/excel/excel.shapescaletype)). Um parâmetro [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) opcional especifica de onde a forma é dimensionada (canto superior esquerdo, meio ou canto inferior direito). Se a `lockAspectRatio` propriedade for `true`, os métodos de escala manterão a taxa de proporção atual da forma também ajustando a outra dimensão.

> [!NOTE]
> As alterações diretas nas `height` propriedades e `width` afetam apenas essa propriedade, independentemente `lockAspectRatio` do valor da propriedade.

O exemplo de código a seguir mostra uma forma sendo dimensionada para 1,25 vezes seu tamanho original e girada 30 graus.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## <a name="text-in-shapes"></a>Texto em formas

Formas Geométricas podem conter texto. As formas têm uma `textFrame` propriedade do tipo [TextFrame](/javascript/api/excel/excel.textframe). O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto). `TextFrame.textRange` é um [objeto TextRange](/javascript/api/excel/excel.textrange) com o conteúdo de texto e as configurações de fonte.

O exemplo de código a seguir cria uma forma geométrica chamada "Wave" com o texto "Texto da Forma". Ele também ajusta as cores da forma e do texto, bem como define o alinhamento horizontal do texto para o centro.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

O `addTextBox` método de criar `ShapeCollection` um tipo com `GeometricShape` uma `Rectangle` tela de fundo branca e um texto preto. Isso é o mesmo que o que é criado pelo botão Caixa de Texto **do** Excel na **guia** Inserir. `addTextBox` usa um argumento de cadeia de caracteres para definir o texto de `TextRange`.

O exemplo de código a seguir mostra a criação de uma caixa de texto com o texto "Olá!".

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="shape-groups"></a>Grupos de formas

As formas podem ser agrupadas. Isso permite que um usuário os trate como uma única entidade para posicionamento, dimensionamento e outras tarefas relacionadas. Um [ShapeGroup](/javascript/api/excel/excel.shapegroup) é um tipo de `Shape`, portanto, seu suplemento trata o grupo como uma única forma.

O exemplo de código a seguir mostra três formas sendo agrupadas. O exemplo de código subsequente mostra que o grupo de formas está sendo movido para os 50 pixels à direita.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> As formas individuais dentro do grupo são referenciadas por meio da `ShapeGroup.shapes` propriedade, que é do tipo [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). Eles não são mais acessíveis por meio da coleção de formas da planilha após serem agrupados. Por exemplo, se sua planilha tivesse três formas e todas elas fossem agrupadas, `shapes.getCount` o método da planilha retornaria uma contagem de 1.

## <a name="export-shapes-as-images"></a>Exportar formas como imagens

Qualquer `Shape` objeto pode ser convertido em uma imagem. [Shape.getAsImage retorna](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) a cadeia de caracteres codificada em base64. O formato da imagem é especificado como uma [enumeração PictureFormat](/javascript/api/excel/excel.pictureformat) passada para `getAsImage`.

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
});
```

## <a name="delete-shapes"></a>Excluir formas

As formas são removidas da planilha com `Shape` o método do `delete` objeto. Nenhum outro metadado é necessário.

O exemplo de código a seguir exclui todas as formas de **MyWorksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Trabalhar com gráficos usando a API JavaScript do Excel](excel-add-ins-charts.md)
