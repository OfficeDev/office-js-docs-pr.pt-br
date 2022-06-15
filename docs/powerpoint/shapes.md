---
title: Trabalhar com formas usando a POWERPOINT JavaScript
description: Saiba como adicionar, remover e formatar formas em PowerPoint slides.
ms.date: 06/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f314cfebb26450e79dbabe1e65ac9e4c8fe9799
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091101"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api"></a>Trabalhar com formas usando a POWERPOINT JavaScript

Este artigo descreve como usar formas geométricas, linhas e caixas de texto em conjunto com as APIs [Shape](/javascript/api/powerpoint/powerpoint.shape) e [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) .

## <a name="create-shapes"></a>Criar formas

As formas são criadas por meio e armazenadas na coleção de formas de um slide (`slide.shapes`). `ShapeCollection` tem vários `.add*` métodos para essa finalidade. Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção. Essas são as propriedades `name` e as `id` propriedades, respectivamente. `name` pode ser definido pelo suplemento.

### <a name="geometric-shapes"></a>Formas geométricas

Uma forma geométrica é criada com uma das sobrecargas de `ShapeCollection.addGeometricShape`. O primeiro parâmetro é uma enumeração [GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) ou a cadeia de caracteres equivalente a um dos valores da enumeração. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar o tamanho inicial da forma e sua posição em relação aos lados superior e esquerdo do slide, medido em pontos. Ou essas propriedades podem ser definidas depois que a forma é criada.

O exemplo de código a seguir cria um retângulo chamado **"Quadrado"** posicionado a 100 pontos dos lados superior e esquerdo do slide. O método retorna um `Shape` objeto.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### <a name="lines"></a>Linhas

Uma linha é criada com uma das sobrecargas de `ShapeCollection.addLine`. O primeiro parâmetro é uma enumeração [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) ou a cadeia de caracteres equivalente a um dos valores da enumeração para especificar como a linha contorts entre pontos de extremidade. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar os pontos inicial e final da linha. Ou essas propriedades podem ser definidas depois que a forma é criada. O método retorna um `Shape` objeto.

> [!NOTE]
> Quando a forma é uma linha, `top` `left` `Shape` `ShapeAddOptions` as propriedades e os objetos especificam o ponto inicial da linha em relação às bordas superior e esquerda do slide. As `height` propriedades `width` e o ponto de extremidade da linha *são especificados em relação ao ponto inicial*. Portanto, o ponto final relativo às bordas superior e esquerda do slide é (`top` + `height`) por ().`left` + `width` A unidade de medida para todas as propriedades é pontos e valores negativos são permitidos.

O exemplo de código a seguir cria uma linha reta no slide.

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### <a name="text-boxes"></a>Caixas de texto

Uma caixa de texto é criada com o [método addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) . O primeiro parâmetro é o texto que deve aparecer na caixa inicialmente. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar o tamanho inicial da caixa de texto e sua posição em relação aos lados superior e esquerdo do slide. Ou essas propriedades podem ser definidas depois que a forma é criada.

O exemplo de código a seguir mostra como criar uma caixa de texto no primeiro slide.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Mover e redimensionar formas

As formas ficam sobre o slide. Seu posicionamento é definido pelo e `left` pelas `top` propriedades. Elas atuam como margens das respectivas bordas do slide, medida em pontos, com `left: 0` `top: 0` e sendo o canto superior esquerdo. O tamanho da forma é especificado pelas propriedades `height` `width` . Seu código pode mover ou redimensionar a forma redefinindo essas propriedades. (Essas propriedades têm um significado ligeiramente diferente quando a forma é uma linha. Ver [Linhas](#lines).)

## <a name="text-in-shapes"></a>Texto em formas

Formas geométricas podem conter texto. As formas têm uma `textFrame` propriedade do tipo [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto). `TextFrame.textRange` é um [objeto TextRange](/javascript/api/powerpoint/powerpoint.textrange) com o conteúdo de texto e as configurações de fonte.

O exemplo de código a seguir cria uma forma geométrica chamada **"Chaves"** com o texto **"Texto da forma"**. Ele também ajusta as cores da forma e do texto, bem como define o alinhamento vertical do texto para o centro.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## <a name="delete-shapes"></a>Excluir formas

As formas são removidas do slide com `Shape` o método do `delete` objeto.

O exemplo de código a seguir mostra como excluir formas.

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```
