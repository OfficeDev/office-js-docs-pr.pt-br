---
title: Trabalhar com formas usando a POWERPOINT JavaScript
description: Saiba como adicionar, remover e formatar formas em PowerPoint slides.
ms.date: 10/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: e510ff47f4c54cd465be5c97c5828aad81041c5e
ms.sourcegitcommit: fb4a55764fb60e826ad06d15d1539e41df503b65
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/14/2021
ms.locfileid: "60356857"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>Trabalhar com formas usando a POWERPOINT JavaScript (visualização)

Este artigo descreve como usar formas geométricas, linhas e caixas de texto em conjunto com as APIs [Shape](/javascript/api/powerpoint/powerpoint.shape) e [ShapeCollection.](/javascript/api/powerpoint/powerpoint.shapecollection)

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>Criar formas

As formas são criadas por meio e armazenadas na coleção de formas de um slide ( `slide.shapes` ). `ShapeCollection` tem vários `.add*` métodos para essa finalidade. Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção. Estas são as `name` propriedades `id` e, respectivamente. `name` pode ser definido pelo seu complemento.

### <a name="geometric-shapes"></a>Formas geométricas

Uma forma geométrica é criada com uma das sobrecargas de `ShapeCollection.addGeometricShape` . O primeiro parâmetro é um número [GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) ou a cadeia de caracteres equivalente a um dos valores do número. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar o tamanho inicial da forma e sua posição em relação aos lados superior e esquerdo do slide, medido em pontos. Ou essas propriedades podem ser definidas depois que a forma for criada.

O exemplo de código a seguir cria um retângulo chamado **"Quadrado"** posicionado a 100 pontos dos lados superior e esquerdo do slide. O método retorna um `Shape` objeto.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
});
```

### <a name="lines"></a>Linhas

Uma linha é criada com uma das sobrecargas de `ShapeCollection.addLine` . O primeiro parâmetro é um número [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) ou a cadeia de caracteres equivalente a um dos valores do número para especificar como a linha se contorce entre pontos de extremidade. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar os pontos inicial e final da linha. Ou essas propriedades podem ser definidas depois que a forma for criada. O método retorna um `Shape` objeto.

> [!NOTE]
> Quando a forma é uma linha, as propriedades e dos objetos e especificam o ponto inicial da linha em relação às bordas superior e `top` `left` esquerda do `Shape` `ShapeAddOptions` slide. As `height` propriedades `width` e especificam o ponto de extremidade da *linha em relação ao ponto inicial*. Portanto, o ponto final em relação às bordas superior e esquerda do slide é ( `top`  +  `height` ) por ( `left`  +  `width` ). A unidade de medida para todas as propriedades é pontos e os valores negativos são permitidos.

O exemplo de código a seguir cria uma linha reta no slide.

```js
// This sample creates a straight line on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    return context.sync();
});
```

### <a name="text-boxes"></a>Caixas de texto

Uma caixa de texto é criada com o [método addTextBox.](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_) O primeiro parâmetro é o texto que deve aparecer na caixa inicialmente. Há um segundo parâmetro opcional do tipo [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) que pode especificar o tamanho inicial da caixa de texto e sua posição em relação aos lados superior e esquerdo do slide. Ou essas propriedades podem ser definidas depois que a forma for criada.

O exemplo de código a seguir mostra como criar uma caixa de texto no primeiro slide.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    return context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Mover e resize formas

As formas ficam na parte superior do slide. Seu posicionamento é definido pelas `left` propriedades `top` e. Elas atuam como margens das respectivas bordas do slide, medida em pontos, com e `left: 0` sendo o canto superior `top: 0` esquerdo. O tamanho da forma é especificado pelas `height` propriedades `width` e. Seu código pode mover ou reorganizar a forma redefinindo essas propriedades. (Essas propriedades têm um significado ligeiramente diferente quando a forma é uma linha. Consulte [Linhas](#lines).)

## <a name="text-in-shapes"></a>Texto em formas

Formas geométricas podem conter texto. As formas têm `textFrame` uma propriedade do tipo [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto). `TextFrame.textRange` é um [objeto TextRange](/javascript/api/powerpoint/powerpoint.textrange) com o conteúdo de texto e as configurações de fonte.

O exemplo de código a seguir cria uma forma geométrica chamada **"Chaves"** com o texto **"Shape text"**. Ele também ajusta as cores da forma e do texto, bem como define o alinhamento vertical do texto para o centro.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends and adds the purple text "Shape text" to the center.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    return context.sync();
});
```

## <a name="delete-shapes"></a>Excluir formas

As formas são removidas do slide com `Shape` o método do `delete` objeto.

O exemplo de código a seguir mostra como excluir formas.

```js
PowerPoint.run(function (context) {
    // Delete all shapes from the first slide.
    var sheet = context.presentation.slides.getItemAt(0);
    var shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync()
        .then(function () {
            shapes.items.forEach(function (shape) {
                shape.delete()
            });
            return context.sync();
        })
       .catch(errorHandlerFunction);
});
```
