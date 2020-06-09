---
title: Trabalhar com formas usando a API JavaScript do Excel
description: Saiba como o Excel define formas como qualquer objeto que se encontra na camada de desenho do Excel.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 7b9a4dba02e28187eeb0f932e245489ca61fcbcc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609738"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a><span data-ttu-id="120d8-103">Trabalhar com formas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="120d8-103">Work with shapes using the Excel JavaScript API</span></span>

<span data-ttu-id="120d8-104">O Excel define formas como qualquer objeto que esteja na camada de desenho do Excel.</span><span class="sxs-lookup"><span data-stu-id="120d8-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="120d8-105">Isso significa que algo fora de uma célula é uma forma.</span><span class="sxs-lookup"><span data-stu-id="120d8-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="120d8-106">Este artigo descreve como usar formas geométricas, linhas e imagens em conjunto com as APIs [Shape](/javascript/api/excel/excel.shape) e [ShapeCollection](/javascript/api/excel/excel.shapecollection) .</span><span class="sxs-lookup"><span data-stu-id="120d8-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="120d8-107">Os [gráficos](/javascript/api/excel/excel.chart) são abordados em seu próprio artigo, [trabalhar com gráficos usando a API JavaScript do Excel](excel-add-ins-charts.md).</span><span class="sxs-lookup"><span data-stu-id="120d8-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

<span data-ttu-id="120d8-108">A imagem a seguir mostra formas que formam um termômetro.</span><span class="sxs-lookup"><span data-stu-id="120d8-108">The following image shows shapes which form a thermometer.</span></span>
<span data-ttu-id="120d8-109">![Imagem de um termômetro criado como uma forma do Excel](../images/excel-shapes.png)</span><span class="sxs-lookup"><span data-stu-id="120d8-109">![Image of a thermometer made as an Excel shape](../images/excel-shapes.png)</span></span>

## <a name="create-shapes"></a><span data-ttu-id="120d8-110">Criar formas</span><span class="sxs-lookup"><span data-stu-id="120d8-110">Create shapes</span></span>

<span data-ttu-id="120d8-111">As formas são criadas e armazenadas na coleção Shape de uma planilha ( `Worksheet.shapes` ).</span><span class="sxs-lookup"><span data-stu-id="120d8-111">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="120d8-112">`ShapeCollection`tem vários `.add*` métodos para essa finalidade.</span><span class="sxs-lookup"><span data-stu-id="120d8-112">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="120d8-113">Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção.</span><span class="sxs-lookup"><span data-stu-id="120d8-113">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="120d8-114">São as `name` Propriedades e `id` , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="120d8-114">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="120d8-115">`name`pode ser definido pelo suplemento para fácil recuperação com o `ShapeCollection.getItem(name)` método.</span><span class="sxs-lookup"><span data-stu-id="120d8-115">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="120d8-116">Os seguintes tipos de formas são adicionados usando o método associado:</span><span class="sxs-lookup"><span data-stu-id="120d8-116">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="120d8-117">Forma</span><span class="sxs-lookup"><span data-stu-id="120d8-117">Shape</span></span> | <span data-ttu-id="120d8-118">Add Method</span><span class="sxs-lookup"><span data-stu-id="120d8-118">Add Method</span></span> | <span data-ttu-id="120d8-119">Assinatura</span><span class="sxs-lookup"><span data-stu-id="120d8-119">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="120d8-120">Forma geométrica</span><span class="sxs-lookup"><span data-stu-id="120d8-120">Geometric Shape</span></span> | [<span data-ttu-id="120d8-121">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="120d8-121">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="120d8-122">Imagem (JPEG ou PNG)</span><span class="sxs-lookup"><span data-stu-id="120d8-122">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="120d8-123">AddImage</span><span class="sxs-lookup"><span data-stu-id="120d8-123">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="120d8-124">Linha</span><span class="sxs-lookup"><span data-stu-id="120d8-124">Line</span></span> | [<span data-ttu-id="120d8-125">addLine</span><span class="sxs-lookup"><span data-stu-id="120d8-125">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="120d8-126">SVG</span><span class="sxs-lookup"><span data-stu-id="120d8-126">SVG</span></span> | [<span data-ttu-id="120d8-127">addSvg</span><span class="sxs-lookup"><span data-stu-id="120d8-127">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="120d8-128">Caixa de Texto</span><span class="sxs-lookup"><span data-stu-id="120d8-128">Text Box</span></span> | [<span data-ttu-id="120d8-129">addTextBox</span><span class="sxs-lookup"><span data-stu-id="120d8-129">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="120d8-130">Formas geométricas</span><span class="sxs-lookup"><span data-stu-id="120d8-130">Geometric shapes</span></span>

<span data-ttu-id="120d8-131">Uma forma geométrica é criada com o `ShapeCollection.addGeometricShape` .</span><span class="sxs-lookup"><span data-stu-id="120d8-131">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="120d8-132">Esse método utiliza uma enumeração [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) como um argumento.</span><span class="sxs-lookup"><span data-stu-id="120d8-132">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="120d8-133">O exemplo de código a seguir cria um retângulo de 150x150 chamado **"Square"** que é posicionado 100 pixels a partir da parte superior e esquerda dos lados da planilha.</span><span class="sxs-lookup"><span data-stu-id="120d8-133">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="120d8-134">Imagens</span><span class="sxs-lookup"><span data-stu-id="120d8-134">Images</span></span>

<span data-ttu-id="120d8-135">Imagens JPEG, PNG e SVG podem ser inseridas em uma planilha como formas.</span><span class="sxs-lookup"><span data-stu-id="120d8-135">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="120d8-136">O `ShapeCollection.addImage` método usa uma cadeia de caracteres codificada em base64 como um argumento.</span><span class="sxs-lookup"><span data-stu-id="120d8-136">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="120d8-137">É uma imagem JPEG ou PNG no formato de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="120d8-137">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="120d8-138">`ShapeCollection.addSvg`o também usa uma cadeia de caracteres, embora esse argumento seja XML que define o gráfico.</span><span class="sxs-lookup"><span data-stu-id="120d8-138">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="120d8-139">O exemplo de código a seguir mostra um arquivo de imagem sendo carregado por um [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="120d8-139">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="120d8-140">A cadeia de caracteres tem os metadados "base64", removidos antes da forma ser criada.</span><span class="sxs-lookup"><span data-stu-id="120d8-140">The string has the metadata "base64," removed before the shape is created.</span></span>

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

### <a name="lines"></a><span data-ttu-id="120d8-141">Linhas</span><span class="sxs-lookup"><span data-stu-id="120d8-141">Lines</span></span>

<span data-ttu-id="120d8-142">Uma linha é criada com o `ShapeCollection.addLine` .</span><span class="sxs-lookup"><span data-stu-id="120d8-142">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="120d8-143">Esse método precisa das margens esquerda e superior dos pontos inicial e final da linha.</span><span class="sxs-lookup"><span data-stu-id="120d8-143">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="120d8-144">Ele também utiliza um [ConnectorType](/javascript/api/excel/excel.connectortype) enum para especificar como a linha se deformará entre pontos de extremidade.</span><span class="sxs-lookup"><span data-stu-id="120d8-144">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="120d8-145">O exemplo de código a seguir cria uma linha reta na planilha.</span><span class="sxs-lookup"><span data-stu-id="120d8-145">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="120d8-146">As linhas podem ser conectadas a outros objetos Shape.</span><span class="sxs-lookup"><span data-stu-id="120d8-146">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="120d8-147">Os `connectBeginShape` `connectEndShape` métodos e anexam o início e o fim de uma linha a formas nos pontos de conexão especificados.</span><span class="sxs-lookup"><span data-stu-id="120d8-147">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="120d8-148">Os locais desses pontos variam por forma, mas o `Shape.connectionSiteCount` pode ser usado para garantir que seu suplemento não se conecte a um ponto que está fora do limite.</span><span class="sxs-lookup"><span data-stu-id="120d8-148">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="120d8-149">Uma linha é desconectada de formas anexadas usando os `disconnectBeginShape` `disconnectEndShape` métodos e.</span><span class="sxs-lookup"><span data-stu-id="120d8-149">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="120d8-150">O exemplo de código a seguir conecta a linha **"myline"** a duas formas chamadas **"LeftShape"** e **"RightShape"**.</span><span class="sxs-lookup"><span data-stu-id="120d8-150">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="120d8-151">Mover e redimensionar formas</span><span class="sxs-lookup"><span data-stu-id="120d8-151">Move and resize shapes</span></span>

<span data-ttu-id="120d8-152">As formas ficam na parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="120d8-152">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="120d8-153">O posicionamento é definido pela `left` propriedade e `top` .</span><span class="sxs-lookup"><span data-stu-id="120d8-153">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="120d8-154">Eles atuam como margens das bordas da planilha, com [0, 0] sendo o canto superior esquerdo.</span><span class="sxs-lookup"><span data-stu-id="120d8-154">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="120d8-155">Eles podem ser definidos diretamente ou ajustados a partir da sua posição atual com os `incrementLeft` `incrementTop` métodos e.</span><span class="sxs-lookup"><span data-stu-id="120d8-155">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="120d8-156">O quanto uma forma é girada a partir da posição padrão também é estabelecida dessa maneira, com a `rotation` propriedade sendo o valor absoluto e o `incrementRotation` método ajustando a rotação existente.</span><span class="sxs-lookup"><span data-stu-id="120d8-156">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="120d8-157">A profundidade de uma forma em relação a outras formas é definida pela `zorderPosition` propriedade.</span><span class="sxs-lookup"><span data-stu-id="120d8-157">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="120d8-158">Isso é definido usando o `setZOrder` método, que usa um [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span><span class="sxs-lookup"><span data-stu-id="120d8-158">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="120d8-159">`setZOrder`ajusta a ordenação da forma atual em relação às outras formas.</span><span class="sxs-lookup"><span data-stu-id="120d8-159">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="120d8-160">O suplemento tem algumas opções para alterar a altura e a largura das formas.</span><span class="sxs-lookup"><span data-stu-id="120d8-160">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="120d8-161">A definição da `height` `width` Propriedade or altera a dimensão especificada sem alterar a outra dimensão.</span><span class="sxs-lookup"><span data-stu-id="120d8-161">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="120d8-162">O `scaleHeight` e `scaleWidth` ajustam as respectivas dimensões da forma em relação ao tamanho atual ou original (com base no valor do [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)fornecido).</span><span class="sxs-lookup"><span data-stu-id="120d8-162">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="120d8-163">Um parâmetro [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) opcional especifica de onde a forma é dimensionada (canto superior esquerdo, médio ou canto inferior direito).</span><span class="sxs-lookup"><span data-stu-id="120d8-163">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="120d8-164">Se a `lockAspectRatio` propriedade for **true**, os métodos Scale mantêm a taxa de proporção atual da forma ajustando também a outra dimensão.</span><span class="sxs-lookup"><span data-stu-id="120d8-164">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="120d8-165">As alterações diretas `height` nas `width` Propriedades e só afetam essa propriedade, independentemente do `lockAspectRatio` valor da propriedade.</span><span class="sxs-lookup"><span data-stu-id="120d8-165">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="120d8-166">O exemplo de código a seguir mostra uma forma que está sendo dimensionada para 1,25 vezes seu tamanho original e girado 30 graus.</span><span class="sxs-lookup"><span data-stu-id="120d8-166">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="120d8-167">Texto em formas</span><span class="sxs-lookup"><span data-stu-id="120d8-167">Text in shapes</span></span>

<span data-ttu-id="120d8-168">Formas geométricas podem conter texto.</span><span class="sxs-lookup"><span data-stu-id="120d8-168">Geometric Shapes can contain text.</span></span> <span data-ttu-id="120d8-169">As formas têm uma `textFrame` Propriedade do tipo [TextFrame](/javascript/api/excel/excel.textframe).</span><span class="sxs-lookup"><span data-stu-id="120d8-169">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="120d8-170">O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto).</span><span class="sxs-lookup"><span data-stu-id="120d8-170">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="120d8-171">`TextFrame.textRange`é um objeto [TextRange](/javascript/api/excel/excel.textrange) com o conteúdo de texto e as configurações de fonte.</span><span class="sxs-lookup"><span data-stu-id="120d8-171">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="120d8-172">O exemplo de código a seguir cria uma forma geométrica chamada "Wave" com o texto "Shape Text".</span><span class="sxs-lookup"><span data-stu-id="120d8-172">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="120d8-173">Também ajusta a forma e as cores do texto, bem como define o alinhamento horizontal do texto para o centro.</span><span class="sxs-lookup"><span data-stu-id="120d8-173">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="120d8-174">O `addTextBox` método de `ShapeCollection` criar um `GeometricShape` tipo `Rectangle` com um texto de fundo branco e preto.</span><span class="sxs-lookup"><span data-stu-id="120d8-174">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="120d8-175">É o mesmo que o que é criado pelo botão **caixa de texto** do Excel na guia **Inserir** . `addTextBox` leva um argumento de cadeia de caracteres para definir o texto do `TextRange` .</span><span class="sxs-lookup"><span data-stu-id="120d8-175">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="120d8-176">O exemplo de código a seguir mostra a criação de uma caixa de texto com o texto "Hello!".</span><span class="sxs-lookup"><span data-stu-id="120d8-176">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="120d8-177">Grupos de formas</span><span class="sxs-lookup"><span data-stu-id="120d8-177">Shape groups</span></span>

<span data-ttu-id="120d8-178">As formas podem ser agrupadas juntas.</span><span class="sxs-lookup"><span data-stu-id="120d8-178">Shapes can be grouped together.</span></span> <span data-ttu-id="120d8-179">Isso permite que um usuário o trate como uma única entidade de posicionamento, dimensionamento e outras tarefas relacionadas.</span><span class="sxs-lookup"><span data-stu-id="120d8-179">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="120d8-180">Um grupo de [formas](/javascript/api/excel/excel.shapegroup) é um tipo de `Shape` , para que o suplemento trate o grupo como uma única forma.</span><span class="sxs-lookup"><span data-stu-id="120d8-180">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="120d8-181">O exemplo de código a seguir mostra três formas que estão sendo agrupadas.</span><span class="sxs-lookup"><span data-stu-id="120d8-181">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="120d8-182">O exemplo de código subsequente mostra que o grupo de formas sendo movido para a direita 50 pixels.</span><span class="sxs-lookup"><span data-stu-id="120d8-182">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="120d8-183">As formas individuais dentro do grupo são referenciadas por meio da `ShapeGroup.shapes` propriedade, que é do tipo [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span><span class="sxs-lookup"><span data-stu-id="120d8-183">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="120d8-184">Eles não podem mais ser acessados por meio da coleção Shape da planilha após serem agrupados.</span><span class="sxs-lookup"><span data-stu-id="120d8-184">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="120d8-185">Por exemplo, se a sua planilha tivesse três formas e todos foram agrupadas, o método da planilha `shapes.getCount` retornará uma contagem de 1.</span><span class="sxs-lookup"><span data-stu-id="120d8-185">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="120d8-186">Exportar formas como imagens</span><span class="sxs-lookup"><span data-stu-id="120d8-186">Export shapes as images</span></span>

<span data-ttu-id="120d8-187">Qualquer `Shape` objeto pode ser convertido em uma imagem.</span><span class="sxs-lookup"><span data-stu-id="120d8-187">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="120d8-188">[Shape. getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) retorna Cadeia de caracteres codificada em base64.</span><span class="sxs-lookup"><span data-stu-id="120d8-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="120d8-189">O formato da imagem é especificado como um enum [PictureFormat](/javascript/api/excel/excel.pictureformat) passado para `getAsImage` .</span><span class="sxs-lookup"><span data-stu-id="120d8-189">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="120d8-190">Excluir formas</span><span class="sxs-lookup"><span data-stu-id="120d8-190">Delete shapes</span></span>

<span data-ttu-id="120d8-191">As formas são removidas da planilha com o `Shape` método do objeto `delete` .</span><span class="sxs-lookup"><span data-stu-id="120d8-191">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="120d8-192">Nenhum outro metadado é necessário.</span><span class="sxs-lookup"><span data-stu-id="120d8-192">No other metadata is needed.</span></span>

<span data-ttu-id="120d8-193">O exemplo de código a seguir exclui todas as formas de **myworksheet**.</span><span class="sxs-lookup"><span data-stu-id="120d8-193">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="120d8-194">Confira também</span><span class="sxs-lookup"><span data-stu-id="120d8-194">See also</span></span>

- [<span data-ttu-id="120d8-195">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="120d8-195">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="120d8-196">Trabalhar com gráficos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="120d8-196">Work with charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
