---
title: Trabalhar com formas usando a API JavaScript do Excel
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: e4d01c387fff01d68cb26369240a1e06e723a54c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448259"
---
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a><span data-ttu-id="efc00-102">Trabalhar com formas usando a API JavaScript do Excel (visualização)</span><span class="sxs-lookup"><span data-stu-id="efc00-102">Work with Shapes using the Excel JavaScript API (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="efc00-103">As APIs abordadas neste artigo atualmente estão disponíveis somente na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="efc00-103">The APIs discussed in this article are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="efc00-104">O Excel define formas como qualquer objeto que esteja na camada de desenho do Excel.</span><span class="sxs-lookup"><span data-stu-id="efc00-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="efc00-105">Isso significa que algo fora de uma célula é uma forma.</span><span class="sxs-lookup"><span data-stu-id="efc00-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="efc00-106">Este artigo descreve como usar formas geométricas, linhas e imagens em conjunto com as APIs [Shape]/JavaScript/API/Excel/Excel.Shape) [](/javascript/api/excel/excel.shapecollection) e ShapeCollection.</span><span class="sxs-lookup"><span data-stu-id="efc00-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape]/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="efc00-107">Os [gráficos](/javascript/api/excel/excel.chart) são abordados em seus próprios artigos, [trabalhar com gráficos usando a API JavaScript do Excel] (Excel-Add-ins-Charts.MD)).</span><span class="sxs-lookup"><span data-stu-id="efc00-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with Charts using the Excel JavaScript API]](excel-add-ins-charts.md)).</span></span>

## <a name="create-shapes"></a><span data-ttu-id="efc00-108">Criar formas</span><span class="sxs-lookup"><span data-stu-id="efc00-108">Create shapes</span></span>

<span data-ttu-id="efc00-109">As formas são criadas e armazenadas na coleção Shape de uma planilha (`Worksheet.shapes`).</span><span class="sxs-lookup"><span data-stu-id="efc00-109">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="efc00-110">`ShapeCollection`tem vários `.add*` métodos para essa finalidade.</span><span class="sxs-lookup"><span data-stu-id="efc00-110">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="efc00-111">Todas as formas têm nomes e IDs gerados para elas quando são adicionadas à coleção.</span><span class="sxs-lookup"><span data-stu-id="efc00-111">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="efc00-112">São as `name` Propriedades e `id` , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="efc00-112">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="efc00-113">`name`pode ser definido pelo suplemento para fácil recuperação com o `ShapeCollection.getItem(name)` método.</span><span class="sxs-lookup"><span data-stu-id="efc00-113">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="efc00-114">Os seguintes tipos de formas são adicionados usando o método associado:</span><span class="sxs-lookup"><span data-stu-id="efc00-114">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="efc00-115">Forma</span><span class="sxs-lookup"><span data-stu-id="efc00-115">Shape</span></span> | <span data-ttu-id="efc00-116">Add Method</span><span class="sxs-lookup"><span data-stu-id="efc00-116">Add Method</span></span> | <span data-ttu-id="efc00-117">Assinatura</span><span class="sxs-lookup"><span data-stu-id="efc00-117">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="efc00-118">Forma geoMétrica</span><span class="sxs-lookup"><span data-stu-id="efc00-118">Geometric Shape</span></span> | [<span data-ttu-id="efc00-119">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="efc00-119">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="efc00-120">Imagem (JPEG ou PNG)</span><span class="sxs-lookup"><span data-stu-id="efc00-120">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="efc00-121">addImage</span><span class="sxs-lookup"><span data-stu-id="efc00-121">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="efc00-122">Linha</span><span class="sxs-lookup"><span data-stu-id="efc00-122">Line</span></span> | [<span data-ttu-id="efc00-123">addLine</span><span class="sxs-lookup"><span data-stu-id="efc00-123">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="efc00-124">SVG</span><span class="sxs-lookup"><span data-stu-id="efc00-124">SVG</span></span> | [<span data-ttu-id="efc00-125">addSvg</span><span class="sxs-lookup"><span data-stu-id="efc00-125">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="efc00-126">Caixa de Texto</span><span class="sxs-lookup"><span data-stu-id="efc00-126">Text Box</span></span> | [<span data-ttu-id="efc00-127">addTextBox</span><span class="sxs-lookup"><span data-stu-id="efc00-127">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="efc00-128">Formas geoMétricas</span><span class="sxs-lookup"><span data-stu-id="efc00-128">Geometric shapes</span></span>

<span data-ttu-id="efc00-129">Uma forma geométrica é criada `ShapeCollection.addGeometricShape`com o.</span><span class="sxs-lookup"><span data-stu-id="efc00-129">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="efc00-130">Esse método utiliza uma enumeração [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) como um argumento.</span><span class="sxs-lookup"><span data-stu-id="efc00-130">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="efc00-131">O exemplo de código a seguir cria um retângulo de 150x150 chamado **"Square"** que é posicionado 100 pixels a partir da parte superior e esquerda dos lados da planilha.</span><span class="sxs-lookup"><span data-stu-id="efc00-131">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="efc00-132">Imagens</span><span class="sxs-lookup"><span data-stu-id="efc00-132">Images</span></span>

<span data-ttu-id="efc00-133">Imagens JPEG, PNG e SVG podem ser inseridas em uma planilha como formas.</span><span class="sxs-lookup"><span data-stu-id="efc00-133">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="efc00-134">O `ShapeCollection.addImage` método usa uma cadeia de caracteres codificada em base64 como um argumento.</span><span class="sxs-lookup"><span data-stu-id="efc00-134">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="efc00-135">É uma imagem JPEG ou PNG no formato de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc00-135">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="efc00-136">`ShapeCollection.addSvg`o também usa uma cadeia de caracteres, embora esse argumento seja XML que define o gráfico.</span><span class="sxs-lookup"><span data-stu-id="efc00-136">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="efc00-137">O exemplo de código a seguir mostra um arquivo de imagem sendo [](https://developer.mozilla.org/docs/Web/API/FileReader) carregado por um FileReader como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc00-137">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="efc00-138">A cadeia de caracteres tem os metadados "base64", removidos antes da forma ser criada.</span><span class="sxs-lookup"><span data-stu-id="efc00-138">The string has the metadata "base64," removed before the shape is created.</span></span>

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = event.target.result.indexOf("base64,");
        var myBase64 = event.target.result.substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a><span data-ttu-id="efc00-139">Linhas</span><span class="sxs-lookup"><span data-stu-id="efc00-139">Lines</span></span>

<span data-ttu-id="efc00-140">Uma linha é criada com `ShapeCollection.addLine`o.</span><span class="sxs-lookup"><span data-stu-id="efc00-140">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="efc00-141">Esse método precisa das margens esquerda e superior dos pontos inicial e final da linha.</span><span class="sxs-lookup"><span data-stu-id="efc00-141">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="efc00-142">Ele também utiliza um [ConnectorType](/javascript/api/excel/excel.connectortype) enum para especificar como a linha se deformará entre pontos de extremidade.</span><span class="sxs-lookup"><span data-stu-id="efc00-142">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="efc00-143">O exemplo de código a seguir cria uma linha reta na planilha.</span><span class="sxs-lookup"><span data-stu-id="efc00-143">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="efc00-144">As linhas podem ser conectadas a outros objetos Shape.</span><span class="sxs-lookup"><span data-stu-id="efc00-144">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="efc00-145">Os `connectBeginShape` métodos `connectEndShape` e anexam o início e o fim de uma linha a formas nos pontos de conexão especificados.</span><span class="sxs-lookup"><span data-stu-id="efc00-145">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="efc00-146">Os locais desses pontos variam por forma, mas o `Shape.connectionSiteCount` pode ser usado para garantir que seu suplemento não se conecte a um ponto que está fora do limite.</span><span class="sxs-lookup"><span data-stu-id="efc00-146">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="efc00-147">Uma linha é desconectada de formas anexadas `disconnectBeginShape` usando `disconnectEndShape` os métodos e.</span><span class="sxs-lookup"><span data-stu-id="efc00-147">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="efc00-148">O exemplo de código a seguir conecta a linha **"myline"** a duas formas chamadas **"LeftShape"** e **"RightShape"**.</span><span class="sxs-lookup"><span data-stu-id="efc00-148">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="efc00-149">Mover e redimensionar formas</span><span class="sxs-lookup"><span data-stu-id="efc00-149">Move and resize shapes</span></span>

<span data-ttu-id="efc00-150">As formas ficam na parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="efc00-150">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="efc00-151">O `left` posicionamento é definido pela propriedade e `top` .</span><span class="sxs-lookup"><span data-stu-id="efc00-151">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="efc00-152">Eles atuam como margens das bordas da planilha, com [0, 0] sendo o canto superior esquerdo.</span><span class="sxs-lookup"><span data-stu-id="efc00-152">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="efc00-153">Eles podem ser definidos diretamente ou ajustados a partir da sua posição atual `incrementLeft` com `incrementTop` os métodos e.</span><span class="sxs-lookup"><span data-stu-id="efc00-153">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="efc00-154">O quanto uma forma é girada a partir da posição padrão também é estabelecida dessa maneira, com a `rotation` propriedade sendo o valor absoluto e o `incrementRotation` método ajustando a rotação existente.</span><span class="sxs-lookup"><span data-stu-id="efc00-154">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="efc00-155">A profundidade de uma forma em relação a outras formas é definida `zorderPosition` pela propriedade.</span><span class="sxs-lookup"><span data-stu-id="efc00-155">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="efc00-156">Isso é definido usando o `setZOrder` método, que usa um [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span><span class="sxs-lookup"><span data-stu-id="efc00-156">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="efc00-157">`setZOrder`ajusta a ordenação da forma atual em relação às outras formas.</span><span class="sxs-lookup"><span data-stu-id="efc00-157">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="efc00-158">O suplemento tem algumas opções para alterar a altura e a largura das formas.</span><span class="sxs-lookup"><span data-stu-id="efc00-158">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="efc00-159">A definição da `height` propriedade `width` or altera a dimensão especificada sem alterar a outra dimensão.</span><span class="sxs-lookup"><span data-stu-id="efc00-159">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="efc00-160">O `scaleHeight` e `scaleWidth` ajustam as respectivas dimensões da forma em relação ao tamanho atual ou original (com base no valor do [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)fornecido).</span><span class="sxs-lookup"><span data-stu-id="efc00-160">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="efc00-161">Um parâmetro [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) opcional especifica de onde a forma é dimensionada (canto superior esquerdo, médio ou canto inferior direito).</span><span class="sxs-lookup"><span data-stu-id="efc00-161">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="efc00-162">Se a `lockAspectRatio` propriedade for **true**, os métodos Scale mantêm a taxa de proporção atual da forma ajustando também a outra dimensão.</span><span class="sxs-lookup"><span data-stu-id="efc00-162">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="efc00-163">As alterações diretas `height` nas `width` Propriedades e só afetam essa propriedade, independentemente do `lockAspectRatio` valor da propriedade.</span><span class="sxs-lookup"><span data-stu-id="efc00-163">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="efc00-164">O exemplo de código a seguir mostra uma forma que está sendo dimensionada para 1,25 vezes seu tamanho original e girado 30 graus.</span><span class="sxs-lookup"><span data-stu-id="efc00-164">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="efc00-165">Texto em formas</span><span class="sxs-lookup"><span data-stu-id="efc00-165">Text in shapes</span></span>

<span data-ttu-id="efc00-166">Formas geoMétricas podem conter texto.</span><span class="sxs-lookup"><span data-stu-id="efc00-166">Geometric Shapes can contain text.</span></span> <span data-ttu-id="efc00-167">As formas têm `textFrame` uma propriedade do tipo TextFrame. [](/javascript/api/excel/excel.textframe)</span><span class="sxs-lookup"><span data-stu-id="efc00-167">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="efc00-168">O `TextFrame` objeto gerencia as opções de exibição de texto (como margens e estouro de texto).</span><span class="sxs-lookup"><span data-stu-id="efc00-168">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="efc00-169">`TextFrame.textRange`é um objeto [TextRange](/javascript/api/excel/excel.textrange) com o conteúdo de texto e as configurações de fonte.</span><span class="sxs-lookup"><span data-stu-id="efc00-169">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="efc00-170">O exemplo de código a seguir cria uma forma geométrica chamada "Wave" com o texto "Shape Text".</span><span class="sxs-lookup"><span data-stu-id="efc00-170">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="efc00-171">Também ajusta a forma e as cores do texto, bem como define o alinhamento horizontal do texto para o centro.</span><span class="sxs-lookup"><span data-stu-id="efc00-171">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="efc00-172">O `addTextBox` método de `ShapeCollection` criar um `GeometricShape` tipo `Rectangle` com um texto de fundo branco e preto.</span><span class="sxs-lookup"><span data-stu-id="efc00-172">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="efc00-173">É o mesmo que o que é criado pelo botão **caixa de texto** do Excel na guia **Inserir** . `addTextBox` Obtém um argumento de cadeia de caracteres para definir o `TextRange`texto do.</span><span class="sxs-lookup"><span data-stu-id="efc00-173">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="efc00-174">O exemplo de código a seguir mostra a criação de uma caixa de texto com o texto "Hello!".</span><span class="sxs-lookup"><span data-stu-id="efc00-174">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="efc00-175">Grupos de formas</span><span class="sxs-lookup"><span data-stu-id="efc00-175">Shape groups</span></span>

<span data-ttu-id="efc00-176">As formas podem ser agrupadas juntas.</span><span class="sxs-lookup"><span data-stu-id="efc00-176">Shapes can be grouped together.</span></span> <span data-ttu-id="efc00-177">Isso permite que um usuário o trate como uma única entidade de posicionamento, dimensionamento e outras tarefas relacionadas.</span><span class="sxs-lookup"><span data-stu-id="efc00-177">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="efc00-178">Um grupo de [formas](/javascript/api/excel/excel.shapegroup) é um tipo `Shape`de, para que o suplemento trate o grupo como uma única forma.</span><span class="sxs-lookup"><span data-stu-id="efc00-178">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="efc00-179">O exemplo de código a seguir mostra três formas que estão sendo agrupadas.</span><span class="sxs-lookup"><span data-stu-id="efc00-179">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="efc00-180">O exemplo de código subsequente mostra que o grupo de formas sendo movido para a direita 50 pixels.</span><span class="sxs-lookup"><span data-stu-id="efc00-180">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="efc00-181">As formas individuais dentro do grupo são referenciadas `ShapeGroup.shapes` por meio da propriedade, que é do tipo [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span><span class="sxs-lookup"><span data-stu-id="efc00-181">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="efc00-182">Eles não podem mais ser acessados por meio da coleção Shape da planilha após serem agrupados.</span><span class="sxs-lookup"><span data-stu-id="efc00-182">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="efc00-183">Por exemplo, se a sua planilha tivesse três formas e todos foram agrupadas, o método da `shapes.getCount` planilha retornará uma contagem de 1.</span><span class="sxs-lookup"><span data-stu-id="efc00-183">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="efc00-184">Exportar formas como imagens</span><span class="sxs-lookup"><span data-stu-id="efc00-184">Export shapes as images</span></span>

<span data-ttu-id="efc00-185">Qualquer `Shape` objeto pode ser convertido em uma imagem.</span><span class="sxs-lookup"><span data-stu-id="efc00-185">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="efc00-186">[Shape. getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) retorna Cadeia de caracteres codificada em base64.</span><span class="sxs-lookup"><span data-stu-id="efc00-186">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="efc00-187">O formato da imagem é especificado como um enum [PictureFormat](/javascript/api/excel/excel.pictureformat) passado para `getAsImage`.</span><span class="sxs-lookup"><span data-stu-id="efc00-187">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="efc00-188">Excluir formas</span><span class="sxs-lookup"><span data-stu-id="efc00-188">Delete shapes</span></span>

<span data-ttu-id="efc00-189">As formas são removidas da planilha `Shape` com o `delete` método do objeto.</span><span class="sxs-lookup"><span data-stu-id="efc00-189">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="efc00-190">Nenhum outro metadado é necessário.</span><span class="sxs-lookup"><span data-stu-id="efc00-190">No other metadata is needed.</span></span>

<span data-ttu-id="efc00-191">O exemplo de código a seguir exclui todas as \*\*\*\* formas de myworksheet.</span><span class="sxs-lookup"><span data-stu-id="efc00-191">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="efc00-192">Confira também</span><span class="sxs-lookup"><span data-stu-id="efc00-192">See also</span></span>

- [<span data-ttu-id="efc00-193">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="efc00-193">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="efc00-194">Trabalhar com gráficos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="efc00-194">Work with Charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
