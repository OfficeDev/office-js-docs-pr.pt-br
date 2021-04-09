---
title: Aplicar formatação condicional a intervalos com a API JavaScript do Excel
description: Saiba mais sobre formatação condicional no contexto dos complementos JavaScript do Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 5736923ac3c4623342d9a77bd58d5682a551a2a6
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650909"
---
# <a name="apply-conditional-formatting-to-excel-ranges"></a><span data-ttu-id="bed8d-103">Aplicar formatação condicional a intervalos do Excel</span><span class="sxs-lookup"><span data-stu-id="bed8d-103">Apply conditional formatting to Excel ranges</span></span>

<span data-ttu-id="bed8d-104">A Biblioteca de JavaScript do Excel fornece APIs para aplicar a formatação condicional aos intervalos de dados nas suas planilhas.</span><span class="sxs-lookup"><span data-stu-id="bed8d-104">The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets.</span></span> <span data-ttu-id="bed8d-105">Esse recurso simplifica a visualização da análise de grandes conjuntos de dados.</span><span class="sxs-lookup"><span data-stu-id="bed8d-105">This functionality makes large sets of data easy to visually parse.</span></span> <span data-ttu-id="bed8d-106">A formatação também atualiza dinamicamente com base nas alterações no intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-106">The formatting also dynamically updates based on changes within the range.</span></span>

> [!NOTE]
> <span data-ttu-id="bed8d-107">Este artigo aborda a formatação condicional no contexto dos suplementos do JavaScript do Excel. Os artigos a seguir fornecem informações detalhadas sobre os recursos completos de formatação condicionais do Excel.</span><span class="sxs-lookup"><span data-stu-id="bed8d-107">This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.</span></span>
> -  [<span data-ttu-id="bed8d-108">Adicionar, alterar ou limpar formatações condicionais</span><span class="sxs-lookup"><span data-stu-id="bed8d-108">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
> -  [<span data-ttu-id="bed8d-109">Use fórmulas com o acesso condicional</span><span class="sxs-lookup"><span data-stu-id="bed8d-109">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

## <a name="programmatic-control-of-conditional-formatting"></a><span data-ttu-id="bed8d-110">Controle de programação de formatação condicional</span><span class="sxs-lookup"><span data-stu-id="bed8d-110">Programmatic control of conditional formatting</span></span>

<span data-ttu-id="bed8d-111">A `Range.conditionalFormats` propriedade é uma coleção de objetos [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) que se aplicam ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-111">The `Range.conditionalFormats` property is a collection of [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) objects that apply to the range.</span></span>  <span data-ttu-id="bed8d-112">O `ConditionalFormat` objeto contém várias propriedades que definem o formato a ser aplicado com o [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="bed8d-112">The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span></span>

-    `cellValue`
-    `colorScale`
-    `custom`
-    `dataBar`
-    `iconSet`
-    `preset`
-    `textComparison`
-    `topBottom`

> [!NOTE]
> <span data-ttu-id="bed8d-113">Cada uma das seguintes propriedades de formatação tem uma variante `*OrNullObject` correspondente.</span><span class="sxs-lookup"><span data-stu-id="bed8d-113">Each of these formatting properties has a corresponding `*OrNullObject` variant.</span></span> <span data-ttu-id="bed8d-114">Saiba mais sobre esse padrão na [ \* seção Métodos OrNullObject.](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)</span><span class="sxs-lookup"><span data-stu-id="bed8d-114">Learn more about that pattern in the [\*OrNullObject methods](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) section.</span></span>

<span data-ttu-id="bed8d-115">Somente um tipo de formato pode ser definido para o objeto ConditionalFormat.</span><span class="sxs-lookup"><span data-stu-id="bed8d-115">Only one format type can be set for the ConditionalFormat object.</span></span> <span data-ttu-id="bed8d-116">Isso é determinado pela `type` propriedade, que é uma enumeração de valor[ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="bed8d-116">This is determined by the `type` property, which is a [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) enum value.</span></span> <span data-ttu-id="bed8d-117">`type` é definido quando um formato condicional para um intervalo é adicionado.</span><span class="sxs-lookup"><span data-stu-id="bed8d-117">`type` is set when adding a conditional format to a range.</span></span>

## <a name="creating-conditional-formatting-rules"></a><span data-ttu-id="bed8d-118">Criando regras de formatação condicional</span><span class="sxs-lookup"><span data-stu-id="bed8d-118">Creating conditional formatting rules</span></span>

<span data-ttu-id="bed8d-119">Formatos condicionais são adicionados a um intervalo usando `conditionalFormats.add`.</span><span class="sxs-lookup"><span data-stu-id="bed8d-119">Conditional formats are added to a range by using `conditionalFormats.add`.</span></span> <span data-ttu-id="bed8d-120">Após a adição, propriedades específicas podem ser definidas  para o formato condicional.</span><span class="sxs-lookup"><span data-stu-id="bed8d-120">Once added, the properties specific to the conditional format can be set.</span></span> <span data-ttu-id="bed8d-121">Os exemplos a seguir mostram a criação de diferentes tipos de formatação.</span><span class="sxs-lookup"><span data-stu-id="bed8d-121">The following examples show the creation of different formatting types.</span></span>

### <a name="cell-value"></a>[<span data-ttu-id="bed8d-122">Valor da célula</span><span class="sxs-lookup"><span data-stu-id="bed8d-122">Cell value</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)

<span data-ttu-id="bed8d-123">A formatação condicional de valor de célula aplica um formato definidas pelo usuário com base em uma ou duas fórmulas em [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span><span class="sxs-lookup"><span data-stu-id="bed8d-123">Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule).</span></span> <span data-ttu-id="bed8d-124">A `operator` propriedade é um[ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) que define como expressões resultantes se relacionam com a formatação.</span><span class="sxs-lookup"><span data-stu-id="bed8d-124">The `operator` property is a [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.</span></span>

<span data-ttu-id="bed8d-125">O exemplo a seguir mostra a cor de fonte vermelha aplicada a qualquer valor no intervalo menor que zero.</span><span class="sxs-lookup"><span data-stu-id="bed8d-125">The following example shows red font coloring applied to any value in the range less than zero.</span></span>

![Um intervalo com números negativos em vermelho.](../images/excel-conditional-format-cell-value.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
);

// set the font of negative numbers to red
conditionalFormat.cellValue.format.font.color = "red";
conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };

await context.sync();
```

### <a name="color-scale"></a>[<span data-ttu-id="bed8d-127">Escala de cores</span><span class="sxs-lookup"><span data-stu-id="bed8d-127">Color scale</span></span>](/javascript/api/excel/excel.colorscaleconditionalformat)

<span data-ttu-id="bed8d-128">Formatação condicional de escala de cores aplica um gradiente de cor para o intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="bed8d-128">Color scale conditional formatting applies a color gradient across the data range.</span></span> <span data-ttu-id="bed8d-129">A `criteria` propriedade na `ColorScaleConditionalFormat` define três [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`e, opcionalmente, `midpoint`.</span><span class="sxs-lookup"><span data-stu-id="bed8d-129">The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`.</span></span> <span data-ttu-id="bed8d-130">Cada um dos pontos de escala critério têm três propriedades:</span><span class="sxs-lookup"><span data-stu-id="bed8d-130">Each of the criterion scale points have three properties:</span></span>

-    <span data-ttu-id="bed8d-131">`color` – O código de cor HTML para o ponto de extremidade.</span><span class="sxs-lookup"><span data-stu-id="bed8d-131">`color` - The HTML color code for the endpoint.</span></span>
-    <span data-ttu-id="bed8d-132">`formula` – Um número ou uma fórmula que representa o ponto de extremidade.</span><span class="sxs-lookup"><span data-stu-id="bed8d-132">`formula` - A number or formula representing the endpoint.</span></span> <span data-ttu-id="bed8d-133">Isso será `null` caso `type` está `lowestValue` ou `highestValue`.</span><span class="sxs-lookup"><span data-stu-id="bed8d-133">This will be `null` if `type` is `lowestValue` or `highestValue`.</span></span>
-    <span data-ttu-id="bed8d-134">`type` Como a fórmula deve ser avaliada.</span><span class="sxs-lookup"><span data-stu-id="bed8d-134">`type` - How the formula should be evaluated.</span></span> <span data-ttu-id="bed8d-135">`highestValue` e `lowestValue` fazem referência a valores no intervalo a ser formatado.</span><span class="sxs-lookup"><span data-stu-id="bed8d-135">`highestValue` and `lowestValue` refer to values in the range being formatted.</span></span>

<span data-ttu-id="bed8d-136">O exemplo a seguir mostra um intervalo a ser colorido de azul para amarelo para vermelho.</span><span class="sxs-lookup"><span data-stu-id="bed8d-136">The following example shows a range being colored blue to yellow to red.</span></span> <span data-ttu-id="bed8d-137">Observe que `minimum` e `maximum` são os valores mais altos e mais baixos, respectivamente e usam `null` fórmulas.</span><span class="sxs-lookup"><span data-stu-id="bed8d-137">Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas.</span></span> <span data-ttu-id="bed8d-138">`midpoint` está usando o `percentage` tipo com uma fórmula de `"=50"` então a célula yellowest é o valor médio.</span><span class="sxs-lookup"><span data-stu-id="bed8d-138">`midpoint` is using the `percentage` type with a formula of `"=50"` so the yellowest cell is the mean value.</span></span>

![Um intervalo com o número baixo em azul, o número médio em amarelo e o número alto é vermelho, com gradientes entre valores.](../images/excel-conditional-format-color-scale.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
      Excel.ConditionalFormatType.colorScale
);

// color the backgrounds of the cells from blue to yellow to red based on value
const criteria = {
      minimum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.lowestValue,
           color: "blue"
      },
      midpoint: {
           formula: "50",
           type: Excel.ConditionalFormatColorCriterionType.percent,
           color: "yellow"
      },
      maximum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.highestValue,
           color: "red"
      }
};
conditionalFormat.colorScale.criteria = criteria;

await context.sync();
```

### <a name="custom"></a>[<span data-ttu-id="bed8d-140">Personalizados</span><span class="sxs-lookup"><span data-stu-id="bed8d-140">Custom</span></span>](/javascript/api/excel/excel.customconditionalformat)

<span data-ttu-id="bed8d-141">A formatação condicional personalizada aplica um formato definido pelo usuário para as células com base em uma fórmula de complexidade arbitrária.</span><span class="sxs-lookup"><span data-stu-id="bed8d-141">Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity.</span></span> <span data-ttu-id="bed8d-142">O objeto [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) permite que você defina a fórmula em notações diferentes:</span><span class="sxs-lookup"><span data-stu-id="bed8d-142">The [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:</span></span>

-    <span data-ttu-id="bed8d-143">`formula` -Anotação padrão.</span><span class="sxs-lookup"><span data-stu-id="bed8d-143">`formula` - Standard notation.</span></span>
-    <span data-ttu-id="bed8d-144">`formulaLocal` - Localizado com base no idioma do usuário.</span><span class="sxs-lookup"><span data-stu-id="bed8d-144">`formulaLocal` - Localized based on the user's language.</span></span>
-    <span data-ttu-id="bed8d-145">`formulaR1C1` -Notação estilo R1C1.</span><span class="sxs-lookup"><span data-stu-id="bed8d-145">`formulaR1C1` - R1C1-style notation.</span></span>

<span data-ttu-id="bed8d-146">O exemplo de cores a seguir as fontes de verde nas células com valores maiores que a célula à esquerda.</span><span class="sxs-lookup"><span data-stu-id="bed8d-146">The following example colors the fonts green of cells with higher values than the cell to their left.</span></span>

![Um intervalo com números verdes para locais em que o valor da coluna anterior nessa linha é inferior.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell's font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### <a name="data-bar"></a>[<span data-ttu-id="bed8d-148">Barra de dados</span><span class="sxs-lookup"><span data-stu-id="bed8d-148">Data bar</span></span>](/javascript/api/excel/excel.databarconditionalformat)

<span data-ttu-id="bed8d-149">A barra de formatação condicional de dados adiciona barras de dados nas células.</span><span class="sxs-lookup"><span data-stu-id="bed8d-149">Data bar conditional formatting adds data bars to the cells.</span></span> <span data-ttu-id="bed8d-150">Por padrão, os valores mínimos e máximos no intervalo formam limites e tamanhos proporcionais às barras de dados.</span><span class="sxs-lookup"><span data-stu-id="bed8d-150">By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars.</span></span> <span data-ttu-id="bed8d-151">O `DataBarConditionalFormat` objeto tem várias propriedades para controlar a aparência da barra.</span><span class="sxs-lookup"><span data-stu-id="bed8d-151">The `DataBarConditionalFormat` object has several properties to control the bar's appearance.</span></span> 

<span data-ttu-id="bed8d-152">O exemplo a seguir formata o intervalo com barras de dados preenchidas da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="bed8d-152">The following example formats the range with data bars filling left-to-right.</span></span>

![Um intervalo com databars atrás dos valores nas células.](../images/excel-conditional-format-databar.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.dataBar
);

// give left-to-right, default-appearance data bars to all the cells
conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
await context.sync();
```

### <a name="icon-set"></a>[<span data-ttu-id="bed8d-154">Conjunto de ícones</span><span class="sxs-lookup"><span data-stu-id="bed8d-154">Icon set</span></span>](/javascript/api/excel/excel.iconsetconditionalformat)

<span data-ttu-id="bed8d-155">A formatação condicional do conjunto de ícones usa os [ícones](/javascript/api/excel/excel.icon) do Excel para realçar células.</span><span class="sxs-lookup"><span data-stu-id="bed8d-155">Icon set conditional formatting uses Excel [Icons](/javascript/api/excel/excel.icon) to highlight cells.</span></span> <span data-ttu-id="bed8d-156">A `criteria` propriedade é uma matriz [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), que define o símbolo a ser inserido e a condição em que ele é inserido.</span><span class="sxs-lookup"><span data-stu-id="bed8d-156">The `criteria` property is an array of [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), which define the symbol to be inserted and the condition under which it is inserted.</span></span> <span data-ttu-id="bed8d-157">Essa matriz é automaticamente pré-preenchida com critério de elementos com propriedades padrão.</span><span class="sxs-lookup"><span data-stu-id="bed8d-157">This array is automatically prepopulated with criterion elements with default properties.</span></span> <span data-ttu-id="bed8d-158">Propriedades individuais não podem ser substituídas.</span><span class="sxs-lookup"><span data-stu-id="bed8d-158">Individual properties cannot be overwritten.</span></span> <span data-ttu-id="bed8d-159">Em vez disso, todo o objeto de critérios deve ser substituído.</span><span class="sxs-lookup"><span data-stu-id="bed8d-159">Instead, the whole criteria object must be replaced.</span></span> 

<span data-ttu-id="bed8d-160">O exemplo a seguir mostra um conjunto de ícones de três triângulos aplicado ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-160">The following example shows a three-triangle icon set applied across the range.</span></span>

![Um intervalo com triângulos verdes para cima para valores acima de 1000, linhas amarelas para valores entre 700 e 1000 e triângulos vermelhos para baixo para valores inferiores.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*" icon set style, such as "threeTriangles", the third
    element in the criteria array (criteria[2]) defines the "top" icon;
    e.g., a green triangle. The second (criteria[1]) defines the "middle"
    icon, The first (criteria[0]) defines the "low" icon, but it can often 
    be left empty as this method does below, because every cell that
   does not match the other two criteria always gets the low icon.
*/
iconSetCF.criteria = [
    {} as any,
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=700"
      },
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=1000"
      }
];

await context.sync();
```

### <a name="preset-criteria"></a>[<span data-ttu-id="bed8d-162">Critérios predefinidos</span><span class="sxs-lookup"><span data-stu-id="bed8d-162">Preset criteria</span></span>](/javascript/api/excel/excel.presetcriteriaconditionalformat)

<span data-ttu-id="bed8d-163">A formatação condicional predefinida aplica um formato definido pelo usuário ao intervalo com base em uma regra padrão selecionada.</span><span class="sxs-lookup"><span data-stu-id="bed8d-163">Preset conditional formatting applies a user-defined format to the range based on a selected standard rule.</span></span> <span data-ttu-id="bed8d-164">Essas regras são definidas pelo[ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) no [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span><span class="sxs-lookup"><span data-stu-id="bed8d-164">These rules are defined by the [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) in the [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).</span></span> 

<span data-ttu-id="bed8d-165">O exemplo a seguir colore a fonte branca onde quer que o valor de uma célula seja pelo menos um desvio padrão acima da média do intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-165">The following example colors the font white wherever a cell's value is at least one standard deviation above the range's average.</span></span>

![Um intervalo com células de fonte branca onde os valores tem pelo menos um desvio padrão acima da média.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell's font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### <a name="text-comparison"></a>[<span data-ttu-id="bed8d-167">Comparação de texto</span><span class="sxs-lookup"><span data-stu-id="bed8d-167">Text comparison</span></span>](/javascript/api/excel/excel.textconditionalformat)

<span data-ttu-id="bed8d-168">A formatação condicional de texto comparação usa comparações de cadeias como condição.</span><span class="sxs-lookup"><span data-stu-id="bed8d-168">Text comparison conditional formatting uses string comparisons as the condition.</span></span> <span data-ttu-id="bed8d-169">As`rule` propriedade é [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) definindo uma cadeia de caracteres a ser comparada com a célula e um operador para especificar o tipo de comparação.</span><span class="sxs-lookup"><span data-stu-id="bed8d-169">The `rule` property is a [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison.</span></span> 

<span data-ttu-id="bed8d-170">O exemplo a seguir formata a cor da fonte vermelha quando o texto de uma célula contém "Atrasado".</span><span class="sxs-lookup"><span data-stu-id="bed8d-170">The following example formats the font color red when a cell's text contains "Delayed".</span></span>

![Um intervalo com células que contêm "Atrasado" em vermelho.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing "Delayed"
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### <a name="topbottom"></a>[<span data-ttu-id="bed8d-172">Superiores/inferiores</span><span class="sxs-lookup"><span data-stu-id="bed8d-172">Top/bottom</span></span>](/javascript/api/excel/excel.TopBottomconditionalformat)

<span data-ttu-id="bed8d-173">A formatação condicional superiores/inferiores aplica um formato para maiores ou menores valores em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-173">Top/bottom conditional formatting applies a format to the highest or lowest values in a range.</span></span> <span data-ttu-id="bed8d-174">As `rule` propriedade é do tipo [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), define a condição se baseia no maior ou menor, e se a avaliação é ordenada ou na baseada na porcentagem.</span><span class="sxs-lookup"><span data-stu-id="bed8d-174">The `rule` property, which is of type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based.</span></span> 

<span data-ttu-id="bed8d-175">O exemplo a seguir aplica um destaque em verde na maior célula valor do intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-175">The following example applies a green highlight to the highest value cell in the range.</span></span>


![Um intervalo com o maior número realçado em verde.](../images/excel-conditional-format-topbottom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.topBottom
);

// for the highest valued cell in the range, make the background green
conditionalFormat.topBottom.format.fill.color = "green"
conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems"}

await context.sync();
```

## <a name="multiple-formats-and-priority"></a><span data-ttu-id="bed8d-177">Vários formatos e prioridades</span><span class="sxs-lookup"><span data-stu-id="bed8d-177">Multiple formats and priority</span></span>

<span data-ttu-id="bed8d-178">Você pode aplicar vários formatos condicionais em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-178">You can apply multiple conditional formats to a range.</span></span> <span data-ttu-id="bed8d-179">Se os formatos tem elementos conflitantes, como cores de fonte diferentes apenas um formato aplica-se a esse elemento determinado.</span><span class="sxs-lookup"><span data-stu-id="bed8d-179">If the formats have conflicting elements, such as differing font colors, only one format applies that particular element.</span></span> <span data-ttu-id="bed8d-180">Precedência é definida pela propriedade `ConditionalFormat.priority`.</span><span class="sxs-lookup"><span data-stu-id="bed8d-180">Precedence is defined by the `ConditionalFormat.priority` property.</span></span> <span data-ttu-id="bed8d-181">Prioridade é um número (igual ao índice a `ConditionalFormatCollection`) e pode ser definido ao criar o formato.</span><span class="sxs-lookup"><span data-stu-id="bed8d-181">Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format.</span></span> <span data-ttu-id="bed8d-182">Quanto mais baixo o `priority` valor for, maior a prioridade do formato é.</span><span class="sxs-lookup"><span data-stu-id="bed8d-182">The lowerer the `priority` value, the higher the priority of the format is.</span></span>

<span data-ttu-id="bed8d-183">O exemplo a seguir mostra uma opção de cor da fonte conflitante entre os dois formatos.</span><span class="sxs-lookup"><span data-stu-id="bed8d-183">The following example shows a conflicting font color choice between the two formats.</span></span> <span data-ttu-id="bed8d-184">Números negativos receberão uma fonte em negrito, mas não a fonte vermelha, porque a prioridade é o formato que oferece uma fonte azul.</span><span class="sxs-lookup"><span data-stu-id="bed8d-184">Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.</span></span>

![Um intervalo com números menores em negrito e em vermelhos, números negativos em azul com telas de fundo verdes.](../images/excel-conditional-format-priority.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();


// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and set priority 0.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;

await context.sync();

```

### <a name="mutually-exclusive-conditional-formats"></a><span data-ttu-id="bed8d-186">Formatos condicionais mutuamente exclusivos </span><span class="sxs-lookup"><span data-stu-id="bed8d-186">Mutually exclusive conditional formats</span></span>

<span data-ttu-id="bed8d-187">As `stopIfTrue` propriedade de `ConditionalFormat` impede que os formatos condicionais de prioridade inferiores sejam aplicados ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-187">The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range.</span></span> <span data-ttu-id="bed8d-188">Quando um intervalo que corresponde ao formato condicional com `stopIfTrue === true` é aplicado, nenhum formato condicional subsequente é aplicado, mesmo se os detalhes da formatação não forem contraditórios.</span><span class="sxs-lookup"><span data-stu-id="bed8d-188">When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.</span></span>

<span data-ttu-id="bed8d-189">O exemplo a seguir mostra dois formatos condicionais adicionados a um intervalo.</span><span class="sxs-lookup"><span data-stu-id="bed8d-189">The following example shows two conditional formats being added to a range.</span></span> <span data-ttu-id="bed8d-190">Números negativos terão uma fonte azul com um fundo verde suave, independentemente da condição de formatação ser verdadeira.</span><span class="sxs-lookup"><span data-stu-id="bed8d-190">Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.</span></span>

![Um intervalo com números baixos em negrito e em vermelho, a menos que sejam negativos; nesse caso, eles não estão em negrito, em azul e têm um plano de fundo verde.](../images/excel-conditional-format-stopiftrue.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();

// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and 
// set priority 0, but set stopIfTrue to true, so none of the 
// formatting of the conditional format with the higher priority
// value will apply, not even the bolding of the font.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;
cellValueFormat.stopIfTrue = true;

await context.sync();
```

## <a name="see-also"></a><span data-ttu-id="bed8d-192">Confira também</span><span class="sxs-lookup"><span data-stu-id="bed8d-192">See also</span></span>

- [<span data-ttu-id="bed8d-193">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bed8d-193">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bed8d-194">Objeto ConditionalFormat (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="bed8d-194">ConditionalFormat Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.conditionalformat)
- [<span data-ttu-id="bed8d-195">Adicionar, alterar ou limpar formatações condicionais</span><span class="sxs-lookup"><span data-stu-id="bed8d-195">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
- [<span data-ttu-id="bed8d-196">Use fórmulas com o acesso condicional</span><span class="sxs-lookup"><span data-stu-id="bed8d-196">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)
