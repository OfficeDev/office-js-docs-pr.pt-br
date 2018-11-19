# <a name="apply-conditional-formatting-to-excel-ranges"></a><span data-ttu-id="523c4-101">Aplicar formatação condicional de intervalos do Excel</span><span class="sxs-lookup"><span data-stu-id="523c4-101">Apply conditional formatting to Excel ranges</span></span>

<span data-ttu-id="523c4-102">A Biblioteca de JavaScript do Excel fornece APIs para aplicar a formatação condicional aos intervalos de dados nas suas planilhas.</span><span class="sxs-lookup"><span data-stu-id="523c4-102">The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets.</span></span> <span data-ttu-id="523c4-103">Esse recurso simplifica a visualização da análise de grandes conjuntos de dados.</span><span class="sxs-lookup"><span data-stu-id="523c4-103">This functionality makes large sets of data easy to visually parse.</span></span> <span data-ttu-id="523c4-104">A formatação também atualiza dinamicamente com base nas alterações no intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-104">The formatting also dynamically updates based on changes within the range.</span></span> 

> [!NOTE] 
> <span data-ttu-id="523c4-105">Este artigo aborda a formatação condicional no contexto dos suplementos do JavaScript do Excel. Os artigos a seguir fornecem informações detalhadas sobre os recursos completos de formatação condicionais do Excel.</span><span class="sxs-lookup"><span data-stu-id="523c4-105">This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.</span></span>
-   [<span data-ttu-id="523c4-106">Adicionar, alterar ou limpar formatações condicionais</span><span class="sxs-lookup"><span data-stu-id="523c4-106">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
-   [<span data-ttu-id="523c4-107">Use fórmulas com o acesso condicional</span><span class="sxs-lookup"><span data-stu-id="523c4-107">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

## <a name="programmatic-control-of-conditional-formatting"></a><span data-ttu-id="523c4-108">Controle de programação de formatação condicional</span><span class="sxs-lookup"><span data-stu-id="523c4-108">Programmatic control of conditional formatting</span></span>

<span data-ttu-id="523c4-109">A `Range.conditionalFormats` propriedade é uma coleção de objetos [ConditionalFormat](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformat) que se aplicam ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-109">The `Range.conditionalFormats` property is a collection of [ConditionalFormat](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformat) objects that apply to the range.</span></span>  <span data-ttu-id="523c4-110">O `ConditionalFormat` objeto contém várias propriedades que definem o formato a ser aplicado com o [ConditionalFormatType](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="523c4-110">The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformattype).</span></span> 

-   `cellValue`
-   `colorScale`
-   `custom`
-   `dataBar`
-   `iconSet`
-   `preset`
-   `textComparison`
-   `topBottom`

> [!NOTE]
> <span data-ttu-id="523c4-111">Cada uma das seguintes propriedades de formatação tem uma variante `*OrNullObject` correspondente.</span><span class="sxs-lookup"><span data-stu-id="523c4-111">Each of these formatting properties has a corresponding `*OrNullObject` variant.</span></span> <span data-ttu-id="523c4-112">Saiba mais sobre esse padrão na seção [\* OrNullObject métodos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="523c4-112">Learn more about that pattern in the [\*OrNullObject methods](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) section.</span></span>

<span data-ttu-id="523c4-113">Somente um tipo de formato pode ser definido para o objeto ConditionalFormat.</span><span class="sxs-lookup"><span data-stu-id="523c4-113">Only one format type can be set for the ConditionalFormat object.</span></span> <span data-ttu-id="523c4-114">Isso é determinado pela `type` propriedade, que é uma enumeração de valor[ConditionalFormatType](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformattype).</span><span class="sxs-lookup"><span data-stu-id="523c4-114">This is determined by the `type` property, which is a [ConditionalFormatType](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformattype) enum value.</span></span> <span data-ttu-id="523c4-115">`type` é definido quando um formato condicional para um intervalo é adicionado.</span><span class="sxs-lookup"><span data-stu-id="523c4-115">`type` is set when adding a conditional format to a range.</span></span> 

## <a name="creating-conditional-formatting-rules"></a><span data-ttu-id="523c4-116">Criando regras de formatação condicional</span><span class="sxs-lookup"><span data-stu-id="523c4-116">Creating conditional formatting rules</span></span>

<span data-ttu-id="523c4-117">Formatos condicionais são adicionados a um intervalo usando `conditionalFormats.add`.</span><span class="sxs-lookup"><span data-stu-id="523c4-117">Conditional formats are added to a range by using `conditionalFormats.add`.</span></span> <span data-ttu-id="523c4-118">Após a adição, propriedades específicas podem ser definidas  para o formato condicional.</span><span class="sxs-lookup"><span data-stu-id="523c4-118">Once added, the properties specific to the conditional format can be set.</span></span> <span data-ttu-id="523c4-119">Os exemplos a seguir mostram a criação de diferentes tipos de formatação.</span><span class="sxs-lookup"><span data-stu-id="523c4-119">The following examples show the creation of different formatting types.</span></span>

### <a name="cell-valuehttpsdocsmicrosoftcomjavascriptapiexcelexcelcellvalueconditionalformat"></a>[<span data-ttu-id="523c4-120">Valor da célula</span><span class="sxs-lookup"><span data-stu-id="523c4-120">Cell value</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.cellvalueconditionalformat)

<span data-ttu-id="523c4-121">A formatação condicional de valor de célula aplica um formato definidas pelo usuário com base em uma ou duas fórmulas em [ConditionalCellValueRule]( https://docs.microsoft.com/javascript/api/excel/excel.conditionalcellvaluerule).</span><span class="sxs-lookup"><span data-stu-id="523c4-121">Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule]( https://docs.microsoft.com/javascript/api/excel/excel.conditionalcellvaluerule).</span></span> <span data-ttu-id="523c4-122">A `operator` propriedade é um[ConditionalCellValueOperator]( https://docs.microsoft.com/javascript/api/excel/excel.conditionalcellvalueoperator) que define como expressões resultantes se relacionam com a formatação.</span><span class="sxs-lookup"><span data-stu-id="523c4-122">The `operator` property is a [ConditionalCellValueOperator]( https://docs.microsoft.com/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.</span></span> 

<span data-ttu-id="523c4-123">O exemplo a seguir mostra a cor de fonte vermelha aplicada a qualquer valor no intervalo menor que zero.</span><span class="sxs-lookup"><span data-stu-id="523c4-123">The following example shows red font coloring applied to any value in the range less than zero.</span></span>

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

### <a name="color-scalehttpsdocsmicrosoftcomjavascriptapiexcelexcelcolorscaleconditionalformat"></a>[<span data-ttu-id="523c4-125">Escala de cores</span><span class="sxs-lookup"><span data-stu-id="523c4-125">Color scale</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.colorscaleconditionalformat)

<span data-ttu-id="523c4-126">Formatação condicional de escala de cores aplica um gradiente de cor para o intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="523c4-126">Color scale conditional formatting applies a color gradient across the data range.</span></span> <span data-ttu-id="523c4-127">A `criteria` propriedade na `ColorScaleConditionalFormat` define três [ConditionalColorScaleCriterion](https://docs.microsoft.com/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`e, opcionalmente, `midpoint`.</span><span class="sxs-lookup"><span data-stu-id="523c4-127">The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](https://docs.microsoft.com/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`.</span></span> <span data-ttu-id="523c4-128">Cada um dos pontos de escala critério têm três propriedades:</span><span class="sxs-lookup"><span data-stu-id="523c4-128">Each of the criterion scale points have three properties:</span></span>

-   <span data-ttu-id="523c4-129">`color` – O código de cor HTML para o ponto de extremidade.</span><span class="sxs-lookup"><span data-stu-id="523c4-129">`color` - The HTML color code for the endpoint.</span></span>
-   <span data-ttu-id="523c4-130">`formula` – Um número ou uma fórmula que representa o ponto de extremidade.</span><span class="sxs-lookup"><span data-stu-id="523c4-130">`formula` - A number or formula representing the endpoint.</span></span> <span data-ttu-id="523c4-131">Isso será `null` caso `type` está `lowestValue` ou `highestValue`.</span><span class="sxs-lookup"><span data-stu-id="523c4-131">This will be `null` if `type` is `lowestValue` or `highestValue`.</span></span>
-   <span data-ttu-id="523c4-132">`type` Como a fórmula deve ser avaliada.</span><span class="sxs-lookup"><span data-stu-id="523c4-132">`type` - How the formula should be evaluated.</span></span> <span data-ttu-id="523c4-133">`highestValue` e `lowestValue` fazem referência a valores no intervalo a ser formatado.</span><span class="sxs-lookup"><span data-stu-id="523c4-133">`highestValue` and `lowestValue` refer to values in the range being formatted.</span></span>

<span data-ttu-id="523c4-134">O exemplo a seguir mostra um intervalo a ser colorido de azul para amarelo para vermelho.</span><span class="sxs-lookup"><span data-stu-id="523c4-134">The following example shows a range being colored blue to yellow to red.</span></span> <span data-ttu-id="523c4-135">Observe que `minimum` e `maximum` são os valores mais altos e mais baixos, respectivamente e usam `null` fórmulas.</span><span class="sxs-lookup"><span data-stu-id="523c4-135">Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas.</span></span> <span data-ttu-id="523c4-136">`midpoint` está usando o `percentage` tipo com uma fórmula de `”=50”` então a célula yellowest é o valor médio.</span><span class="sxs-lookup"><span data-stu-id="523c4-136">`midpoint` is using the `percentage` type with a formula of `”=50”` so the yellowest cell is the mean value.</span></span>

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

### <a name="customhttpsdocsmicrosoftcomjavascriptapiexcelexcelcustomconditionalformat"></a>[<span data-ttu-id="523c4-138">Personalizados</span><span class="sxs-lookup"><span data-stu-id="523c4-138">Custom</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.customconditionalformat) 

<span data-ttu-id="523c4-139">A formatação condicional personalizada aplica um formato definido pelo usuário para as células com base em uma fórmula de complexidade arbitrária.</span><span class="sxs-lookup"><span data-stu-id="523c4-139">Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity.</span></span> <span data-ttu-id="523c4-140">O objeto [ConditionalFormatRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformatrule) permite que você defina a fórmula em notações diferentes:</span><span class="sxs-lookup"><span data-stu-id="523c4-140">The [ConditionalFormatRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:</span></span>

-   <span data-ttu-id="523c4-141">`formula` -Anotação padrão.</span><span class="sxs-lookup"><span data-stu-id="523c4-141">`formula` - Standard notation.</span></span> 
-   <span data-ttu-id="523c4-142">`formulaLocal` – Localizados com base no idioma do usuário.</span><span class="sxs-lookup"><span data-stu-id="523c4-142">`formulaLocal` - Localized based on the user’s language.</span></span>
-   <span data-ttu-id="523c4-143">`formulaR1C1` -Notação estilo R1C1.</span><span class="sxs-lookup"><span data-stu-id="523c4-143">`formulaR1C1` - R1C1-style notation.</span></span>

<span data-ttu-id="523c4-144">O exemplo de cores a seguir as fontes de verde nas células com valores maiores que a célula à esquerda.</span><span class="sxs-lookup"><span data-stu-id="523c4-144">The following example colors the fonts green of cells with higher values than the cell to their left.</span></span>

![Um intervalo com números verdes para locais em que o valor da coluna anterior nessa linha é inferior.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell’s font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### <a name="data-barhttpsdocsmicrosoftcomjavascriptapiexcelexceldatabarconditionalformat"></a>[<span data-ttu-id="523c4-146">Barra de dados</span><span class="sxs-lookup"><span data-stu-id="523c4-146">Data bar</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.databarconditionalformat)

<span data-ttu-id="523c4-147">A barra de formatação condicional de dados adiciona barras de dados nas células.</span><span class="sxs-lookup"><span data-stu-id="523c4-147">Data bar conditional formatting adds data bars to the cells.</span></span> <span data-ttu-id="523c4-148">Por padrão, os valores mínimos e máximos no intervalo formam limites e tamanhos proporcionais às barras de dados.</span><span class="sxs-lookup"><span data-stu-id="523c4-148">By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars.</span></span> <span data-ttu-id="523c4-149">O `DataBarConditionalFormat` objeto tem várias propriedades de controle da aparência da barra.</span><span class="sxs-lookup"><span data-stu-id="523c4-149">The `DataBarConditionalFormat` object has several properties to control the bar’s appearance.</span></span> 

<span data-ttu-id="523c4-150">O exemplo a seguir formata o intervalo com barras de dados preenchidas da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="523c4-150">The following example formats the range with data bars filling left-to-right.</span></span>

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

### <a name="icon-sethttpsdocsmicrosoftcomjavascriptapiexcelexceliconsetconditionalformat"></a>[<span data-ttu-id="523c4-152">Conjunto de ícones</span><span class="sxs-lookup"><span data-stu-id="523c4-152">Icon set</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.iconsetconditionalformat)

<span data-ttu-id="523c4-153">A formatação condicional do conjunto de ícones usa os [ícones]( https://docs.microsoft.com/javascript/api/excel/excel.icon) do Excel para realçar células.</span><span class="sxs-lookup"><span data-stu-id="523c4-153">Icon set conditional formatting uses Excel [Icons]( https://docs.microsoft.com/javascript/api/excel/excel.icon) to highlight cells.</span></span> <span data-ttu-id="523c4-154">A `criteria` propriedade é uma matriz [ConditionalIconCriterion](https://docs.microsoft.com/javascript/api/excel/excel.ConditionalIconCriterion), que define o símbolo a ser inserido e a condição em que ele é inserido.</span><span class="sxs-lookup"><span data-stu-id="523c4-154">The `criteria` property is an array of [ConditionalIconCriterion](https://docs.microsoft.com/javascript/api/excel/excel.ConditionalIconCriterion), which define the symbol to be inserted and the condition under which it is inserted.</span></span> <span data-ttu-id="523c4-155">Essa matriz é automaticamente pré-preenchida com critério de elementos com propriedades padrão.</span><span class="sxs-lookup"><span data-stu-id="523c4-155">This array is automatically prepopulated with criterion elements with default properties.</span></span> <span data-ttu-id="523c4-156">Propriedades individuais não podem ser substituídas.</span><span class="sxs-lookup"><span data-stu-id="523c4-156">Individual properties cannot be overwritten.</span></span> <span data-ttu-id="523c4-157">Em vez disso, todo o objeto de critérios deve ser substituído.</span><span class="sxs-lookup"><span data-stu-id="523c4-157">Instead, the whole criteria object must be replaced.</span></span> 

<span data-ttu-id="523c4-158">O exemplo a seguir mostra um conjunto de ícones de três triângulos aplicado ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-158">The following example shows a three-triangle icon set applied across the range.</span></span>

![Um intervalo com triângulos verdes para cima para valores acima de 1000 linhas amarelas para values entre 1000 e 700 e triângulos vermelhos para baixo para valores mais baixos.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*” icon set style, such as "threeTriangles", the third
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

### <a name="preset-criteriahttpsdocsmicrosoftcomjavascriptapiexcelexcelpresetcriteriaconditionalformat"></a>[<span data-ttu-id="523c4-160">Critérios predefinidos</span><span class="sxs-lookup"><span data-stu-id="523c4-160">Preset criteria</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.presetcriteriaconditionalformat)

<span data-ttu-id="523c4-161">A formatação condicional predefinida aplica um formato definido pelo usuário ao intervalo com base em uma regra padrão selecionada.</span><span class="sxs-lookup"><span data-stu-id="523c4-161">Preset conditional formatting applies a user-defined format to the range based on a selected standard rule.</span></span> <span data-ttu-id="523c4-162">Essas regras são definidas pelo[ConditionalFormatPresetCriterion](https://docs.microsoft.com/javascript/api/excel/excel.ConditionalFormatPresetCriterion) no [ConditionalPresetCriteriaRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionalpresetcriteriarule).</span><span class="sxs-lookup"><span data-stu-id="523c4-162">These rules are defined by the [ConditionalFormatPresetCriterion](https://docs.microsoft.com/javascript/api/excel/excel.ConditionalFormatPresetCriterion) in the [ConditionalPresetCriteriaRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionalpresetcriteriarule).</span></span> 

<span data-ttu-id="523c4-163">O exemplo a seguir cor da fonte é branca onde o valor da célula tem pelo menos um desvio padrão da acima do intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-163">The following example colors the font white wherever a cell’s value is at least one standard deviation above the range’s average.</span></span>

![Um intervalo com células de fonte branca onde os valores tem pelo menos um desvio padrão acima da média.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell’s font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### <a name="text-comparisonhttpsdocsmicrosoftcomjavascriptapiexcelexceltextconditionalformat"></a>[<span data-ttu-id="523c4-165">Comparação de texto</span><span class="sxs-lookup"><span data-stu-id="523c4-165">Text comparison</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.textconditionalformat)

<span data-ttu-id="523c4-166">A formatação condicional de texto comparação usa comparações de cadeias como condição.</span><span class="sxs-lookup"><span data-stu-id="523c4-166">Text comparison conditional formatting uses string comparisons as the condition.</span></span> <span data-ttu-id="523c4-167">As`rule` propriedade é [ConditionalTextComparisonRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionaltextcomparisonrule) definindo uma cadeia de caracteres a ser comparada com a célula e um operador para especificar o tipo de comparação.</span><span class="sxs-lookup"><span data-stu-id="523c4-167">The `rule` property is a [ConditionalTextComparisonRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison.</span></span> 

<span data-ttu-id="523c4-168">O exemplo a seguir mostra a cor da fonte vermelha quando o texto de uma célula contém "Atrasada".</span><span class="sxs-lookup"><span data-stu-id="523c4-168">The following example formats the font color red when a cell’s text contains “Delayed”.</span></span>

![Um intervalo com células que contêm "Atrasado" em vermelho.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing “Delayed”
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### <a name="topbottomhttpsdocsmicrosoftcomjavascriptapiexcelexceltopbottomconditionalformat"></a>[<span data-ttu-id="523c4-170">Superiores/inferiores</span><span class="sxs-lookup"><span data-stu-id="523c4-170">TopBottom</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.TopBottomconditionalformat)

<span data-ttu-id="523c4-171">A formatação condicional superiores/inferiores aplica um formato para maiores ou menores valores em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-171">Top/bottom conditional formatting applies a format to the highest or lowest values in a range.</span></span> <span data-ttu-id="523c4-172">As `rule` propriedade é do tipo [ConditionalTopBottomRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionaltopbottomrule), define a condição se baseia no maior ou menor, e se a avaliação é ordenada ou na baseada na porcentagem.</span><span class="sxs-lookup"><span data-stu-id="523c4-172">The `rule` property, which is of type [ConditionalTopBottomRule](https://docs.microsoft.com/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based.</span></span> 

<span data-ttu-id="523c4-173">O exemplo a seguir aplica um destaque em verde na maior célula valor do intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-173">The following example applies a green highlight to the highest value cell in the range.</span></span>


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

## <a name="multiple-formats-and-priority"></a><span data-ttu-id="523c4-175">Vários formatos e prioridades</span><span class="sxs-lookup"><span data-stu-id="523c4-175">Multiple formats and priority</span></span>

<span data-ttu-id="523c4-176">Você pode aplicar vários formatos condicionais em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-176">You can apply multiple conditional formats to a range.</span></span> <span data-ttu-id="523c4-177">Se os formatos tem elementos conflitantes, como cores de fonte diferentes apenas um formato aplica-se a esse elemento determinado.</span><span class="sxs-lookup"><span data-stu-id="523c4-177">If the formats have conflicting elements, such as differing font colors, only one format applies that particular element.</span></span> <span data-ttu-id="523c4-178">Precedência é definida pela propriedade `ConditionalFormat.priority`.</span><span class="sxs-lookup"><span data-stu-id="523c4-178">Precedence is defined by the `ConditionalFormat.priority` property.</span></span> <span data-ttu-id="523c4-179">Prioridade é um número (igual ao índice a `ConditionalFormatCollection`) e pode ser definido ao criar o formato.</span><span class="sxs-lookup"><span data-stu-id="523c4-179">Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format.</span></span> <span data-ttu-id="523c4-180">Quanto mais baixo o `priority` valor for, maior a prioridade do formato é.</span><span class="sxs-lookup"><span data-stu-id="523c4-180">The lowerer the `priority` value, the higher the priority of the format is.</span></span>

<span data-ttu-id="523c4-181">O exemplo a seguir mostra uma opção de cor da fonte conflitante entre os dois formatos.</span><span class="sxs-lookup"><span data-stu-id="523c4-181">The following example shows a conflicting font color choice between the two formats.</span></span> <span data-ttu-id="523c4-182">Números negativos receberão uma fonte em negrito, mas não a fonte vermelha, porque a prioridade é o formato que oferece uma fonte azul.</span><span class="sxs-lookup"><span data-stu-id="523c4-182">Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.</span></span>

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

### <a name="mutually-exclusive-conditional-formats"></a><span data-ttu-id="523c4-184">Formatos condicionais mutuamente exclusivos </span><span class="sxs-lookup"><span data-stu-id="523c4-184">Mutually exclusive conditional formats</span></span>

<span data-ttu-id="523c4-185">As `stopIfTrue` propriedade de `ConditionalFormat` impede que os formatos condicionais de prioridade inferiores sejam aplicados ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-185">The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range.</span></span> <span data-ttu-id="523c4-186">Quando um intervalo que corresponde ao formato condicional com `stopIfTrue === true` é aplicado, nenhum formato condicional subsequente é aplicado, mesmo se os detalhes da formatação não forem contraditórios.</span><span class="sxs-lookup"><span data-stu-id="523c4-186">When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.</span></span>

<span data-ttu-id="523c4-187">O exemplo a seguir mostra dois formatos condicionais adicionados a um intervalo.</span><span class="sxs-lookup"><span data-stu-id="523c4-187">The following example shows two conditional formats being added to a range.</span></span> <span data-ttu-id="523c4-188">Números negativos terão uma fonte azul com um fundo verde suave, independentemente da condição de formatação ser verdadeira.</span><span class="sxs-lookup"><span data-stu-id="523c4-188">Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="523c4-190">Confira também</span><span class="sxs-lookup"><span data-stu-id="523c4-190">See also</span></span>
-   [<span data-ttu-id="523c4-191">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="523c4-191">Fundamental programming concepts with the Excel JavaScript API</span></span>]( https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-core-concepts)
-   [<span data-ttu-id="523c4-192">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="523c4-192">Work with ranges using the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges)
-   [<span data-ttu-id="523c4-193">Objeto ConditionalFormat (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="523c4-193">ConditionalFormat Object (JavaScript API for Excel)</span></span>]( https://docs.microsoft.com/javascript/api/excel/excel.conditionalformat)
-   [<span data-ttu-id="523c4-194">Adicionar, alterar ou limpar formatações condicionais</span><span class="sxs-lookup"><span data-stu-id="523c4-194">Add, change, or clear conditional formats</span></span>](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
-   [<span data-ttu-id="523c4-195">Use fórmulas com o acesso condicional</span><span class="sxs-lookup"><span data-stu-id="523c4-195">Use formulas with conditional formatting</span></span>](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)
