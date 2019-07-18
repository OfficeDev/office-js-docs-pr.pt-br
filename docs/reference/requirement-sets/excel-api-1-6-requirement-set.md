---
title: Conjunto de requisitos de API JavaScript do Excel 1,6
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,6
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e1a3375d19d8c1cb0fbddac50fabf826b96d7cc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771971"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Quais são as novidades na API JavaScript do Excel 1.6

## <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite os seguintes tipos de formatação condicional:

* Escala de cores
* Barra de dados
* Conjunto de ícones
* Personalizado

Além disso:

* Retorna o intervalo ao qual o formatato condicional é aplicada.
* Remoção da formatação condicional.
* Fornece a capacidade de priority e stopifTrue.
* Obtém a coleção de toda a formatação condicional em um determinado intervalo.
* Limpa todos os formatos condicionais ativos no intervalo atual especificado.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho, para garantir que todas as dependências sejam propagadas.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Representa o objeto Regra neste formato condicional.|
||[Set (Propriedades: Excel. CellValueConditionalFormat)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. CellValueConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[CellValueConditionalFormatData](/javascript/api/excel/excel.cellvalueconditionalformatdata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatdata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.cellvalueconditionalformatdata#rule)|Representa o objeto Regra neste formato condicional.|
|[CellValueConditionalFormatLoadOptions](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#rule)|Representa o objeto Regra neste formato condicional.|
|[CellValueConditionalFormatUpdateData](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#rule)|Representa o objeto Regra neste formato condicional.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Os critérios da escala de cores. O ponto médio é opcional ao usar uma escala de cores de dois pontos.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Se true, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, terá dois (mínimo, máximo).|
||[Set (Propriedades: Excel. ColorScaleConditionalFormat)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ColorScaleConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ColorScaleConditionalFormatData](/javascript/api/excel/excel.colorscaleconditionalformatdata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatdata#criteria)|Os critérios da escala de cores. O ponto médio é opcional ao usar uma escala de cores de dois pontos.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatdata#threecolorscale)|Se true, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, terá dois (mínimo, máximo).|
|[ColorScaleConditionalFormatLoadOptions](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#criteria)|Os critérios da escala de cores. O ponto médio é opcional ao usar uma escala de cores de dois pontos.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#threecolorscale)|Se true, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, terá dois (mínimo, máximo).|
|[ColorScaleConditionalFormatUpdateData](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata#criteria)|Os critérios da escala de cores. O ponto médio é opcional ao usar uma escala de cores de dois pontos.|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[operador](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|O operador do formato condicional de texto.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|O critério de escala de cores de ponto máximo.|
||[Central](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|O critério de escala de cores de ponto mínimo.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Representação do código de cor HTML da cor de escala de cores. Por exemplo #FF0000 representa vermelho.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Um número, uma fórmula ou nulo (se Type for LowestValue).|
||[tipo](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|O que a fórmula condicional de critério deve se basear.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|
||[Set (Propriedades: Excel. ConditionalDataBarNegativeFormat)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalDataBarNegativeFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ConditionalDataBarNegativeFormatData](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivebordercolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivefillcolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|
|[ConditionalDataBarNegativeFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivebordercolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivefillcolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|
|[ConditionalDataBarNegativeFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivebordercolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivefillcolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Representação booliana para indicar se a DataBar tem um gradiente ou não.|
||[Set (Propriedades: Excel. ConditionalDataBarPositiveFormat)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalDataBarPositiveFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ConditionalDataBarPositiveFormatData](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#gradientfill)|Representação booliana para indicar se a DataBar tem um gradiente ou não.|
|[ConditionalDataBarPositiveFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#gradientfill)|Representação booliana para indicar se a DataBar tem um gradiente ou não.|
|[ConditionalDataBarPositiveFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#gradientfill)|Representação booliana para indicar se a DataBar tem um gradiente ou não.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|A fórmula, se necessário, para avaliar a regra databar.|
||[tipo](/javascript/api/excel/excel.conditionaldatabarrule#type)|O tipo de regra para o databar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Exclui esse formato condicional.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Retorna o intervalo ao qual a formatação condicional é aplicada. Gera um erro se a formatação condicional for aplicada a vários intervalos. Somente leitura.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Retorna o intervalo ao qual o formato conditonal é aplicado, ou um objeto NULL, se o formato condicional for aplicado a vários intervalos. Somente leitura.|
||[prioriza](/javascript/api/excel/excel.conditionalformat#priority)|A prioridade (ou índice) dentro da coleção de formato condicional em que esse formato condicional existe atualmente no. Alterar isso também|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|
||[cliente](/javascript/api/excel/excel.conditionalformat#custom)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado. Somente leitura.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado. Somente leitura.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset. Somente leitura.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset. Somente leitura.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|A prioridade do formato condicional na atual ConditionalFormatCollection. Somente leitura.|
||[predefinido](/javascript/api/excel/excel.conditionalformat#preset)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[textcomparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[tipo](/javascript/api/excel/excel.conditionalformat#type)|Um tipo de formato condicional. Apenas um pode ser definido por vez. Somente leitura.|
||[Set (Propriedades: Excel. ConditionalFormat)](/javascript/api/excel/excel.conditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Add (tipo: "Custom" \| "databar" \| "ColorScale" \| " \| TopBottom \| " "PresetCriteria" \| "ContainsText" \| "cellvalue")](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adiciona um novo formato condicional à coleção na prioridade First/Top.|
||[Adicionar (tipo: Excel. Valorconditionalformattype)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adiciona um novo formato condicional à coleção na prioridade First/Top.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Retorna um formato condicional para o ID fornecido.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Retorna um formato condicional no índice fornecido.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ConditionalFormatCollectionLoadOptions](/javascript/api/excel/excel.conditionalformatcollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalue)|Para cada ITEM na coleção: retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalueornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[colorScale](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscale)|Para cada ITEM na coleção: retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscaleornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[cliente](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#custom)|Para cada ITEM na coleção: retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#customornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databar)|Para cada ITEM na coleção: retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databarornullobject)|Para cada ITEM na coleção: retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[iconSet](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconset)|Para cada ITEM na coleção: retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de conjunto de ícones.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconsetornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de conjunto de ícones.|
||[id](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#id)|Para cada ITEM na coleção: a prioridade do formato condicional dentro do ConditionalFormatCollection atual. Somente leitura.|
||[predefinido](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#preset)|Para cada ITEM na coleção: retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#presetornullobject)|Para cada ITEM na coleção: retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[prioriza](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#priority)|Para cada ITEM na coleção: a prioridade (ou índice) dentro da coleção de formato condicional que este formato condicional existe atualmente no. Alterar isso também|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#stopiftrue)|Para cada ITEM na coleção: se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa deverá ter efeito nessa célula.|
||[textcomparison](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparison)|Para cada ITEM na coleção: retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparisonornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottom)|Para cada ITEM na coleção: retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottomornullobject)|Para cada ITEM na coleção: retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[tipo](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#type)|Para cada ITEM na coleção: um tipo de formato condicional. Apenas um pode ser definido por vez. Somente leitura.|
|[ConditionalFormatData](/javascript/api/excel/excel.conditionalformatdata)|[cellValue](/javascript/api/excel/excel.conditionalformatdata#cellvalue)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatdata#cellvalueornullobject)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[colorScale](/javascript/api/excel/excel.conditionalformatdata#colorscale)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatdata#colorscaleornullobject)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|
||[cliente](/javascript/api/excel/excel.conditionalformatdata#custom)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado. Somente leitura.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatdata#customornullobject)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado. Somente leitura.|
||[dataBar](/javascript/api/excel/excel.conditionalformatdata#databar)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatdata#databarornullobject)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|
||[iconSet](/javascript/api/excel/excel.conditionalformatdata#iconset)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset. Somente leitura.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#iconsetornullobject)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset. Somente leitura.|
||[id](/javascript/api/excel/excel.conditionalformatdata#id)|A prioridade do formato condicional na atual ConditionalFormatCollection. Somente leitura.|
||[predefinido](/javascript/api/excel/excel.conditionalformatdata#preset)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#presetornullobject)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[prioriza](/javascript/api/excel/excel.conditionalformatdata#priority)|A prioridade (ou índice) dentro da coleção de formato condicional em que esse formato condicional existe atualmente no. Alterar isso também|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatdata#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
||[textcomparison](/javascript/api/excel/excel.conditionalformatdata#textcomparison)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatdata#textcomparisonornullobject)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformatdata#topbottom)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatdata#topbottomornullobject)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[tipo](/javascript/api/excel/excel.conditionalformatdata#type)|Um tipo de formato condicional. Apenas um pode ser definido por vez. Somente leitura.|
|[ConditionalFormatLoadOptions](/javascript/api/excel/excel.conditionalformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalue)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalueornullobject)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[colorScale](/javascript/api/excel/excel.conditionalformatloadoptions#colorscale)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#colorscaleornullobject)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[cliente](/javascript/api/excel/excel.conditionalformatloadoptions#custom)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#customornullobject)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformatloadoptions#databar)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#databarornullobject)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[iconSet](/javascript/api/excel/excel.conditionalformatloadoptions#iconset)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#iconsetornullobject)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset.|
||[id](/javascript/api/excel/excel.conditionalformatloadoptions#id)|A prioridade do formato condicional na atual ConditionalFormatCollection. Somente leitura.|
||[predefinido](/javascript/api/excel/excel.conditionalformatloadoptions#preset)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#presetornullobject)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[prioriza](/javascript/api/excel/excel.conditionalformatloadoptions#priority)|A prioridade (ou índice) dentro da coleção de formato condicional em que esse formato condicional existe atualmente no. Alterar isso também|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatloadoptions#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
||[textcomparison](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparison)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparisonornullobject)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformatloadoptions#topbottom)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#topbottomornullobject)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[tipo](/javascript/api/excel/excel.conditionalformatloadoptions#type)|Um tipo de formato condicional. Apenas um pode ser definido por vez. Somente leitura.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
||[Set (Propriedades: Excel. ConditionalFormatRule)](/javascript/api/excel/excel.conditionalformatrule#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalFormatRuleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalformatrule#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ConditionalFormatRuleData](/javascript/api/excel/excel.conditionalformatruledata)|[formula](/javascript/api/excel/excel.conditionalformatruledata#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruledata#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruledata#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
|[ConditionalFormatRuleLoadOptions](/javascript/api/excel/excel.conditionalformatruleloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatruleloadoptions#$all)||
||[formula](/javascript/api/excel/excel.conditionalformatruleloadoptions#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleloadoptions#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleloadoptions#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
|[ConditionalFormatRuleUpdateData](/javascript/api/excel/excel.conditionalformatruleupdatedata)|[formula](/javascript/api/excel/excel.conditionalformatruleupdatedata#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleupdatedata#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleupdatedata#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
|[ConditionalFormatUpdateData](/javascript/api/excel/excel.conditionalformatupdatedata)|[cellValue](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalue)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalueornullobject)|Retorna as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo Cellvalue.|
||[colorScale](/javascript/api/excel/excel.conditionalformatupdatedata#colorscale)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#colorscaleornullobject)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[cliente](/javascript/api/excel/excel.conditionalformatupdatedata#custom)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#customornullobject)|Retorna as propriedades de formato condicional personalizado se o formato condicional atual for um tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformatupdatedata#databar)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#databarornullobject)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[iconSet](/javascript/api/excel/excel.conditionalformatupdatedata#iconset)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#iconsetornullobject)|Retorna as propriedades de formato condicional do Iconset se o formato condicional atual for um tipo de Íconeset.|
||[predefinido](/javascript/api/excel/excel.conditionalformatupdatedata#preset)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#presetornullobject)|Retorna o formato condicional de critérios predefinidos. Confira Excel. PresetCriteriaConditionalFormat para obter mais detalhes.|
||[prioriza](/javascript/api/excel/excel.conditionalformatupdatedata#priority)|A prioridade (ou índice) dentro da coleção de formato condicional em que esse formato condicional existe atualmente no. Alterar isso também|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatupdatedata#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
||[textcomparison](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparison)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparisonornullobject)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformatupdatedata#topbottom)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#topbottomornullobject)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um tipo TopBottom.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Um número ou uma fórmula, dependendo do tipo.|
||[operador](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual para cada tipo de regra para o formato condicional de ícone.|
||[tipo](/javascript/api/excel/excel.conditionaliconcriterion#type)|No que a fórmula condicional de ícone deve se basear.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critério](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|O critério do formato condicional.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valor constante que indica o lado específico da borda. Consulte Excel. ConditionalRangeBorderIndex para obter detalhes. Somente leitura.|
||[Set (Propriedades: Excel. ConditionalRangeBorder)](/javascript/api/excel/excel.conditionalrangeborder#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalRangeBorderUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangeborder#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "EdgeRight")](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItem (index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtém um objeto Border usando o respectivo índice.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtém a borda inferior. Somente leitura.|
||[Count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Número de objetos de borda da coleção. Somente leitura.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtém a borda esquerda. Somente leitura.|
||[direita](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtém a borda direita. Somente leitura.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtém a borda superior. Somente leitura.|
|[ConditionalRangeBorderCollectionLoadOptions](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#color)|Para cada ITEM na coleção: código de cor HTML que representa a cor da linha de borda, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#sideindex)|Para cada ITEM da coleção: valor constante que indica o lado específico da borda. Consulte Excel. ConditionalRangeBorderIndex para obter detalhes. Somente leitura.|
||[style](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#style)|Para cada ITEM na coleção: uma das constantes de estilo de linha que especifica o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeBorderCollectionUpdateData](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#bottom)|Obtém a borda inferior.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#left)|Obtém a borda esquerda.|
||[direita](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#right)|Obtém a borda direita.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#top)|Obtém a borda superior.|
|[ConditionalRangeBorderData](/javascript/api/excel/excel.conditionalrangeborderdata)|[color](/javascript/api/excel/excel.conditionalrangeborderdata#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderdata#sideindex)|Valor constante que indica o lado específico da borda. Consulte Excel. ConditionalRangeBorderIndex para obter detalhes. Somente leitura.|
||[style](/javascript/api/excel/excel.conditionalrangeborderdata#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeBorderLoadOptions](/javascript/api/excel/excel.conditionalrangeborderloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangeborderloadoptions#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderloadoptions#sideindex)|Valor constante que indica o lado específico da borda. Consulte Excel. ConditionalRangeBorderIndex para obter detalhes. Somente leitura.|
||[style](/javascript/api/excel/excel.conditionalrangeborderloadoptions#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeBorderUpdateData](/javascript/api/excel/excel.conditionalrangeborderupdatedata)|[color](/javascript/api/excel/excel.conditionalrangeborderupdatedata#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[style](/javascript/api/excel/excel.conditionalrangeborderupdatedata#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Redefine o preenchimento.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[Set (Propriedades: Excel. ConditionalRangeFill)](/javascript/api/excel/excel.conditionalrangefill#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalRangeFillUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangefill#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ConditionalRangeFillData](/javascript/api/excel/excel.conditionalrangefilldata)|[color](/javascript/api/excel/excel.conditionalrangefilldata#color)|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFillLoadOptions](/javascript/api/excel/excel.conditionalrangefillloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangefillloadoptions#color)|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFillUpdateData](/javascript/api/excel/excel.conditionalrangefillupdatedata)|[color](/javascript/api/excel/excel.conditionalrangefillupdatedata#color)|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Representa o status da fonte em negrito.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Redefine os formatos de fonte.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Representa o status da fonte em itálico.|
||[Set (Propriedades: Excel. ConditionalRangeFont)](/javascript/api/excel/excel.conditionalrangefont#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalRangeFontUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangefont#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Representa o status de tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ConditionalRangeFontUnderlineStyle para obter detalhes.|
|[ConditionalRangeFontData](/javascript/api/excel/excel.conditionalrangefontdata)|[bold](/javascript/api/excel/excel.conditionalrangefontdata#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.conditionalrangefontdata#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.conditionalrangefontdata#italic)|Representa o status da fonte em itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontdata#strikethrough)|Representa o status de tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefontdata#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ConditionalRangeFontUnderlineStyle para obter detalhes.|
|[ConditionalRangeFontLoadOptions](/javascript/api/excel/excel.conditionalrangefontloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.conditionalrangefontloadoptions#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.conditionalrangefontloadoptions#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.conditionalrangefontloadoptions#italic)|Representa o status da fonte em itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontloadoptions#strikethrough)|Representa o status de tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefontloadoptions#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ConditionalRangeFontUnderlineStyle para obter detalhes.|
|[ConditionalRangeFontUpdateData](/javascript/api/excel/excel.conditionalrangefontupdatedata)|[bold](/javascript/api/excel/excel.conditionalrangefontupdatedata#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.conditionalrangefontupdatedata#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.conditionalrangefontupdatedata#italic)|Representa o status da fonte em itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontupdatedata#strikethrough)|Representa o status de tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefontupdatedata#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ConditionalRangeFontUnderlineStyle para obter detalhes.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado. Desmarcada se NULL for passado.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Coleção de objetos Border que se aplicam ao intervalo de formato condicional geral. Somente leitura.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Retorna o objeto Fill definido no intervalo de formato condicional geral. Somente leitura.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Retorna o objeto Font definido no intervalo de formato condicional geral. Somente leitura.|
||[Set (Propriedades: Excel. ConditionalRangeFormat)](/javascript/api/excel/excel.conditionalrangeformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ConditionalRangeFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.conditionalrangeformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ConditionalRangeFormatData](/javascript/api/excel/excel.conditionalrangeformatdata)|[Borders](/javascript/api/excel/excel.conditionalrangeformatdata#borders)|Coleção de objetos Border que se aplicam ao intervalo de formato condicional geral. Somente leitura.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatdata#fill)|Retorna o objeto Fill definido no intervalo de formato condicional geral. Somente leitura.|
||[font](/javascript/api/excel/excel.conditionalrangeformatdata#font)|Retorna o objeto Font definido no intervalo de formato condicional geral. Somente leitura.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatdata#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado. Desmarcada se NULL for passado.|
|[ConditionalRangeFormatLoadOptions](/javascript/api/excel/excel.conditionalrangeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeformatloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.conditionalrangeformatloadoptions#borders)|Coleção de objetos Border que se aplicam ao intervalo de formato condicional geral.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatloadoptions#fill)|Retorna o objeto Fill definido no intervalo de formato condicional geral.|
||[font](/javascript/api/excel/excel.conditionalrangeformatloadoptions#font)|Retorna o objeto Font definido no intervalo de formato condicional geral.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatloadoptions#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado. Desmarcada se NULL for passado.|
|[ConditionalRangeFormatUpdateData](/javascript/api/excel/excel.conditionalrangeformatupdatedata)|[Borders](/javascript/api/excel/excel.conditionalrangeformatupdatedata#borders)|Coleção de objetos Border que se aplicam ao intervalo de formato condicional geral.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatupdatedata#fill)|Retorna o objeto Fill definido no intervalo de formato condicional geral.|
||[font](/javascript/api/excel/excel.conditionalrangeformatupdatedata#font)|Retorna o objeto Font definido no intervalo de formato condicional geral.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatupdatedata#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado. Desmarcada se NULL for passado.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operador](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|O operador do formato condicional de texto.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|O valor de texto do formato condicional.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[Classificação](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|
||[tipo](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formatar valores com base na classificação superior ou inferior.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.customconditionalformat#rule)|Representa o objeto Regra neste formato condicional. Somente leitura.|
||[Set (Propriedades: Excel. CustomConditionalFormat)](/javascript/api/excel/excel.customconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. CustomConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.customconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[CustomConditionalFormatData](/javascript/api/excel/excel.customconditionalformatdata)|[format](/javascript/api/excel/excel.customconditionalformatdata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.customconditionalformatdata#rule)|Representa o objeto Regra neste formato condicional. Somente leitura.|
|[CustomConditionalFormatLoadOptions](/javascript/api/excel/excel.customconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.customconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.customconditionalformatloadoptions#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.customconditionalformatloadoptions#rule)|Representa o objeto Regra neste formato condicional.|
|[CustomConditionalFormatUpdateData](/javascript/api/excel/excel.customconditionalformatupdatedata)|[format](/javascript/api/excel/excel.customconditionalformatupdatedata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.customconditionalformatupdatedata#rule)|Representa o objeto Regra neste formato condicional.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representação de como o eixo é determinado para uma barra de dados do Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Representa a direção na qual o gráfico da barra de dados deve se basear.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|
||[Set (Propriedades: Excel. DataBarConditionalFormat)](/javascript/api/excel/excel.databarconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. DataBarConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.databarconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[DataBarConditionalFormatData](/javascript/api/excel/excel.databarconditionalformatdata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatdata#axiscolor)|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatdata#axisformat)|Representação de como o eixo é determinado para uma barra de dados do Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatdata#bardirection)|Representa a direção na qual o gráfico da barra de dados deve se basear.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatdata#negativeformat)|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatdata#positiveformat)|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatdata#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[DataBarConditionalFormatLoadOptions](/javascript/api/excel/excel.databarconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.databarconditionalformatloadoptions#$all)||
||[axisColor](/javascript/api/excel/excel.databarconditionalformatloadoptions#axiscolor)|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#axisformat)|Representação de como o eixo é determinado para uma barra de dados do Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatloadoptions#bardirection)|Representa a direção na qual o gráfico da barra de dados deve se basear.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#negativeformat)|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#positiveformat)|Representação de todos os valores à direita do eixo em uma barra de dados do Excel.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatloadoptions#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[DataBarConditionalFormatUpdateData](/javascript/api/excel/excel.databarconditionalformatupdatedata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatupdatedata#axiscolor)|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#axisformat)|Representação de como o eixo é determinado para uma barra de dados do Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatupdatedata#bardirection)|Representa a direção na qual o gráfico da barra de dados deve se basear.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#negativeformat)|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#positiveformat)|Representação de todos os valores à direita do eixo em uma barra de dados do Excel.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatupdatedata#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Uma matriz de critérios e IconSets para as regras e possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto Type, Formula e Operator serão ignorados quando set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Se true, inverte as ordens de ícone para o Iconset. Observe que isso não poderá ser definido se os ícones personalizados forem usados.|
||[Set (Propriedades: Excel. IconSetConditionalFormat)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. IconSetConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Se definido, exibe a opção Iconset para o formato condicional.|
|[IconSetConditionalFormatData](/javascript/api/excel/excel.iconsetconditionalformatdata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatdata#criteria)|Uma matriz de critérios e IconSets para as regras e possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto Type, Formula e Operator serão ignorados quando set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatdata#reverseiconorder)|Se true, inverte as ordens de ícone para o Iconset. Observe que isso não poderá ser definido se os ícones personalizados forem usados.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatdata#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatdata#style)|Se definido, exibe a opção Iconset para o formato condicional.|
|[IconSetConditionalFormatLoadOptions](/javascript/api/excel/excel.iconsetconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#criteria)|Uma matriz de critérios e IconSets para as regras e possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto Type, Formula e Operator serão ignorados quando set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#reverseiconorder)|Se true, inverte as ordens de ícone para o Iconset. Observe que isso não poderá ser definido se os ícones personalizados forem usados.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#style)|Se definido, exibe a opção Iconset para o formato condicional.|
|[IconSetConditionalFormatUpdateData](/javascript/api/excel/excel.iconsetconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#criteria)|Uma matriz de critérios e IconSets para as regras e possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto Type, Formula e Operator serão ignorados quando set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#reverseiconorder)|Se true, inverte as ordens de ícone para o Iconset. Observe que isso não poderá ser definido se os ícones personalizados forem usados.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#style)|Se definido, exibe a opção Iconset para o formato condicional.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|A regra da formatação condicional.|
||[Set (Propriedades: Excel. PresetCriteriaConditionalFormat)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PresetCriteriaConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[PresetCriteriaConditionalFormatData](/javascript/api/excel/excel.presetcriteriaconditionalformatdata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#rule)|A regra da formatação condicional.|
|[PresetCriteriaConditionalFormatLoadOptions](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#rule)|A regra da formatação condicional.|
|[PresetCriteriaConditionalFormatUpdateData](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#rule)|A regra da formatação condicional.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcula um intervalo de células em uma planilha.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Conjunto de ConditionalFormats que interseccionam o intervalo. Somente leitura.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[conditionalFormats](/javascript/api/excel/excel.rangedata#conditionalformats)|Conjunto de ConditionalFormats que interseccionam o intervalo. Somente leitura.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.textconditionalformat#rule)|A regra da formatação condicional.|
||[Set (Propriedades: Excel. TextConditionalFormat)](/javascript/api/excel/excel.textconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TextConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.textconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[TextConditionalFormatData](/javascript/api/excel/excel.textconditionalformatdata)|[format](/javascript/api/excel/excel.textconditionalformatdata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.textconditionalformatdata#rule)|A regra da formatação condicional.|
|[TextConditionalFormatLoadOptions](/javascript/api/excel/excel.textconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.textconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.textconditionalformatloadoptions#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.textconditionalformatloadoptions#rule)|A regra da formatação condicional.|
|[TextConditionalFormatUpdateData](/javascript/api/excel/excel.textconditionalformatupdatedata)|[format](/javascript/api/excel/excel.textconditionalformatupdatedata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.textconditionalformatupdatedata#rule)|A regra da formatação condicional.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Os critérios do formato condicional superior/inferior.|
||[Set (Propriedades: Excel. TopBottomConditionalFormat)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TopBottomConditionalFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[TopBottomConditionalFormatData](/javascript/api/excel/excel.topbottomconditionalformatdata)|[format](/javascript/api/excel/excel.topbottomconditionalformatdata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.topbottomconditionalformatdata#rule)|Os critérios do formato condicional superior/inferior.|
|[TopBottomConditionalFormatLoadOptions](/javascript/api/excel/excel.topbottomconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#rule)|Os critérios do formato condicional superior/inferior.|
|[TopBottomConditionalFormatUpdateData](/javascript/api/excel/excel.topbottomconditionalformatupdatedata)|[format](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#rule)|Os critérios do formato condicional superior/inferior.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calcular (markAllDirty: booliano)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcula todas as células em uma planilha.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
