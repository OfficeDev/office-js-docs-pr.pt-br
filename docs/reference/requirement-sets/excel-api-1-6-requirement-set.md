---
title: Excel conjunto de requisitos da API JavaScript 1.6
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: d68bfae3494ec21df1eee5909ac2df532a0537b9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745838"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Quais são as novidades na API JavaScript do Excel 1.6

## <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite os seguintes tipos de formatação condicional.

- Escala de cores
- Barra de dados
- Conjunto de ícones
- Personalizado

Além disso:

- Retorna o intervalo ao qual o formatato condicional é aplicada.
- Remoção da formatação condicional.
- Fornece prioridade e `stopifTrue` funcionalidade.
- Obtém a coleção de toda a formatação condicional em um determinado intervalo.
- Limpa todos os formatos condicionais ativos no intervalo atual especificado.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.6. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.6 ou anterior, consulte Excel APIs no conjunto de requisitos [1.6 ou anterior](/javascript/api/excel?view=excel-js-1.6&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendapicalculationuntilnextsync-member(1))|Suspende o cálculo até que o próximo `context.sync()` seja chamado.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-format-member)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-rule-member)|Especifica o objeto rule neste formato condicional.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[critério](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-criteria-member)|Os critérios da escala de cores.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-threecolorscale-member)|Se `true`, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, ela terá dois (mínimo, máximo).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula1-member)|A fórmula, se necessário, na qual avaliar a regra de formato condicional.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula2-member)|A fórmula, se necessário, na qual avaliar a regra de formato condicional.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-operator-member)|O operador do formato condicional do valor da célula.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-maximum-member)|O ponto máximo do critério de escala de cores.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-midpoint-member)|O ponto médio do critério de escala de cores, se a escala de cores for uma escala de 3 cores.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-minimum-member)|O ponto mínimo do critério de escala de cores.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-color-member)|Representação de código de cor HTML da cor da escala de cores (por exemplo, #FF0000 representa Vermelho).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-formula-member)|Um número, uma fórmula ou `null` (se `type` for `lowestValue`).|
||[tipo](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-type-member)|Em que a fórmula condicional do critério deve se basear.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-bordercolor-member)|Código de cor HTML que representa a cor da linha de borda, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-fillcolor-member)|Código de cor HTML que representa a cor de preenchimento, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivebordercolor-member)|Especifica se a barra de dados negativa tem a mesma cor de borda que a barra de dados positiva.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivefillcolor-member)|Especifica se a barra de dados negativa tem a mesma cor de preenchimento que a barra de dados positiva.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-bordercolor-member)|Código de cor HTML que representa a cor da linha de borda, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-fillcolor-member)|Código de cor HTML que representa a cor de preenchimento, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-gradientfill-member)|Especifica se a barra de dados tem um gradiente.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-formula-member)|A fórmula, se necessário, na qual avaliar a regra da barra de dados.|
||[tipo](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-type-member)|O tipo de regra para a barra de dados.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalue-member)|Retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um `CellValue` tipo.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalueornullobject-member)|Retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um `CellValue` tipo.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscale-member)|Retorna as propriedades de formato condicional da escala de cores se o formato condicional atual for um `ColorScale` tipo.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscaleornullobject-member)|Retorna as propriedades de formato condicional da escala de cores se o formato condicional atual for um `ColorScale` tipo.|
||[custom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-custom-member)|Retorna as propriedades de formato condicional personalizadas se o formato condicional atual for um tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-customornullobject-member)|Retorna as propriedades de formato condicional personalizadas se o formato condicional atual for um tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databar-member)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databarornullobject-member)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[delete()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-delete-member(1))|Exclui esse formato condicional.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrange-member(1))|Retorna o intervalo ao qual a formatação condicional é aplicada.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrangeornullobject-member(1))|Retorna o intervalo ao qual o formato conditonal é aplicado.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconset-member)|Retorna as propriedades de formato condicional do conjunto de ícones se o formato condicional atual for um `IconSet` tipo.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconsetornullobject-member)|Retorna as propriedades de formato condicional do conjunto de ícones se o formato condicional atual for um `IconSet` tipo.|
||[id](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-id-member)|A prioridade do formato condicional no `ConditionalFormatCollection`atual .|
||[preset](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-preset-member)|Retorna o formato condicional de critérios predefinidos.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-presetornullobject-member)|Retorna o formato condicional de critérios predefinidos.|
||[priority](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-priority-member)|A prioridade (ou índice) na coleção de formato condicional em que esse formato condicional existe no momento.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-stopiftrue-member)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparison-member)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparisonornullobject-member)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottom-member)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um `TopBottom` tipo.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottomornullobject-member)|Retorna as propriedades de formato condicional superior/inferior se o formato condicional atual for um `TopBottom` tipo.|
||[tipo](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-type-member)|Um tipo de formato condicional.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-add-member(1))|Adiciona um novo formato condicional à coleção na prioridade primeiro/superior.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-clearall-member(1))|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getcount-member(1))|Retorna o número de formatos condicionais na guia de trabalho.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitem-member(1))|Retorna um formato condicional para o ID fornecido.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemat-member(1))|Retorna um formato condicional no índice fornecido.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formula-member)|A fórmula, se necessário, na qual avaliar a regra de formato condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formulalocal-member)|A fórmula, se necessário, na qual avaliar a regra de formato condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formular1c1-member)|A fórmula, se necessário, na qual avaliar a regra de formato condicional na notação de estilo R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-customicon-member)|O ícone personalizado do critério atual, se diferente do conjunto de ícones padrão, será `null` retornado.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-formula-member)|Um número ou uma fórmula, dependendo do tipo.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-operator-member)|`greaterThan` ou `greaterThanOrEqual` para cada um dos tipos de regra para o formato condicional do ícone.|
||[tipo](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-type-member)|No que a fórmula condicional de ícone deve se basear.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#excel-excel-conditionalpresetcriteriarule-criterion-member)|O critério do formato condicional.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-color-member)|Código de cor HTML que representa a cor da linha de borda, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-sideindex-member)|Valor constante que indica o lado específico da borda.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-style-member)|Uma das constantes de estilo de linha especificando o estilo de linha da borda.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-bottom-member)|Obtém a borda inferior.|
||[Count](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-count-member)|Número de objetos de borda da coleção.|
||[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitem-member(1))|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitemat-member(1))|Obtém um objeto Border usando o respectivo índice.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-left-member)|Obtém a borda esquerda.|
||[direita](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-right-member)|Obtém a borda direita.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-top-member)|Obtém a borda superior.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-clear-member(1))|Redefine o preenchimento.|
||[color](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-color-member)|Código de cor HTML que representa a cor do preenchimento, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-bold-member)|Especifica se a fonte está em negrito.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-clear-member(1))|Redefine os formatos de fonte.|
||[color](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-color-member)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-italic-member)|Especifica se a fonte é itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-strikethrough-member)|Especifica o status tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-underline-member)|O tipo de sublinhado aplicado à fonte.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[Borders](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-borders-member)|Coleção de objetos de borda que se aplicam ao intervalo geral de formato condicional.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-fill-member)|Retorna o objeto fill definido no intervalo geral de formato condicional.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-font-member)|Retorna o objeto font definido no intervalo geral de formato condicional.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-numberformat-member)|Representa Excel código de formato de número para o intervalo determinado.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-operator-member)|O operador do formato condicional de texto.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-text-member)|O valor de texto do formato condicional.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[classificação](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-rank-member)|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|
||[tipo](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-type-member)|Formatar valores com base na classificação superior ou inferior.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-format-member)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-rule-member)|Especifica o objeto `Rule` nesse formato condicional.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axiscolor-member)|Código de cor HTML que representa a cor da linha Axis, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axisformat-member)|Representação de como o eixo é determinado para uma Excel de dados.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-bardirection-member)|Especifica a direção na qual o gráfico da barra de dados deve ser baseado.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-lowerboundrule-member)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-negativeformat-member)|Representação de todos os valores à esquerda do eixo em uma Excel de dados.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-positiveformat-member)|Representação de todos os valores à direita do eixo em uma Excel de dados.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-showdatabaronly-member)|If `true`, oculta os valores das células onde a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-upperboundrule-member)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[critério](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-criteria-member)|Uma matriz de critérios e conjuntos de ícones para as regras e ícones personalizados potenciais para ícones condicionais.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-reverseiconorder-member)|Se `true`, inverte as ordens de ícone para o conjunto de ícones.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-showicononly-member)|Se `true`, oculta os valores e mostra apenas ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-style-member)|Se definido, exibe a opção de conjunto de ícones para o formato condicional.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-format-member)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-rule-member)|A regra da formatação condicional.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#excel-excel-range-calculate-member(1))|Calcula um intervalo de células em uma planilha.|
||[conditionalFormats](/javascript/api/excel/excel.range#excel-excel-range-conditionalformats-member)|A coleção de `ConditionalFormats` que intercepta o intervalo.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-format-member)|Retorna um objeto format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades do formato condicional.|
||[rule](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-rule-member)|A regra da formatação condicional.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-format-member)|Retorna um objeto format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades do formato condicional.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-rule-member)|Os critérios do formato condicional superior/inferior.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-calculate-member(1))|Calcula todas as células em uma planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
