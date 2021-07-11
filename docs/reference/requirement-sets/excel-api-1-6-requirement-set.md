---
title: Excel Conjunto de requisitos da API JavaScript 1.6
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.6.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bc2eb8f182a329808a46f172868b818027f5e367
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350103"
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
- Fornece prioridade `stopifTrue` e funcionalidade.
- Obtém a coleção de toda a formatação condicional em um determinado intervalo.
- Limpa todos os formatos condicionais ativos no intervalo atual especificado.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.6. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.6 ou anterior, consulte Excel APIs no conjunto de requisitos [1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspende o cálculo até que o próximo "context.sync()" seja chamado.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Especifica o objeto Rule neste formato condicional.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Os critérios da escala de cores.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Se for verdadeiro, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, ela terá dois (mínimo, máximo).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|O operador do formato condicional do valor da célula.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|O critério de escala de cores de ponto máximo.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|O critério de escala de cores de ponto mínimo.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Representação de código de cor HTML da cor da escala de cores (por exemplo, #FF0000 representa Vermelho).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Um número, uma fórmula ou nulo (se Type for LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Em que a fórmula condicional do critério deve se basear.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Código de cor HTML que representa a cor da linha de borda do formulário #RRGGBB (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Código de cor HTML que representa a cor de preenchimento, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Especifica se a Barra de Dados negativa tem a mesma cor de borda que a Barra de Dados positiva.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Especifica se a Barra de Dados negativa tem a mesma cor de preenchimento que a Barra de Dados positiva.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Código de cor HTML que representa a cor da linha de borda do formulário #RRGGBB (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Código de cor HTML que representa a cor de preenchimento, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Especifica se a Barra de Dados tem um gradiente.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|A fórmula, se necessário, para avaliar a regra databar.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|O tipo de regra para a barra de dados.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Exclui esse formato condicional.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Retorna o intervalo ao qual a formatação condicional é aplicada.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Retorna o intervalo ao que o formato conditonal é aplicado ou um objeto nulo se o formato condicional for aplicado a vários intervalos.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|A prioridade (ou índice) na coleção de formato condicional em que esse formato condicional existe no momento.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um tipo CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Retorna as propriedades de formato condicional do valor da célula se o formato condicional atual for um tipo CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Retorna as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Retorna as propriedades de formato condicional personalizadas se o formato condicional atual for um tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Retorna as propriedades de formato condicional personalizadas se o formato condicional atual for um tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Retorna as propriedades da barra de dados se o formato condicional atual for uma barra de dados.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Retorna as propriedades de formato condicional IconSet se o formato condicional atual for um tipo IconSet.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Retorna as propriedades de formato condicional IconSet se o formato condicional atual for um tipo IconSet.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|A prioridade do formato condicional na atual ConditionalFormatCollection.|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Retorna o formato condicional de critérios predefinidos.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Retorna o formato condicional de critérios predefinidos.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Retorna as propriedades de formato condicional de texto específico se o formato condicional atual for um tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Retorna as propriedades de formato condicional Superior/Inferior se o formato condicional atual for um tipo TopBottom.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Retorna as propriedades de formato condicional Superior/Inferior se o formato condicional atual for um tipo TopBottom.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Um tipo de formato condicional.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adiciona um novo formato condicional à coleção na prioridade primeiro/superior.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Retorna o número de formatos condicionais na guia de trabalho.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Retorna um formato condicional para o ID fornecido.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Retorna um formato condicional no índice fornecido.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Um número ou uma fórmula, dependendo do tipo.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual para cada um dos tipos de regra para o formato condicional Icon.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|No que a fórmula condicional de ícone deve se basear.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|O critério do formato condicional.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Código de cor HTML que representa a cor da linha de borda do formulário #RRGGBB (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valor constante que indica o lado específico da borda.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtém um objeto Border usando o respectivo índice.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtém a borda inferior.|
||[Count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Número de objetos de borda da coleção.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtém a borda esquerda.|
||[direita](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtém a borda direita.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtém a borda superior.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Redefine o preenchimento.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Código de cor HTML que representa a cor do preenchimento, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Especifica se a fonte está em negrito.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Redefine os formatos de fonte.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Especifica se a fonte é itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Especifica o status tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|O tipo de sublinhado aplicado à fonte.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Representa Excel código de formato de número para o intervalo determinado.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Coleção de objetos de borda que se aplicam ao intervalo geral de formato condicional.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Retorna o objeto fill definido no intervalo geral de formato condicional.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Retorna o objeto font definido no intervalo geral de formato condicional.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|O operador do formato condicional de texto.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|O valor de texto do formato condicional.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[classificação](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formatar valores com base na classificação superior ou inferior.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Especifica o objeto Rule neste formato condicional.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Código de cor HTML que representa a cor da linha Axis, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representação de como o eixo é determinado para uma Excel de dados.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Especifica a direção na qual o gráfico da barra de dados deve ser baseado.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representação de todos os valores à esquerda do eixo em uma Excel de dados.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representação de todos os valores à direita do eixo em uma Excel de dados.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Uma matriz de Criteria e IconSets para as regras e ícones personalizados potenciais para ícones condicionais.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Se for true, reverterá as ordens de ícone do IconSet.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Se definido, exibe a opção IconSet para o formato condicional.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|A regra da formatação condicional.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcula um intervalo de células em uma planilha.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|A coleção de ConditionalFormats que intercepta o intervalo.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|A regra da formatação condicional.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Retorna um objeto format, encapsulando a fonte de formatos condicionais, preenchimento, bordas e outras propriedades.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Os critérios do formato condicional Superior/Inferior.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcula todas as células em uma planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
