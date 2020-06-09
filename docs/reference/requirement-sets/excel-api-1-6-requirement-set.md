---
title: Conjunto de requisitos de API JavaScript do Excel 1,6
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,6
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f86a470d39bacfe4940a6c225b9ce7d8903e2092
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611411"
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
* Oferece prioridade e `stopifTrue` capacidade.
* Obtém a coleção de toda a formatação condicional em um determinado intervalo.
* Limpa todos os formatos condicionais ativos no intervalo atual especificado.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,6. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,6 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,6 ou anterior](/javascript/api/excel?view=excel-js-1.6).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho, para garantir que todas as dependências sejam propagadas.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Representa o objeto Regra neste formato condicional.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Os critérios da escala de cores. O ponto médio é opcional ao usar uma escala de cores de dois pontos.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Se true, a escala de cores terá três pontos (mínimo, ponto médio, máximo), caso contrário, terá dois (mínimo, máximo).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[operador](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|O operador do formato condicional de texto.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|O critério de escala de cores de ponto máximo.|
||[Central](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|O critério de escala de cores de ponto mínimo.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Representação do código de cor HTML da cor de escala de cores. Por exemplo #FF0000 representa vermelho.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Um número, uma fórmula ou nulo (se Type for LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|O que a fórmula condicional de critério deve se basear.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Representação booliana para indicar se a DataBar tem um gradiente ou não.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|A fórmula, se necessário, para avaliar a regra databar.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|O tipo de regra para o databar.|
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
||[type](/javascript/api/excel/excel.conditionalformat#type)|Um tipo de formato condicional. Apenas um pode ser definido por vez. Somente leitura.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Adicionar (tipo: Excel. Valorconditionalformattype)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adiciona um novo formato condicional à coleção na prioridade First/Top.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Retorna um formato condicional para o ID fornecido.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Retorna um formato condicional no índice fornecido.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|A fórmula, se necessário, para avaliar a regra de formatação condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Um número ou uma fórmula, dependendo do tipo.|
||[operador](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan ou GreaterThanOrEqual para cada tipo de regra para o formato condicional de ícone.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|No que a fórmula condicional de ícone deve se basear.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[critério](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|O critério do formato condicional.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valor constante que indica o lado específico da borda. Consulte Excel. ConditionalRangeBorderIndex para obter detalhes. Somente leitura.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtém um objeto Border usando o respectivo índice.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtém a borda inferior. Somente leitura.|
||[Count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Número de objetos de borda da coleção. Somente leitura.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtém a borda esquerda. Somente leitura.|
||[direita](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtém a borda direita. Somente leitura.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtém a borda superior. Somente leitura.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Redefine o preenchimento.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Representa o status da fonte em negrito.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Redefine os formatos de fonte.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Representa o status da fonte em itálico.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Representa o status de tachado da fonte.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ConditionalRangeFontUnderlineStyle para obter detalhes.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado. Desmarcada se NULL for passado.|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Coleção de objetos Border que se aplicam ao intervalo de formato condicional geral. Somente leitura.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Retorna o objeto Fill definido no intervalo de formato condicional geral. Somente leitura.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Retorna o objeto Font definido no intervalo de formato condicional geral. Somente leitura.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operador](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|O operador do formato condicional de texto.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|O valor de texto do formato condicional.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[Classificação](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formatar valores com base na classificação superior ou inferior.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.customconditionalformat#rule)|Representa o objeto Regra neste formato condicional. Somente leitura.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representação de como o eixo é determinado para uma barra de dados do Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Representa a direção na qual o gráfico da barra de dados deve se basear.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Uma matriz de critérios e IconSets para as regras e possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto Type, Formula e Operator serão ignorados quando set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Se true, inverte as ordens de ícone para o Iconset. Observe que isso não poderá ser definido se os ícones personalizados forem usados.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Caso verdadeiro, oculta os valores e mostra somente ícones.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Se definido, exibe a opção Iconset para o formato condicional.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais.|
||[norma](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|A regra da formatação condicional.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcula um intervalo de células em uma planilha.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Conjunto de ConditionalFormats que interseccionam o intervalo. Somente leitura.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.textconditionalformat#rule)|A regra da formatação condicional.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|
||[norma](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Os critérios do formato condicional superior/inferior.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calcular (markAllDirty: booliano)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcula todas as células em uma planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.6)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
