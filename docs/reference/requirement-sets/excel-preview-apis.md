---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel
ms.date: 03/19/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fda0721bd5d7cbec6349c4800a97132d61a26ab9
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891198"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Configurações de cultura](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings-preview) | Obtém configurações culturais do sistema para a pasta de trabalho, como formatação de número. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Inserir pasta de trabalho](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insira uma pasta de trabalho em outra.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Filtros dinâmicos | Aplica filtros orientados a valores aos campos de uma tabela dinâmica. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) a pasta de trabalho | Salve e feche a pasta de trabalho.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Excel atualmente em versão prévia. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs previamente lançadas), consulte [todas as APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fornece informações com base nas configurações de cultura do sistema atual. Isso inclui os nomes de cultura, a formatação de números e outras configurações dependentes de cultura.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos. Isso é baseado nas configurações locais do Excel.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos. Isso é baseado nas configurações locais do Excel.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Especifica se os separadores de sistema do Microsoft Excel estão habilitados.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Representa o ângulo no qual o texto é orientado para o título do eixo do gráfico. O valor deve ser um inteiro de-90 a 90 ou o inteiro 180 para texto orientado verticalmente.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimensão: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtém os valores de uma única dimensão da série de gráficos. Podem ser valores de categoria ou valores de dados, dependendo da dimensão especificada e de como os dados são mapeados para a série de gráficos.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtém o tipo de conteúdo do comentário.|
||[Obtido](/javascript/api/excel/excel.comment#resolved)|Obtém ou define o status do thread de comentários. O valor "true" significa que o thread está resolvido.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Obtém o tipo de conteúdo da resposta.|
||[Obtido](/javascript/api/excel/excel.commentreply#resolved)|Obtém ou define o status da resposta. O valor "true" significa que a resposta está no estado resolvido.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Define o formato culturalmente apropriado para exibir data e hora. Isso é baseado nas configurações atuais de cultura do sistema.|
||[name](/javascript/api/excel/excel.cultureinfo#name)|Obtém o nome da cultura no formato languagecode2-Country/regioncode2 (por exemplo, "zh-CN" ou "en-US"). Isso é baseado nas configurações atuais do sistema.|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Define o formato culturalmente apropriado para exibir números. Isso é baseado nas configurações atuais de cultura do sistema.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtém a cadeia de caracteres usada como o separador de data. Isso é baseado nas configurações atuais do sistema.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data longa. Isso é baseado nas configurações atuais do sistema.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtém a cadeia de caracteres de formato para um valor de tempo longo. Isso é baseado nas configurações atuais do sistema.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data abreviada. Isso é baseado nas configurações atuais do sistema.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtém a cadeia de caracteres usada como o separador de tempo. Isso é baseado nas configurações atuais do sistema.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos. Isso é baseado nas configurações atuais do sistema.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos. Isso é baseado nas configurações atuais do sistema.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparador](/javascript/api/excel/excel.pivotdatefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotdatefilter#condition)|Indica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|O limite inferior do intervalo para a `Between` condição de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|O limite superior do intervalo para a `Between` condição de filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Para as condições `Equals`de filtro `Between` ,, e, indica se as comparações devem ser feitas como dias inteiros. `Before` `After`|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filtro: PivotValueFilter \| PivotLabelFilter \| PivotManualFilter \| PivotDateFilter \| PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Define um ou vários dos PivotFilters atuais do campo e os aplica ao campo.|
||[clearAllFilters ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Limpa todos os critérios de todos os filtros de campo. Isso removerá qualquer filtragem ativa no campo.|
||[clearFilter (FilterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Limpa todos os critérios existentes do filtro do campo de determinado tipo (se houver algum aplicado no momento).|
||[GetFilters ()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtém todos os filtros aplicados no campo no momento.|
||[IsFiltered (FilterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Verifica se há filtros aplicados no campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|O filtro de data atualmente aplicado ao PivotField. Nulo se nenhum for aplicado.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|O filtro de rótulo do PivotField atualmente aplicado. Nulo se nenhum for aplicado.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|O filtro manual aplicado no momento do PivotField. Nulo se nenhum for aplicado.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|O filtro de valor atualmente aplicado ao PivotField. Nulo se nenhum for aplicado.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparador](/javascript/api/excel/excel.pivotlabelfilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotlabelfilter#condition)|Indica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|O limite inferior do intervalo para a condição de filtro between.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|A subcadeia de caracteres `BeginsWith`usada `EndsWith`para as `Contains` condições de filtro, e.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|O limite superior do intervalo para a condição de filtro entre.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias. A célula retornada é a interseção da linha e coluna fornecidas que contém os dados da hierarquia especificada. Esse método é o inverso de chamar getPivotItems e getDataHierarchy em uma célula específica.|
||[tabela dinâmica](/javascript/api/excel/excel.pivotlayout#pivotstyle)|O estilo aplicado à tabela dinâmica.|
||[setStyle (Style: String \| pivotstyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Define o estilo aplicado à tabela dinâmica.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Uma lista de itens selecionados a serem filtrados manualmente. Eles devem ser itens válidos e existentes do campo escolhido.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica se a tabela dinâmica permite o aplicativo de vários PivotFilters em um determinado campo PivotField na tabela.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtém a primeira tabela dinâmica na coleção. As tabelas dinâmicas da coleção são classificadas de cima para baixo e da esquerda para a direita, de forma que a tabela superior esquerda seja a primeira tabela dinâmica na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtém uma Tabela Dinâmica por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotvaluefilter#condition)|Indica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|O limite inferior do intervalo para a `Between` condição de filtro.|
||[SelectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Indica se o filtro é para os N primeiros/últimos itens n, Top/Bottom N% ou soma superior/inferior N.|
||[soleira](/javascript/api/excel/excel.pivotvaluefilter#threshold)|O número de limite de "N" de itens, porcentagem ou soma a ser filtrado para uma condição de filtro Top/Bottom.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|O limite superior do intervalo para a `Between` condição de filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nome do "valor" escolhido no campo pelo qual filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getpivotrs (fullyContained?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtém uma coleção com escopo de tabelas dinâmicas que se sobrepõe ao intervalo.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Somente leitura.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Somente leitura.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Representa a categoria do formato de número de cada célula. Somente leitura.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha. Retorna um objeto Shape que representa a nova imagem.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|O estilo aplicado à segmentação de,.|
||[setStyle (Style: String \| pivotstyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Define o estilo aplicado à segmentação de,.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|O estilo aplicado à tabela.|
||[setStyle (Style: String \| pivotstyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Define o estilo aplicado à segmentação de,.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtém a ID da planilha que contém a tabela.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtém uma coleção de propriedades personalizadas no nível da planilha.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|O endereço do intervalo que concluiu o cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtém a chave da propriedade personalizada. Somente leitura.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtém o valor da propriedade personalizada. Somente leitura.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtém o número de propriedades personalizadas nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Lança se a propriedade personalizada não existe.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto NULL se a propriedade personalizada não existir.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtém a ID da planilha na qual o filtro é aplicado.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento foi acionado. Confira `Excel.RowHiddenChangeType` para obter detalhes.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
