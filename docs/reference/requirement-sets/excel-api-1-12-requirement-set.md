---
title: Excel conjunto de requisitos da API JavaScript 1.12
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-112"></a>Novidades na EXCEL JavaScript 1.12

O ExcelApi 1.12 aumentou o suporte a fórmulas em intervalos adicionando APIs para controlar matrizes dinâmicas e encontrando precedentes diretos de uma fórmula. Ele também adicionou controle API de filtros de tabela dinâmica. Melhorias também foram feitas nas áreas de recurso comentário, configurações de cultura e propriedades personalizadas.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Eventos de comentário](../../excel/excel-add-ins-comments.md#comment-events) | Adiciona eventos para adicionar, alterar e excluir à coleção de comentários.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Configurações de cultura [de data e hora](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Dá acesso a configurações culturais adicionais em torno da formatação de data e hora. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Aplicativo NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Precedentes diretos](../../excel/excel-add-ins-ranges-precedents.md) | Retorna intervalos usados para avaliar a fórmula de uma célula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Filtros pivôs | Aplica filtros orientados por valor aos campos de uma tabela dinâmica. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Vazamento de intervalo](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Permite que os complementos encontrem intervalos associados aos [resultados dinâmicos da matriz](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Propriedades personalizadas no nível da planilha](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Permite que as propriedades personalizadas sejam escopo para o nível da planilha, além de serem escopo para o nível da pasta de trabalho. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.12. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.12 ou anterior, consulte Excel APIs no conjunto de requisitos [1.12 ou anterior](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-textorientation-member)|Especifica o ângulo para o qual o texto é orientado para o título do eixo do gráfico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensionvalues-member(1))|Obtém os valores de uma única dimensão da série de gráficos.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#excel-excel-comment-contenttype-member)|Obtém o tipo de conteúdo do comentário.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-commentdetails-member)|Obtém `CommentDetail` a matriz que contém as IDs e as IDs de comentários de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-source-member)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-changetype-member)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-commentdetails-member)|Obter a `CommentDetail` matriz que contém as IDs de comentário e as IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-source-member)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member)|Ocorre quando os comentários são adicionados.|
||[onChanged](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member)|Ocorre quando comentários ou respostas em uma coleção de comentários são alterados, incluindo quando as respostas são excluídas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member)|Ocorre quando os comentários são excluídos na coleção de comentários.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-commentdetails-member)|Obtém `CommentDetail` a matriz que contém as IDs e as IDs de comentários de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-source-member)|Especifica a origem do evento.|
||[tipo](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o evento aconteceu.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-commentid-member)|Representa a ID do comentário.|
||[replyIds](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-replyids-member)|Representa as IDs das respostas relacionadas que pertencem ao comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-contenttype-member)|O tipo de conteúdo da resposta.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-datetimeformat-member)|Define o formato culturalmente apropriado de exibição de data e hora.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-dateseparator-member)|Obtém a cadeia de caracteres usada como separador de data.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longdatepattern-member)|Obtém a cadeia de caracteres de formato para um valor de data longa.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longtimepattern-member)|Obtém a cadeia de caracteres de formato por um valor de longo tempo.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-shortdatepattern-member)|Obtém a cadeia de caracteres de formato para um valor de data curta.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-timeseparator-member)|Obtém a cadeia de caracteres usada como separador de tempo.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparador](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-comparator-member)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-condition-member)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-exclusive-member)|If `true`, filter *exclui itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-lowerbound-member)|O limite inferior do intervalo para a condição `between` de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-upperbound-member)|O limite superior do intervalo para a condição `between` de filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-wholedays-member)|Para `equals`, `before`, `after`e condições `between` de filtro, indica se as comparações devem ser feitas como dias inteiros.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-applyfilter-member(1))|Define um ou mais dos PivotFilters atuais do campo e os aplica ao campo.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearallfilters-member(1))|Limpa todos os critérios de todos os filtros do campo.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearfilter-member(1))|Limpa todos os critérios existentes do filtro do campo do tipo determinado (se um estiver aplicado no momento).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-getfilters-member(1))|Obtém todos os filtros atualmente aplicados no campo.|
||[isFiltered(filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-isfiltered-member(1))|Verifica se há filtros aplicados no campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-datefilter-member)|O filtro de data aplicado no momento do PivotField.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-labelfilter-member)|O filtro de rótulo aplicado no momento do PivotField.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-manualfilter-member)|O filtro manual aplicado no momento do PivotField.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-valuefilter-member)|O filtro de valor aplicado no momento do PivotField.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparador](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-comparator-member)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-condition-member)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-exclusive-member)|If `true`, filter *exclui itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-lowerbound-member)|O limite inferior do intervalo para a condição `between` de filtro.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-substring-member)|A subdistragem usada para `beginsWith`, `endsWith`e condições `contains` de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-upperbound-member)|O limite superior do intervalo para a condição `between` de filtro.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#excel-excel-pivotmanualfilter-selecteditems-member)|Uma lista de itens selecionados para filtrar manualmente.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-allowmultiplefiltersperfield-member)|Especifica se a Tabela Dinâmica permite a aplicação de vários PivotFilters em um dado PivotField na tabela.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getcount-member(1))|Obtém o número de Tabelas Dinâmicas na coleção.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirst-member(1))|Obtém a primeira Tabela Dinâmica da coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitem-member(1))|Obtém uma Tabela Dinâmica por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitemornullobject-member(1))|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-comparator-member)|O comparador é o valor estático ao qual outros valores são comparados.|
||[condição](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-condition-member)|Especifica a condição do filtro, que define os critérios de filtragem necessários.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-exclusive-member)|If `true`, filter *exclui itens* que atendem aos critérios.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-lowerbound-member)|O limite inferior do intervalo para a condição `between` de filtro.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-selectiontype-member)|Especifica se o filtro é para os itens N superior/inferior, N por cento superior/inferior ou N superior/inferior.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-threshold-member)|O número limite "N" de itens, porcentagem ou soma a ser filtrado para uma condição de filtro superior/inferior.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-upperbound-member)|O limite superior do intervalo para a condição `between` de filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-value-member)|Nome do "valor" escolhido no campo pelo qual filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes diretos de uma célula na mesma planilha ou em várias planilhas.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getpivottables-member(1))|Obtém uma coleção com escopo de Tabelas Dinâmicas que se sobrepõem ao intervalo.|
||[getSpillParent()](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1))|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillparentornullobject-member(1))|Obtém o objeto range que contém a célula âncora para a célula que está sendo descarada.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1))|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorangeornullobject-member(1))|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora.|
||[hasSpill](/javascript/api/excel/excel.range#excel-excel-range-hasspill-member)|Representa se todas as células têm uma borda de despejo.|
||[numberFormatCategories](/javascript/api/excel/excel.range#excel-excel-range-numberformatcategories-member)|Representa a categoria do formato de número de cada célula.|
||[savedAsArray](/javascript/api/excel/excel.range#excel-excel-range-savedasarray-member)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getcount-member(1))|Obtém o número de `RangeAreas` objetos nesta coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getitemat-member(1))|Retorna o `RangeAreas` objeto com base na posição na coleção.|
||[items](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[addresses](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-addresses-member)|Retorna uma matriz de endereços no estilo A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-areas-member)|Retorna o `RangeAreasCollection` objeto.|
||[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasbysheet-member(1))|Retorna o `RangeAreas` objeto com base na ID da planilha ou no nome da coleção.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasornullobjectbysheet-member(1))|Retorna o `RangeAreas` objeto com base no nome da planilha ou na ID da coleção.|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-ranges-member)|Retorna intervalos que compõem esse objeto em um `RangeCollection` objeto.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-customproperties-member)|Obtém uma coleção de propriedades personalizadas no nível da planilha.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-delete-member(1))|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-key-member)|Obtém a chave da propriedade personalizada.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-value-member)|Obtém ou define o valor da propriedade personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-add-member(1))|Adiciona uma nova propriedade personalizada que mapeia para a chave fornecida.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getcount-member(1))|Obtém o número de propriedades personalizadas nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitem-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitemornullobject-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
