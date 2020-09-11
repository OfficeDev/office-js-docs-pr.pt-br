---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel
ms.date: 06/29/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d1701ad393b96e33f0007bfcb5609c93c13608a2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430762"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Configurações de cultura](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) de data e hora | Fornece acesso às configurações culturais adicionais em torno da formatação de data e hora. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Inserir pasta de trabalho](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insira uma pasta de trabalho em outra.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Filtros dinâmicos | Aplica filtros orientados a valores aos campos de uma tabela dinâmica. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|Derramamento de intervalo | Permite que os suplementos encontrem intervalos associados aos resultados da [matriz dinâmica](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Excel atualmente em versão prévia. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs previamente lançadas), consulte [todas as APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimensão: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtém os valores de uma única dimensão da série de gráficos. Podem ser valores de categoria ou valores de dados, dependendo da dimensão especificada e de como os dados são mapeados para a série de gráficos.|
|[Comentário](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtém o tipo de conteúdo do comentário.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtém a `CommentDetail` matriz que contém a ID de comentário e IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Especifica a origem do evento. Confira `Excel.EventSource` para obter detalhes.|
||[tipo](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtém o tipo do evento. Confira `Excel.EventType` para obter detalhes.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento ocorreu.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtém a `CommentDetail` matriz que contém a ID de comentário e IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Especifica a origem do evento. Confira `Excel.EventSource` para obter detalhes.|
||[tipo](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtém o tipo do evento. Confira `Excel.EventType` para obter detalhes.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento ocorreu.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Ocorre quando os comentários são adicionados.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Ocorre quando comentários ou respostas em uma coleção de comentários são alterados, incluindo quando respostas são excluídas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Ocorre quando os comentários são excluídos na coleção comment.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtém a `CommentDetail` matriz que contém a ID de comentário e IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Especifica a origem do evento. Confira `Excel.EventSource` para obter detalhes.|
||[tipo](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtém o tipo do evento. Confira `Excel.EventType` para obter detalhes.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento ocorreu.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Representa a ID do comentário.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Representa as IDs das respostas relacionadas que pertencem ao comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|O tipo de conteúdo da resposta.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Define o formato culturalmente apropriado para exibir data e hora. Isso é baseado nas configurações atuais de cultura do sistema.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtém a cadeia de caracteres usada como o separador de data. Isso é baseado nas configurações atuais do sistema.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data longa. Isso é baseado nas configurações atuais do sistema.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtém a cadeia de caracteres de formato para um valor de tempo longo. Isso é baseado nas configurações atuais do sistema.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtém a cadeia de caracteres de formato para um valor de data abreviada. Isso é baseado nas configurações atuais do sistema.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtém a cadeia de caracteres usada como o separador de tempo. Isso é baseado nas configurações atuais do sistema.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Ativa este modo de exibição de planilha. Isso equivale a usar "mudar para" na interface do usuário do Excel.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Remove o modo de exibição de planilha da planilha.|
||[Duplicate (Name?: String)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Cria uma cópia deste modo de exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do modo de exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Cria um novo modo de exibição de planilha com o nome fornecido.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sai do modo de exibição de planilha ativo no momento.|
||[getactive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtém o modo de exibição de planilha atualmente ativo da planilha.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtém o número de modos de exibição de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtém um modo de exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtém um modo de exibição de planilha por seu índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparador](/javascript/api/excel/excel.pivotdatefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotdatefilter#condition)|Especifica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|O limite inferior do intervalo para a condição de `Between` filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|O limite superior do intervalo para a condição de `Between` filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Para `Equals` `Before` as condições de filtro,, e, `After` `Between` indica se as comparações devem ser feitas como dias inteiros.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filtro: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Define um ou mais PivotFilters atuais do campo e os aplica ao campo.|
||[clearAllFilters ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Limpa todos os critérios de todos os filtros de campo. Isso removerá qualquer filtragem ativa no campo.|
||[clearFilter (FilterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Limpa todos os critérios existentes do filtro do campo de determinado tipo (se houver algum aplicado no momento).|
||[GetFilters ()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtém todos os filtros aplicados no campo no momento.|
||[IsFiltered (FilterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Verifica se há filtros aplicados no campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|O filtro de data atualmente aplicado ao PivotField. Nulo se nenhum for aplicado.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|O filtro de rótulo do PivotField atualmente aplicado. Nulo se nenhum for aplicado.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|O filtro manual aplicado no momento do PivotField. Nulo se nenhum for aplicado.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|O filtro de valor atualmente aplicado ao PivotField. Nulo se nenhum for aplicado.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparador](/javascript/api/excel/excel.pivotlabelfilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotlabelfilter#condition)|Especifica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|O limite inferior do intervalo para a condição de filtro between.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|A subcadeia de caracteres usada para as `BeginsWith` `EndsWith` condições de filtro, e `Contains` .|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|O limite superior do intervalo para a condição de filtro entre.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias. A célula retornada é a interseção da linha e coluna fornecidas que contém os dados da hierarquia especificada. Esse método é o inverso de chamar getPivotItems e getDataHierarchy em uma célula específica.|
||[tabela dinâmica](/javascript/api/excel/excel.pivotlayout#pivotstyle)|O estilo aplicado à tabela dinâmica.|
||[setStyle (Style: String \| pivotstyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Define o estilo aplicado à tabela dinâmica.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Uma lista de itens selecionados a serem filtrados manualmente. Eles devem ser itens válidos e existentes do campo escolhido.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica se a tabela dinâmica permite o aplicativo de vários PivotFilters em um determinado campo PivotField na tabela.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotvaluefilter#condition)|Especifica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|O limite inferior do intervalo para a condição de `Between` filtro.|
||[SelectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Especifica se o filtro é para os N itens superiores/inferiores, N superior/inferior% ou soma superior/inferior N.|
||[soleira](/javascript/api/excel/excel.pivotvaluefilter#threshold)|O número de limite de "N" de itens, porcentagem ou soma a ser filtrado para uma condição de filtro Top/Bottom.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|O limite superior do intervalo para a condição de `Between` filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nome do "valor" escolhido no campo pelo qual filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes diretos de uma célula na mesma planilha ou em várias planilhas.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Retorna um objeto RangeAreas que representa as áreas mescladas neste intervalo. Observe que, se a contagem de áreas mescladas neste intervalo for maior que 512, a API falhará ao retornar o resultado.|
||[getprecedentes ()](/javascript/api/excel/excel.range#getprecedents--)|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou em várias planilhas.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Falha se aplicado a um intervalo com mais de uma célula.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Falha se aplicado a um intervalo com mais de uma célula.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Representa a categoria do formato de número de cada célula.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtém o número de objetos RangeAreas nesta coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Retorna o objeto RangeAreas com base na posição na coleção.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtém os itens filhos carregados nesta coleção.|
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
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica se o painel de lista de campos da tabela dinâmica é mostrado no nível da pasta de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Retorna o `RangeAreas` objeto com base na ID ou no nome da planilha na coleção.|
||[getRangeAreasOrNullObjectBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Retorna o `RangeAreas` objeto com base no nome ou na ID da planilha na coleção. Se a planilha não existir, retornará um objeto null.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Retorna uma matriz de endereço em estilo a1. O valor de endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Planilha1! A1: B4, Planilha1! D1: D4 "). Somente leitura.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Retorna o objeto RangeAreasCollection, cada RangeAreas na coleção representa um ou mais intervalos de retângulo em uma planilha.|
||[variações](/javascript/api/excel/excel.workbookrangeareas#ranges)|Retorna uma coleção de intervalos que inclui este objeto.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtém uma coleção de propriedades personalizadas no nível da planilha.|
||[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Retorna uma coleção de modos de exibição de planilha que estão presentes na planilha.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtém a chave da propriedade personalizada. As chaves de propriedades personalizadas não diferenciam maiúsculas de minúsculas. A chave está limitada a 255 caracteres (valores maiores causarão o erro "InvalidArgument".)|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtém ou define o valor da propriedade personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: String, value: String)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Adiciona uma nova propriedade personalizada que é mapeada para a chave fornecida. Isso substitui as propriedades personalizadas existentes por essa chave.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtém o número de propriedades personalizadas nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Lança se a propriedade personalizada não existe.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto NULL se a propriedade personalizada não existir.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtém a ID da planilha na qual o filtro é aplicado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
