---
title: Conjunto de requisitos de API JavaScript do Excel 1,12
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,12.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a88c511e90fe48e1a9997d19cb4a2851cb718f6b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819837"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>O que há de novo na API JavaScript do Excel 1,12

O ExcelApi 1,12 aumentou o suporte para fórmulas em intervalos adicionando APIs para controlar matrizes dinâmicas e localizar os precedentes diretos de uma fórmula. Ele também adicionou controle de API de filtros de tabela dinâmica. Também foram feitas melhorias nas áreas de recurso comentário, configurações de cultura e propriedades personalizadas.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Eventos de comentários](../../excel/excel-add-ins-events.md) | Adiciona eventos para adicionar, alterar e excluir à coleção comment.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Configurações de cultura](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) de data e hora | Fornece acesso às configurações culturais adicionais em torno da formatação de data e hora. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo |
| Precedentes diretos | Retorna intervalos que são usados para avaliar a fórmula de uma célula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Filtros dinâmicos | Aplica filtros orientados a valores aos campos de uma tabela dinâmica. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Derramamento de intervalo](../../excel/excel-add-ins-ranges-advanced.md#handle-dynamic-arrays-and-spilling) | Permite que os suplementos encontrem intervalos associados aos resultados da [matriz dinâmica](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Propriedades personalizadas no nível da planilha](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Permite que as propriedades personalizadas sejam delimitadas ao nível da planilha, além de estarem no escopo da pasta de trabalho. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,12. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,12 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,12 ou anterior](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Especifica o ângulo no qual o texto é orientado para o título do eixo do gráfico. O valor deve ser um inteiro de-90 a 90 ou o inteiro 180 para texto orientado verticalmente.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimensão: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtém os valores de uma única dimensão da série de gráficos. Podem ser valores de categoria ou valores de dados, dependendo da dimensão especificada e de como os dados são mapeados para a série de gráficos.|
|[Comentário](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtém o tipo de conteúdo do comentário.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtenha a matriz CommentDetail que contém a ID de comentário e as IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Especifica a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento ocorreu.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtenha a matriz CommentDetail que contém a ID de comentário e as IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Especifica a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtém a ID da planilha na qual o evento ocorreu.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Ocorre quando os comentários são adicionados.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Ocorre quando comentários ou respostas em uma coleção de comentários são alterados, incluindo quando respostas são excluídas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Ocorre quando os comentários são excluídos na coleção comment.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtenha a matriz CommentDetail que contém a ID de comentário e as IDs de suas respostas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Especifica a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[tabela dinâmica](/javascript/api/excel/excel.pivotlayout#pivotstyle)|O estilo aplicado à tabela dinâmica.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Uma lista de itens selecionados a serem filtrados manualmente. Eles devem ser itens válidos e existentes do campo escolhido.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica se a tabela dinâmica permite o aplicativo de vários PivotFilters em um determinado campo PivotField na tabela.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtém a primeira tabela dinâmica na coleção. As tabelas dinâmicas da coleção são classificadas de cima para baixo e da esquerda para a direita, de forma que a tabela superior esquerda seja a primeira tabela dinâmica na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtém uma Tabela Dinâmica por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#comparator)|O comparador é o valor estático ao qual outros valores são comparados. O tipo de comparação é definido pela condição.|
||[condição](/javascript/api/excel/excel.pivotvaluefilter#condition)|Especifica a condição para o filtro, que define os critérios de filtragem necessários.|
||[Exclude](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Se true, Filter *excluirá* itens que atendem aos critérios. O padrão é false (filtrar para incluir itens que atendam aos critérios).|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|O limite inferior do intervalo para a condição de `Between` filtro.|
||[SelectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Especifica se o filtro é para os N itens superiores/inferiores, N superior/inferior% ou soma superior/inferior N.|
||[soleira](/javascript/api/excel/excel.pivotvaluefilter#threshold)|O número de limite de "N" de itens, porcentagem ou soma a ser filtrado para uma condição de filtro Top/Bottom.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|O limite superior do intervalo para a condição de `Between` filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nome do "valor" escolhido no campo pelo qual filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Retorna um objeto WorkbookRangeAreas que representa o intervalo que contém todos os precedentes diretos de uma célula na mesma planilha ou em várias planilhas.|
||[getpivotrs (fullyContained?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtém uma coleção com escopo de tabelas dinâmicas que se sobrepõe ao intervalo.|
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
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|O estilo aplicado à segmentação de,.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Retorna o `RangeAreas` objeto com base na ID ou no nome da planilha na coleção.|
||[getRangeAreasOrNullObjectBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Retorna o `RangeAreas` objeto com base no nome ou na ID da planilha na coleção. Se a planilha não existir, retornará um objeto null.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Retorna uma matriz de endereço em estilo a1. O valor de endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Planilha1! A1: B4, Planilha1! D1: D4 "). Somente leitura.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Retorna o `RangeAreasCollection` objeto. Cada `RangeAreas` objeto na coleção representa um ou mais intervalos de retângulo em uma planilha.|
||[variações](/javascript/api/excel/excel.workbookrangeareas#ranges)|Retorna intervalos que compõem este objeto em um `RangeCollection` objeto.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtém uma coleção de propriedades personalizadas no nível da planilha.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtém a chave da propriedade personalizada. As chaves de propriedades personalizadas não diferenciam maiúsculas de minúsculas. A chave está limitada a 255 caracteres (valores maiores causarão o erro "InvalidArgument".)|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtém ou define o valor da propriedade personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: String, value: String)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Adiciona uma nova propriedade personalizada que é mapeada para a chave fornecida. Isso substitui as propriedades personalizadas existentes por essa chave.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtém o número de propriedades personalizadas nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Lança se a propriedade personalizada não existe.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto NULL se a propriedade personalizada não existir.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
