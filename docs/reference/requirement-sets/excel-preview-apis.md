---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9ddc1405d4bc13087780e8950b36d9b3b4b04069
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819788"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados ao Excel a partir de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Exibições de planilha nomeadas | Fornece controle programático de modos de exibição de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Excel atualmente em versão prévia. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs previamente lançadas), consulte [todas as APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[DataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|O nome do provedor de dados para o tipo de dados vinculados. Isso pode ser alterado quando as informações são recuperadas do serviço.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|A data e a hora da zona de tempo local desde que a pasta de trabalho foi aberta quando o tipo de dados vinculados foi atualizado pela última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|O nome do tipo de dados vinculados. Isso pode ser alterado quando as informações são recuperadas do serviço.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|A frequência, em segundos, em que o tipo de dados vinculado é atualizado, se `refreshMode` estiver definido como "periódico".|
||[RefreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|O mecanismo pelo qual os dados para o tipo de dados vinculados são recuperados.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|A identificação exclusiva do tipo de dados vinculados.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Retorna uma matriz com todos os modos de atualização compatíveis com o tipo de dados vinculados. O conteúdo da matriz pode ser alterado quando as informações são recuperadas do serviço.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Faz uma solicitação para atualizar o tipo de dados vinculados. Se o serviço estiver ocupado ou temporariamente inacessível, a solicitação não será atendida.|
||[requestSetRefreshMode (RefreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculados.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|A identificação exclusiva do novo tipo de dados vinculados.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtém um tipo de dados vinculado por ID de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtém um tipo de dados vinculado por seu índice na coleção.|
||[getItemOrNullObject (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtém um tipo de dados vinculado por ID. Se o tipo de dados vinculado não existir, um objeto com sua `isNullObject` propriedade definida como `true` . Para obter mais informações, consulte {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | * Métodos e propriedades do OrNullObject}.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|A descrição de texto alt da tabela dinâmica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|O título do texto alt da tabela dinâmica.|
||[displayBlankLineAfterEachItem (exibição: Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Define se deve ou não exibir uma linha em branco após cada item. Isso é definido no nível global para a tabela dinâmica e aplicado a campos PivotFields individuais.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|O texto que é preenchido automaticamente em qualquer célula vazia da tabela dinâmica se `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica se as células vazias da tabela dinâmica devem ser preenchidas com o `emptyCellText` . False por padrão.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias. A célula retornada é a interseção da linha e coluna fornecidas que contém os dados da hierarquia especificada. Esse método é o inverso de chamar getPivotItems e getDataHierarchy em uma célula específica.|
||[repeatAllItemLabels (repeatLabels: Boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Define a configuração "repetir todos os rótulos de item" em todos os campos da tabela dinâmica.|
||[setStyle (Style: String \| pivotstyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Define o estilo aplicado à tabela dinâmica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica se a tabela dinâmica exibe cabeçalhos de campos (legendas de campos e suspensas de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica se a tabela dinâmica é atualizada quando a pasta de trabalho é aberta. Corresponde à configuração "atualizar ao carregar" na interface do usuário.|
|[Range](/javascript/api/excel/excel.range)|[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Retorna um `RangeAreas` objeto que representa as áreas mescladas neste intervalo. Observe que, se a contagem de áreas mescladas neste intervalo for maior que 512, a API falhará ao retornar o resultado.|
||[getprecedentes ()](/javascript/api/excel/excel.range#getprecedents--)|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou em várias planilhas.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[RefreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|A identificação exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica se a solicitação para atualizar foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|A identificação exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[alerta](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha. Retorna um objeto Shape que representa a nova imagem.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[setStyle (Style: String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Define o estilo aplicado à segmentação de,.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|O estilo aplicado à tabela.|
||[setStyle (Style: String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Define o estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtém a ID da planilha que contém a tabela.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da pasta de trabalho.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica se o painel de lista de campos da tabela dinâmica é mostrado no nível da pasta de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Retorna uma coleção de modos de exibição de planilha que estão presentes na planilha.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtém a ID da planilha na qual o filtro é aplicado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
