---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel
ms.date: 01/02/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7b751cfa550813fe3cb57f0d90d40e5632baf953
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969542"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Configurações de cultura | Obtém configurações culturais do sistema para a pasta de trabalho, como formatação de número. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Inserir pasta de trabalho](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insira uma pasta de trabalho em outra.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| [Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) a pasta de trabalho | Salve e feche a pasta de trabalho.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Excel atualmente em versão prévia. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs previamente lançadas), consulte [todas as APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fornece informações com base nas configurações de cultura do sistema atual. Isso inclui os nomes de cultura, a formatação de números e outras configurações dependentes de cultura.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos. Isso é baseado nas configurações locais do Excel.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos. Isso é baseado nas configurações locais do Excel.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Especifica se os separadores de sistema do Microsoft Excel estão habilitados.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimensão: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtém os valores de uma única dimensão da série de gráficos. Podem ser valores de categoria ou valores de dados, dependendo da dimensão especificada e de como os dados são mapeados para a série de gráficos.|
|[Comment](/javascript/api/excel/excel.comment)|[Obtido](/javascript/api/excel/excel.comment#resolved)|Obtém ou define o status do thread de comentários. O valor "true" significa que o thread de comentário está no estado resolvido.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[Obtido](/javascript/api/excel/excel.commentreply#resolved)|Obtém ou define o status de resposta de comentário. O valor "true" significa que a resposta de comentário está no estado resolvido.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtém o nome da cultura no formato languagecode2-Country/regioncode2 (por exemplo, "zh-CN" ou "en-US"). Isso é baseado nas configurações atuais do sistema.|
||[numberFormatInfo](/javascript/api/excel/excel.cultureinfo#numberformatinfo)|Define o formato culturalmente apropriado para exibir números. Isso é baseado nas configurações atuais de cultura do sistema.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos. Isso é baseado nas configurações atuais do sistema.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos. Isso é baseado nas configurações atuais do sistema.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias. A célula retornada é a interseção da linha e coluna fornecidas que contém os dados da hierarquia especificada. Esse método é o inverso de chamar getPivotItems e getDataHierarchy em uma célula específica.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Somente leitura.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Somente leitura.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha. Retorna um objeto Shape que representa a nova imagem.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Representa a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Representa a id da planilha que contém a tabela.|
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
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Representa a id da planilha na qual o filtro é aplicado.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento foi acionado. Confira `Excel.RowHiddenChangeType` para obter detalhes.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
