---
title: Excel Conjunto de requisitos da API JavaScript 1.14
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.14.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9cdf22d35125607237b724c88da2083ae78a9940
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681504"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Novidades na API JavaScript 1.14 Excel JavaScript

O ExcelApi 1.14 adicionou objetos para controlar o recurso de tabela de dados de um gráfico, um método para localizar todas as células precedentes de uma fórmula e eventos de proteção de planilha para rastrear alterações no estado de proteção de uma planilha. Ele também adicionou [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) vários métodos para objetos como , e para melhorar o tratamento de `CommentCollection` `ShapeCollection` `StyleCollection` erros.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Tabelas de dados de gráfico | Controlar a aparência, a formatação e a visibilidade das tabelas de dados nos gráficos. | [Chart,](/javascript/api/excel/excel.chart) [ChartDataTable,](/javascript/api/excel/excel.chartdatatable) [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Precedentes de fórmula | Retorne todas as células precedentes de uma fórmula. | [Range](/javascript/api/excel/excel.range) |
| Consultas | Recupere atributos de Consulta do Power, como nome, data de atualização e contagem de consultas. | [Consulta](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| Eventos de proteção de planilha | Acompanhe as alterações no estado de proteção de uma planilha e na origem dessas alterações. | [WorksheetProtectionChangedEventArgs,](/javascript/api/excel/excel.worksheetprotectionchangedeventargs) [Planilha,](/javascript/api/excel/excel.worksheet) [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.14. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.14 ou anterior, consulte Excel APIs no conjunto de requisitos [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Limpa os critérios de filtro de coluna do AutoFilter.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Representa a direção (como para cima ou para a esquerda) que as células restantes serão deslocadas quando uma célula ou células são excluídas.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Representa a direção (como para baixo ou para a direita) que as células existentes mudarão quando uma nova célula ou células são inseridas.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtém a tabela de dados no gráfico.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtém a tabela de dados no gráfico.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Representa o formato de uma tabela de dados do gráfico, que inclui o formato de preenchimento, fonte e borda.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Especifica se a borda horizontal da tabela de dados deve ser exibida.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Especifica se a chave de legenda da tabela de dados deve ser apresentada.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Especifica se a borda de contorno da tabela de dados deve ser exibida.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Especifica se a borda vertical da tabela de dados deve ser exibida.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Especifica se é preciso mostrar a tabela de dados do gráfico.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[borda](/javascript/api/excel/excel.chartdatatableformat#border)|Representa o formato de borda da tabela de dados do gráfico, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) do objeto atual.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtém um comentário da coleção com base em seu ID.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Retorna uma resposta de comentário identificada pela respectiva ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Retorna um formato condicional identificado por sua ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Obtém uma forma usando seu nome ou ID.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtém a mensagem de erro de consulta de quando a consulta foi atualizada pela última vez.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtém a consulta carregada para o tipo de objeto.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Especifica se a consulta foi carregada para o modelo de dados.|
||[name](/javascript/api/excel/excel.query#name)|Obtém o nome da consulta.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtém a data e a hora em que a consulta foi atualizada pela última vez.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtém o número de linhas que foram carregadas quando a consulta foi atualizada pela última vez.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtém o número de consultas na guia de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtém uma consulta da coleção com base em seu nome.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Retorna um objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtém uma forma usando seu nome ou ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Obtém um estilo por nome.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Obtém uma tabela pelo nome ou ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[consultas](/javascript/api/excel/excel.workbook#queries)|Retorna uma coleção de consultas do Power Query que fazem parte da workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Retorna um valor que representa essa planilha que pode ser lido por Open Office XML.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Representa uma alteração na direção em que as células de uma planilha serão deslocadas quando uma célula ou células são excluídas ou inseridas.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Representa a origem do gatilho do evento.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtém o status de proteção atual da planilha.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtém a ID da planilha na qual o status da proteção é alterado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
