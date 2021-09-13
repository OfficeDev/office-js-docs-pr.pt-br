---
title: Excel Conjunto de requisitos da API JavaScript 1.13
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 677cb9eceba3d25a7f2d0f8677ab63942852a784
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151748"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Novidades na API JavaScript 1.13 Excel JavaScript

O ExcelApi 1.13 adicionou um método para inserir planilhas em uma pasta de trabalho de uma cadeia de caracteres codificada com Base64 e um evento para detectar a ativação de pasta de trabalho. Ele também aumentou o suporte a fórmulas em intervalos adicionando APIs para rastrear alterações nas fórmulas e localizar células dependentes diretas de uma fórmula. Além disso, ele expandiu o suporte à Tabela Dinâmica adicionando APIs pivotLayout para o gerenciamento de células de alt text, style e empty.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Fórmula de eventos alterados](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Acompanhe as alterações nas fórmulas, incluindo a origem e o tipo de evento que causou uma alteração. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| [Dependentes da fórmula](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Localize as células dependentes diretas de uma fórmula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| [Inserir planilhas](../../excel//excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insira planilhas de outra pasta de trabalho na pasta de trabalho atual como uma cadeia de caracteres codificada com Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| [PivotLayout de tabela dinâmica](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Uma expansão da classe PivotLayout, incluindo novo suporte para alt text e gerenciamento de células vazias. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.13. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.13 ou anterior, consulte Excel APIs no conjunto de requisitos [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#cellAddress)|O endereço da célula que contém a fórmula alterada.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousFormula)|Representa a fórmula anterior, antes de ser alterada.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positionType)|A posição de inserção, na pasta de trabalho atual, das novas planilhas.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeTo)|A planilha na pasta de trabalho atual que é referenciada para o `WorksheetPositionType` parâmetro.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)|Os nomes de planilhas individuais a inserir.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#altTextDescription)|A descrição de texto alt da Tabela Dinâmica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#altTextTitle)|O título de texto alt da Tabela Dinâmica.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayBlankLineAfterEachItem_display_)|Define se uma linha em branco deve ou não ser exibida após cada item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptyCellText)|O texto que é preenchido automaticamente em qualquer célula vazia na Tabela Dinâmica se `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillEmptyCells)|Especifica se células vazias na Tabela Dinâmica devem ser preenchidas com `emptyCellText` o .|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatAllItemLabels_repeatLabels_)|Define a configuração "repetir todos os rótulos de item" em todos os campos da Tabela Dinâmica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showFieldHeaders)|Especifica se a Tabela Dinâmica exibe os headers de campo (legendas de campo e drop-downs de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshOnOpen)|Especifica se a Tabela Dinâmica é atualizada quando a workbook é aberta.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getDirectDependents__)|Retorna um objeto que representa o intervalo que contém todos os dependentes diretos de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#getExtendedRange_direction__activeCell_)|Retorna um objeto range que inclui o intervalo atual e até a borda do intervalo, com base na direção fornecida.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getMergedAreasOrNullObject__)|Retorna um objeto RangeAreas que representa as áreas mescladas nesse intervalo.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_)|Retorna um objeto range que é a célula de borda da região de dados que corresponde à direção fornecida.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize_newRange_)|Resize a tabela para o novo intervalo.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_)|Insere as planilhas especificadas de uma pasta de trabalho de origem na pasta de trabalho atual.|
||[onActivated](/javascript/api/excel/excel.workbook#onActivated)|Ocorre quando a guia de trabalho é ativada.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[tipo](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtém o tipo do evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|Ocorre quando uma ou mais fórmulas são alteradas nesta planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)|Ocorre quando uma ou mais fórmulas são alteradas em qualquer planilha dessa coleção.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails)|Obtém uma matriz `FormulaChangedEventDetail` de objetos, que contém os detalhes sobre todas as fórmulas alteradas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetId)|Obtém a ID da planilha na qual a fórmula foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
