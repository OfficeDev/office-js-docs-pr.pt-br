---
title: Excel Conjunto de requisitos da API JavaScript 1.3
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.3.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: c7c39a341f635e3355014f75e32c1501124f99d9
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367457"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

O ExcelApi 1.3 adicionou suporte para vinculação de dados e acesso básico à tabela dinâmica.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no Excel de requisitos da API JavaScript 1.3. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.3 ou anterior, consulte Excel APIs no conjunto de requisitos [1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Associação](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|Especifica a associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|Adiciona uma nova associação a um intervalo específico.|
||[addFromNamedItem(name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|
||[addFromSelection(bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|Adiciona uma nova associação com base na seleção atual.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nome da Tabela Dinâmica.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|Atualiza a Tabela Dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|A planilha que contém a Tabela Dinâmica atual.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|Atualiza todas as tabelas dinâmicas da coleção.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getVisibleView__)|Representa as linhas visíveis do intervalo atual.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|Representa os endereços de célula do `RangeView` .|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|O número de colunas visíveis.|
||[fórmulas](/javascript/api/excel/excel.rangeview#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|Representa a fórmula em notação no estilo L1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|Obtém o intervalo pai associado ao `RangeView` atual .|
||[índice](/javascript/api/excel/excel.rangeview#index)|Retorna um valor que representa o índice do `RangeView` .|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|Representa o código de formato de número do Excel para determinada célula.|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|O número de linhas visíveis.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Representa uma coleção de exibições de tabelas associadas ao intervalo.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valores de texto do intervalo especificado.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|Representa o tipo de dados de cada célula.|
||[values](/javascript/api/excel/excel.rangeview#values)|Representa os valores brutos da exibição do intervalo especificado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|Obtém `RangeView` uma linha por meio de seu índice.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|Especifica se a primeira coluna contém formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|Especifica se a última coluna contém formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|Especifica se as colunas mostram formatação em faixa na qual as colunas ímpares são realçadas de forma diferente de outras, para facilitar a leitura da tabela.|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|Especifica se as linhas mostram formatação em faixa na qual linhas ímpares são realçadas de forma diferente de outras, para facilitar a leitura da tabela.|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|Especifica se os botões de filtro estão visíveis na parte superior de cada header de coluna.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivotTables)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivotTables)|Coleção de Tabelas Dinâmicas que fazem parte da planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
