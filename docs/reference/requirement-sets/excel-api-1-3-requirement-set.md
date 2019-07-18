---
title: Conjunto de requisitos de API JavaScript do Excel 1,3
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,3
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4698b0fad3122c8ecf52117c35d4928305d812fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771992"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

ExcelApi 1,3 adicionado suporte para associação de dados e acesso básico de tabela dinâmica.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Associação](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Especifica a associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: String \| de intervalo, BindingType: "Range \| " "Table \| " "text", ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Adiciona uma nova associação a um intervalo específico.|
||[Add (Range: String \| de intervalo, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Adiciona uma nova associação a um intervalo específico.|
||[addFromNamedItem (Name: String, BindingType: "Range" \| "Table" \| "text", ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|
||[addFromNamedItem (Name: String, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|
||[addFromSelection (BindingType: "Range" \| "Table" \| "text", ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Adiciona uma nova associação com base na seleção atual.|
||[addFromSelection (BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Adiciona uma nova associação com base na seleção atual.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nome da Tabela Dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|A planilha que contém a Tabela Dinâmica atual.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Atualiza a Tabela Dinâmica.|
||[Set (Propriedades: Excel. PivotTable)](/javascript/api/excel/excel.pivottable#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotTableUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivottable#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Atualiza todas as tabelas dinâmicas da coleção.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|Para cada ITEM na coleção: nome da tabela dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|Para cada ITEM na coleção: a planilha que contém a tabela dinâmica atual.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|Nome da Tabela Dinâmica.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|Nome da Tabela Dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|A planilha que contém a Tabela Dinâmica atual.|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|Nome da Tabela Dinâmica.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|Representa as linhas visíveis do intervalo atual.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[fórmulas](/javascript/api/excel/excel.rangeview#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Representa a fórmula em notação no estilo L1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtém o intervalo pai associado à RangeView atual.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Representa o código de formato de número do Excel para determinada célula.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Representa os endereços de célula da RangeView. Somente leitura.|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|Retorna o número de colunas visíveis. Somente leitura.|
||[index](/javascript/api/excel/excel.rangeview#index)|Retorna um valor que representa o índice da RangeView. Somente leitura.|
||[Validação](/javascript/api/excel/excel.rangeview#rowcount)|Retorna o número de linhas visíveis. Somente leitura.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Representa uma coleção de exibições de tabelas associadas ao intervalo. Somente leitura.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Representa o tipo de dados de cada célula. Somente leitura.|
||[Set (Propriedades: Excel. RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. RangeViewUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.rangeview#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[values](/javascript/api/excel/excel.rangeview#values)|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtém uma linha RangeView por meio de seu índice. Indexado com zero.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeViewCollectionLoadOptions](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|Para cada ITEM na coleção: representa os endereços de célula do RangeView. Somente leitura.|
||[columnCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|Para cada ITEM na coleção: retorna o número de colunas visíveis. Somente leitura.|
||[fórmulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|Para cada ITEM na coleção: representa a fórmula em notação de estilo a1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|Para cada ITEM na coleção: representa a fórmula em notação de estilo a1, no idioma do usuário e na localidade de formatação de números.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|Para cada ITEM na coleção: representa a fórmula em notação de estilo L1C1.|
||[index](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|Para cada ITEM na coleção: retorna um valor que representa o índice do RangeView. Somente leitura.|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|Para cada ITEM na coleção: representa o código de formato de número do Excel para a célula especificada.|
||[Validação](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|Para cada ITEM na coleção: retorna o número de linhas visíveis. Somente leitura.|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|Para cada ITEM na coleção: valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|Para cada ITEM na coleção: representa o tipo de dados de cada célula. Somente leitura.|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|Para cada ITEM na coleção: representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeViewData](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|Representa os endereços de célula da RangeView. Somente leitura.|
||[columnCount](/javascript/api/excel/excel.rangeviewdata#columncount)|Retorna o número de colunas visíveis. Somente leitura.|
||[fórmulas](/javascript/api/excel/excel.rangeviewdata#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|Representa a fórmula em notação no estilo L1C1.|
||[index](/javascript/api/excel/excel.rangeviewdata#index)|Retorna um valor que representa o índice da RangeView. Somente leitura.|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|Representa o código de formato de número do Excel para determinada célula.|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|Retorna o número de linhas visíveis. Somente leitura.|
||[rows](/javascript/api/excel/excel.rangeviewdata#rows)|Representa uma coleção de exibições de tabelas associadas ao intervalo. Somente leitura.|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|Representa o tipo de dados de cada célula. Somente leitura.|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeViewLoadOptions](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|Representa os endereços de célula da RangeView. Somente leitura.|
||[columnCount](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|Retorna o número de colunas visíveis. Somente leitura.|
||[fórmulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|Representa a fórmula em notação no estilo L1C1.|
||[index](/javascript/api/excel/excel.rangeviewloadoptions#index)|Retorna um valor que representa o índice da RangeView. Somente leitura.|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|Representa o código de formato de número do Excel para determinada célula.|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|Retorna o número de linhas visíveis. Somente leitura.|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|Representa o tipo de dados de cada célula. Somente leitura.|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeViewUpdateData](/javascript/api/excel/excel.rangeviewupdatedata)|[fórmulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|Representa a fórmula em notação no estilo L1C1.|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|Representa o código de formato de número do Excel para determinada célula.|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indica se a primeira coluna contém uma formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indica se a última coluna contém uma formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|Para cada ITEM na coleção: indica se a primeira coluna contém formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|Para cada ITEM na coleção: indica se a última coluna contém formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|Para cada ITEM na coleção: indica se as colunas mostram a formatação em tiras nas quais as colunas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|Para cada ITEM na coleção: indica se as linhas mostram a formatação em tiras nas quais as linhas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|Para cada ITEM na coleção: indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho de coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|Indica se a primeira coluna contém uma formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|Indica se a última coluna contém uma formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|Indica se a primeira coluna contém uma formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|Indica se a última coluna contém uma formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|Indica se a primeira coluna contém uma formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|Indica se a última coluna contém uma formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[pivotTables](/javascript/api/excel/excel.workbookdata#pivottables)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[pivotTables](/javascript/api/excel/excel.worksheetdata#pivottables)|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
