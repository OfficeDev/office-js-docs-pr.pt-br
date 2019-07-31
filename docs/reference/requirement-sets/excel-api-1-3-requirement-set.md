---
title: Conjunto de requisitos de API JavaScript do Excel 1,3
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,3
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 684b802a32e58591d43d46a37ecc8b53395b652c
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940756"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

ExcelApi 1,3 adicionado suporte para associação de dados e acesso básico de tabela dinâmica.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Associação](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Especifica a associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: String \| de intervalo, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Adiciona uma nova associação a um intervalo específico.|
||[addFromNamedItem (Name: String, BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|
||[addFromSelection (BindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Adiciona uma nova associação com base na seleção atual.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nome da Tabela Dinâmica.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|A planilha que contém a Tabela Dinâmica atual.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Atualiza a Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtém uma Tabela Dinâmica por nome.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Atualiza todas as tabelas dinâmicas da coleção.|
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
||[values](/javascript/api/excel/excel.rangeview#values)|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtém uma linha RangeView por meio de seu índice. Indexado com zero.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indica se a primeira coluna contém uma formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indica se a última coluna contém uma formatação especial.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
