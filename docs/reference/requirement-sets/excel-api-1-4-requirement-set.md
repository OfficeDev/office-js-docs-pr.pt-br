---
title: Excel conjunto de requisitos da API JavaScript 1.4
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bcdbd044c5de562b7c2cc2bc9971af31179f8a9b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746541"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

## <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` - Itens com escopo de planilha ou pasta de trabalho.
* `worksheet` - Retorna a planilha na qual o item nomeado tem escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)` - Adiciona um novo nome à coleção do escopo determinado.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Adiciona um novo nome à coleção do escopo determinado usando a localidade do usuário para a fórmula.

## <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs incluem `getItem()` obter a entrada de configuração por `add()` meio da chave e adicionar o par de configuração key:value especificado à lista de trabalho.

## <a name="others"></a>Outros

* De definir o nome da coluna da tabela.
* Adicione uma coluna de tabela ao final da tabela.
* Adicione várias linhas a uma tabela de cada vez.
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Os [\*métodos e propriedades OrNullObject](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties): Essa funcionalidade permite obter um objeto usando uma chave. Se o objeto não existir, a propriedade do `isNullObject` objeto retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto sem precisar lidar com ele por meio do tratamento de exceção. Um `*OrNullObject` método está disponível na maioria dos objetos da coleção.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.4. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.4 ou anterior, consulte Excel APIs no conjunto de requisitos [1.4 ou anterior](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|Obtém o número de associações da coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|Obtém um objeto de associação pela ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|Retorna o número de gráficos da planilha.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|Obtém um gráfico usando o respectivo nome.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|Retorna o número de pontos do gráfico da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|Retorna o número de série da coleção.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|Especifica o comentário associado a esse nome.|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|Exclui o nome fornecido.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|Retorna o objeto Range associado ao nome.|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|Especifica se o nome tem escopo para a pasta de trabalho ou para uma planilha específica.|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|Retorna a planilha em que o item nomeado tem escopo.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|Retorna a planilha à qual o item nomeado é escopo.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|Adiciona um novo nome à coleção do escopo fornecido.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|Obtém o número de itens nomeados na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|Obtém um `NamedItem` objeto usando seu nome.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|Obtém o número de tabelas dinâmicas na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|Obtém uma Tabela Dinâmica por nome.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersectionornullobject-member(1))|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getusedrangeornullobject-member(1))|Retorna o intervalo usado do objeto de intervalo determinado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|Obtém o número de `RangeView` objetos na coleção.|
|[Configuração](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|Exclui a configuração.|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|A chave que representa a ID da configuração.|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|Representa o valor armazenado para esta configuração.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array \| any)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|Obtém o número de configurações na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|Obtém uma entrada de configuração por meio da chave.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|Obtém uma entrada de configuração por meio da chave.|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|Ocorre quando as configurações no documento são alteradas.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[configurações](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|Obtém o `Setting` objeto que representa a associação que gerou o evento alterado de configurações|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|Obtém o número de tabelas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|Obtém uma tabela pelo nome ou ID.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|Obtém a quantidade de colunas na tabela.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|Obtém um objeto de coluna por nome ou ID.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|Obtém a quantidade de linhas na tabela.|
|[Workbook](/javascript/api/excel/excel.workbook)|[configurações](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|Representa uma coleção de configurações associadas à workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getusedrangeornullobject-member(1))|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas.|
||[names](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|Coleção de nomes com escopo para a planilha atual.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|Obtém o número de planilhas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|Obtém um objeto de planilha usando o nome ou ID dele.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
