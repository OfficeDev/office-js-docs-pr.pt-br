---
title: Excel Conjunto de requisitos da API JavaScript 1.4
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: be71d1e0c063bd3902bf57ba8f2024ae5a78ff1d
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671720"
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

As APIs incluem obter a entrada de configuração por meio da chave e adicionar o par de configuração `getItem()` `add()` key:value especificado à lista de trabalho.

## <a name="others"></a>Outros

* De definir o nome da coluna da tabela.
* Adicione uma coluna de tabela ao final da tabela.
* Adicione várias linhas a uma tabela de cada vez.
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Os [ \* métodos e propriedades OrNullObject](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties): Essa funcionalidade permite obter um objeto usando uma chave. Se o objeto não existir, a propriedade do objeto `isNullObject` retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto sem precisar lidar com ele por meio do tratamento de exceção. Um `*OrNullObject` método está disponível na maioria dos objetos da coleção.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.4. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.4 ou anterior, consulte Excel APIs no conjunto de requisitos [1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getCount__)|Obtém o número de associações da coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getItemOrNullObject_id_)|Obtém um objeto de associação pela ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getCount__)|Retorna o número de gráficos da planilha.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getItemOrNullObject_name_)|Obtém um gráfico usando o respectivo nome.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getCount__)|Retorna o número de pontos do gráfico da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getCount__)|Retorna o número de série da coleção.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Especifica o comentário associado a esse nome.|
||[delete()](/javascript/api/excel/excel.nameditem#delete__)|Exclui o nome fornecido.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getRangeOrNullObject__)|Retorna o objeto Range associado ao nome.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Especifica se o nome tem escopo para a pasta de trabalho ou para uma planilha específica.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Retorna a planilha em que o item nomeado tem escopo.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetOrNullObject)|Retorna a planilha à qual o item nomeado é escopo.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add_name__reference__comment_)|Adiciona um novo nome à coleção do escopo fornecido.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addFormulaLocal_name__formula__comment_)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getCount__)|Obtém o número de itens nomeados na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getItemOrNullObject_name_)|Obtém `NamedItem` um objeto usando seu nome.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getCount__)|Obtém o número de tabelas dinâmicas na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getItemOrNullObject_name_)|Obtém uma Tabela Dinâmica por nome.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersectionOrNullObject_anotherRange_)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getUsedRangeOrNullObject_valuesOnly_)|Retorna o intervalo usado do objeto de intervalo determinado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getCount__)|Obtém o número `RangeView` de objetos na coleção.|
|[Configuração](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete__)|Exclui a configuração.|
||[key](/javascript/api/excel/excel.setting#key)|A chave que representa a ID da configuração.|
||[value](/javascript/api/excel/excel.setting#value)|Representa o valor armazenado para esta configuração.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date Array \| <any> \| any)](/javascript/api/excel/excel.settingcollection#add_key__value_)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getCount__)|Obtém o número de configurações na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getItem_key_)|Obtém uma entrada de configuração por meio da chave.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getItemOrNullObject_key_)|Obtém uma entrada de configuração por meio da chave.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onSettingsChanged)|Ocorre quando as configurações no documento são alteradas.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[configurações](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtém `Setting` o objeto que representa a associação que gerou o evento alterado de configurações|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getCount__)|Obtém o número de tabelas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getItemOrNullObject_key_)|Obtém uma tabela pelo nome ou ID.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getCount__)|Obtém a quantidade de colunas na tabela.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItemOrNullObject_key_)|Obtém um objeto de coluna por nome ou ID.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getCount__)|Obtém a quantidade de linhas na tabela.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[configurações](/javascript/api/excel/excel.workbook#settings)|Representa uma coleção de configurações associadas à workbook.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getUsedRangeOrNullObject_valuesOnly_)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas.|
||[names](/javascript/api/excel/excel.worksheet#names)|Coleção de nomes com escopo para a planilha atual.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getCount_visibleOnly_)|Obtém o número de planilhas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getItemOrNullObject_key_)|Obtém um objeto de planilha usando o nome ou ID dele.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
