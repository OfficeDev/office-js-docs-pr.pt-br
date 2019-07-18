---
title: Conjunto de requisitos de API JavaScript do Excel 1,44
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,4
ms.date: 07/15/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c0cd380a71c98ab63aa955ec0ff2ed005065577c
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771978"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

## <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope`– Itens com escopo de planilha ou pasta de trabalho.
* `worksheet`-Retorna a planilha na qual o item nomeado tem escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)`– Adiciona um novo nome à coleção de um determinado escopo.
* `addFormulaLocal(name: string, formula: string, comment: string)`– Adiciona um novo nome à coleção do escopo fornecido usando a localidade do usuário para a fórmula.

## <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs `getItem()` incluem para obter a entrada de configuração através `add()` da chave e para adicionar o par de definição de valor-chave especificado: Value à pasta de trabalho.

## <a name="others"></a>Outros

* Definir o nome da coluna da tabela.
* Adicione uma coluna de tabela ao final da tabela.
* Adicionar várias linhas a uma tabela de cada vez.
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* A [função de obter item ou objeto nulo](../../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods): essa funcionalidade permite obter o objeto usando uma chave. Se o objeto não existir, a propriedade do `isNullObject` objeto retornado será true. Isso permite que os desenvolvedores verifiquem se um objeto existe ou não sem precisar tratá-lo por meio da manipulação de exceção. O `*OrNullObject` método está disponível na maioria dos objetos Collection.

```javascript
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Obtém o número de associações da coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Obtém um objeto binding pela ID. Se o objeto binding não existir, retornará um objeto null.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Retorna o número de gráficos da planilha.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Retorna o número de pontos do gráfico da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Retorna o número de série da coleção.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[Retire](/javascript/api/excel/excel.nameditem#comment)|Representa o comentário associado a esse nome.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Exclui o nome fornecido.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Retorna o objeto Range associado ao nome. Retornará um objeto null se o tipo do item nomeado não for um intervalo.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Os valores possíveis são: planilha, pasta de trabalho. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Retorna a planilha em que o item nomeado tem escopo. Gera um erro se o item estiver no escopo da pasta de trabalho.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (Name: String, Reference: cadeia \| de caracteres de intervalo, comentário?: cadeia de caracteres)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Adiciona um novo nome à coleção do escopo fornecido.|
||[addFormulaLocal (Name: String, formula: String, comment?: String)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Obtém o número de itens nomeados na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Obtém um objeto NamedItem usando seu nome. Se o objeto getNamedItem não existir, retornará um objeto null.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[Retire](/javascript/api/excel/excel.nameditemcollectionloadoptions#comment)|Para cada ITEM na coleção: representa o comentário associado a esse nome.|
||[scope](/javascript/api/excel/excel.nameditemcollectionloadoptions#scope)|Para cada ITEM na coleção: indica se o nome tem o escopo para a pasta de trabalho ou para uma planilha específica. Os valores possíveis são: planilha, pasta de trabalho. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.nameditemcollectionloadoptions#worksheet)|Para cada ITEM na coleção: retorna a planilha na qual o item nomeado tem escopo. Gera um erro se o item estiver no escopo da pasta de trabalho.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditemcollectionloadoptions#worksheetornullobject)|Para cada ITEM na coleção: retorna a planilha na qual o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[Retire](/javascript/api/excel/excel.nameditemdata#comment)|Representa o comentário associado a esse nome.|
||[scope](/javascript/api/excel/excel.nameditemdata#scope)|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Os valores possíveis são: planilha, pasta de trabalho. Somente leitura.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[Retire](/javascript/api/excel/excel.nameditemloadoptions#comment)|Representa o comentário associado a esse nome.|
||[scope](/javascript/api/excel/excel.nameditemloadoptions#scope)|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Os valores possíveis são: planilha, pasta de trabalho. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.nameditemloadoptions#worksheet)|Retorna a planilha em que o item nomeado tem escopo. Gera um erro se o item estiver no escopo da pasta de trabalho.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditemloadoptions#worksheetornullobject)|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[Retire](/javascript/api/excel/excel.nameditemupdatedata#comment)|Representa o comentário associado a esse nome.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: cadeia \| de caracteres de intervalo)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Retorna o intervalo usado do objeto range determinado. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Obtém o número de objetos RangeView na coleção.|
|[Configuração](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Exclui a configuração.|
||[](/javascript/api/excel/excel.setting#datejsonprefix)||
||[](/javascript/api/excel/excel.setting#datejsonsuffix)||
||[](/javascript/api/excel/excel.setting#replacestringdatewithdate)||
||[key](/javascript/api/excel/excel.setting#key)|Retorna a chave que representa a id da configuração. Somente leitura.|
||[Set (Propriedades: Excel. setting)](/javascript/api/excel/excel.setting#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. SettingUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.setting#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[value](/javascript/api/excel/excel.setting#value)|Representa o valor armazenado para esta configuração.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (Key: String, value: String \| número \| Boolean \| data \| array<any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Obtém o número de Configurações na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Obtém uma entrada de configuração por meio da tecla.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Obtém uma entrada de configuração por meio da tecla. Se a Configuração não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Ocorre quando as Configurações no documento são alteradas.|
|[SettingCollectionLoadOptions](/javascript/api/excel/excel.settingcollectionloadoptions)|[$all](/javascript/api/excel/excel.settingcollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.settingcollectionloadoptions#key)|Para cada ITEM na coleção: retorna a chave que representa a ID da configuração. Somente leitura.|
||[value](/javascript/api/excel/excel.settingcollectionloadoptions#value)|Para cada ITEM na coleção: representa o valor armazenado para esta configuração.|
|[SettingData](/javascript/api/excel/excel.settingdata)|[key](/javascript/api/excel/excel.settingdata#key)|Retorna a chave que representa a id da configuração. Somente leitura.|
||[value](/javascript/api/excel/excel.settingdata#value)|Representa o valor armazenado para esta configuração.|
|[SettingLoadOptions](/javascript/api/excel/excel.settingloadoptions)|[$all](/javascript/api/excel/excel.settingloadoptions#$all)||
||[key](/javascript/api/excel/excel.settingloadoptions#key)|Retorna a chave que representa a id da configuração. Somente leitura.|
||[value](/javascript/api/excel/excel.settingloadoptions#value)|Representa o valor armazenado para esta configuração.|
|[SettingUpdateData](/javascript/api/excel/excel.settingupdatedata)|[value](/javascript/api/excel/excel.settingupdatedata#value)|Representa o valor armazenado para esta configuração.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[configurações](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, retornará um objeto null.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Obtém a quantidade de colunas na tabela.|
||[getItemOrNullObject (Key: String \| de número)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Obtém um objeto column por nome ou ID. Se a coluna não existir, retornará um objeto null.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Obtém a quantidade de linhas na tabela.|
|[Workbook](/javascript/api/excel/excel.workbook)|[configurações](/javascript/api/excel/excel.workbook#settings)|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[configurações](/javascript/api/excel/excel.workbookdata#settings)|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|
||[names](/javascript/api/excel/excel.worksheet#names)|Coleção de nomes com escopo para a planilha atual. Somente leitura.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetCount (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Obtém o número de planilhas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Obtém um objeto worksheet usando o Nome ou ID dele. Se a planilha não existir, retornará um objeto null.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[names](/javascript/api/excel/excel.worksheetdata#names)|Coleção de nomes com escopo para a planilha atual. Somente leitura.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
