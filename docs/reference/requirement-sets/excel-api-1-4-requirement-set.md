---
title: Conjunto de requisitos de API JavaScript do Excel 1,4
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,4.
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ed6a9329ae5c0c5839abf5d97a0c67a44783a579
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819760"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

## <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` – Itens com escopo de planilha ou pasta de trabalho.
* `worksheet` -Retorna a planilha na qual o item nomeado tem escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)` – Adiciona um novo nome à coleção de um determinado escopo.
* `addFormulaLocal(name: string, formula: string, comment: string)` – Adiciona um novo nome à coleção do escopo fornecido usando a localidade do usuário para a fórmula.

## <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs incluem `getItem()` para obter a entrada de configuração através da chave e `add()` para adicionar o par de definição de valor-chave especificado: Value à pasta de trabalho.

## <a name="others"></a>Outros

* Definir o nome da coluna da tabela.
* Adicione uma coluna de tabela ao final da tabela.
* Adicionar várias linhas a uma tabela de cada vez.
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Os [ \* métodos e propriedades do OrNullObject](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties): essa funcionalidade permite obter um objeto usando uma chave. Se o objeto não existir, a propriedade do objeto retornado `isNullObject` será true. Isso permite que os desenvolvedores verifiquem se um objeto existe sem precisar tratá-lo por meio da manipulação de exceção. Um `*OrNullObject` método está disponível na maioria dos objetos coleção.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,4. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,4 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,4 ou anterior](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

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
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (Name: String, Reference: \| cadeia de caracteres de intervalo, comentário?: cadeia de caracteres)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Adiciona um novo nome à coleção do escopo fornecido.|
||[addFormulaLocal (Name: String, formula: String, comment?: String)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Obtém o número de itens nomeados na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Obtém um objeto NamedItem usando seu nome. Se o objeto getNamedItem não existir, retornará um objeto null.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Retorna o intervalo usado do objeto range determinado. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Obtém o número de objetos RangeView na coleção.|
|[Configuração](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Exclui a configuração.|
||[key](/javascript/api/excel/excel.setting#key)|Retorna a chave que representa a id da configuração. Somente leitura.|
||[value](/javascript/api/excel/excel.setting#value)|Representa o valor armazenado para esta configuração.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (Key: String, value: String \| número \| Boolean \| Data \| array <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Obtém o número de Configurações na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Obtém uma entrada de configuração por meio da tecla.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Obtém uma entrada de configuração por meio da tecla. Se a Configuração não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Ocorre quando as Configurações no documento são alteradas.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[configurações](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, retornará um objeto null.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Obtém a quantidade de colunas na tabela.|
||[getItemOrNullObject (Key: String de número \| )](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Obtém um objeto column por nome ou ID. Se a coluna não existir, retornará um objeto null.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Obtém a quantidade de linhas na tabela.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[configurações](/javascript/api/excel/excel.workbook#settings)|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|
||[names](/javascript/api/excel/excel.worksheet#names)|Coleção de nomes com escopo para a planilha atual. Somente leitura.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetCount (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Obtém o número de planilhas na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Obtém um objeto worksheet usando o Nome ou ID dele. Se a planilha não existir, retornará um objeto null.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
