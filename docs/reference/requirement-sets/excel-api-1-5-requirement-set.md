---
title: Conjunto de requisitos de API JavaScript do Excel 1,5
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,5
ms.date: 07/15/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b8f767a83b7e373b422b6fc0d9ac65de90c04f5
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771964"
---
#  <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

ExcelApi 1,5 adiciona partes XML personalizadas. Eles podem ser acessados por meio da [coleção de partes XML personalizadas](/javascript/api/excel/excel.workbook#customxmlparts) no objeto Workbook.

## <a name="custom-xml-part"></a>Parte XML personalizada

* Obtenha partes XML personalizadas usando sua ID.
* Obtenção de um novo conjunto com escopo de partes XML personalizadas cujos namespaces correspondam ao namespace especificado.
* Obtenha uma cadeia de caracteres XML associada a uma parte.
* Forneça a ID e o namespace de uma parte.
* Adicione uma nova parte XML personalizada à pasta de trabalho.
* Defina uma parte de XML inteira.
* Exclua uma parte XML personalizada.
* Exclua um atributo com o nome especificado do elemento identificado por xpath.
* Consulte o conteúdo XML por xpath.
* Inserir, atualizar e excluir atributos.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Exclui a parte XML personalizada.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|A ID da parte XML personalizada. Somente leitura.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|O URI do namespace da parte XML personalizada. Somente leitura.|
||[setXml (XML: String)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Define o conteúdo XML completo da parte XML personalizada.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML: String)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|
||[getByNamespace (namespaceUri: cadeia de caracteres)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtém o número de partes CustomXml na coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#id)|Para cada ITEM na coleção: a ID da parte XML personalizada. Somente leitura.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#namespaceuri)|Para cada ITEM na coleção: o URI do namespace da parte XML personalizada. Somente leitura.|
|[CustomXmlPartData](/javascript/api/excel/excel.customxmlpartdata)|[id](/javascript/api/excel/excel.customxmlpartdata#id)|A ID da parte XML personalizada. Somente leitura.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartdata#namespaceuri)|O URI do namespace da parte XML personalizada. Somente leitura.|
|[CustomXmlPartLoadOptions](/javascript/api/excel/excel.customxmlpartloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartloadoptions#id)|A ID da parte XML personalizada. Somente leitura.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartloadoptions#namespaceuri)|O URI do namespace da parte XML personalizada. Somente leitura.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtém o número de partes CustomXML nesta coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#id)|Para cada ITEM na coleção: a ID da parte XML personalizada. Somente leitura.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#namespaceuri)|Para cada ITEM na coleção: o URI do namespace da parte XML personalizada. Somente leitura.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Id da Tabela Dinâmica. Somente leitura.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[id](/javascript/api/excel/excel.pivottablecollectionloadoptions#id)|Para cada ITEM na coleção: ID da tabela dinâmica. Somente leitura.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[id](/javascript/api/excel/excel.pivottabledata#id)|Id da Tabela Dinâmica. Somente leitura.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[id](/javascript/api/excel/excel.pivottableloadoptions#id)|Id da Tabela Dinâmica. Somente leitura.|
|[Tempo](/javascript/api/excel/excel.runtime)|[Set (Propriedades: Excel. Runtime)](/javascript/api/excel/excel.runtime#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. RuntimeUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.runtime#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[$all](/javascript/api/excel/excel.runtimeloadoptions#$all)||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Representa a coleção de partes XML personalizadas contidas por esta pasta de trabalho. Somente leitura.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[customXmlParts](/javascript/api/excel/excel.workbookdata#customxmlparts)|Representa a coleção de partes XML personalizadas contidas por esta pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtém a planilha que segue esta. Se não houver planilhas após esta, este método gerará um erro.|
||[getNextOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtém a planilha que segue esta. Se não houver planilhas após esta, este método retornará um objeto NULL.|
||[getprevious (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtém a planilha que precede esta. Se não houver planilhas anteriores, este método gerará um erro.|
||[getPreviousOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtém a planilha que precede esta. Se não houver planilhas anteriores, este método retornará um objeto NULL.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetFirst (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtém a primeira planilha na coleção.|
||[GetLast (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtém a última planilha na coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
