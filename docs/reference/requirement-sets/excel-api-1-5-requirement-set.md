---
title: Conjunto de requisitos de API JavaScript do Excel 1,5
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,5.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 901ea29253bdfee3aeeefd1595eda32d70e4f1e1
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996253"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

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

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,5. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,5 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,5 ou anterior](/javascript/api/excel?view=excel-js-1.5&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Exclui a parte XML personalizada.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|A ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|O URI do namespace da parte XML personalizada.|
||[setXml (XML: String)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Define o conteúdo XML completo da parte XML personalizada.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML: String)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|
||[getByNamespace (namespaceUri: cadeia de caracteres)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtém o número de partes CustomXml na coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtém o número de partes CustomXML nesta coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Id da Tabela Dinâmica.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[API Set: ExcelApi 1,5]|
|[Runtime](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Representa a coleção de partes XML personalizadas contidas por esta pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtém a planilha que segue esta.|
||[getNextOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtém a planilha que segue esta.|
||[getprevious (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtém a planilha que precede esta.|
||[getPreviousOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtém a planilha que precede esta.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetFirst (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtém a primeira planilha na coleção.|
||[GetLast (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtém a última planilha na coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
