---
title: Excel Conjunto de requisitos da API JavaScript 1.5
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 01a13a0f531eae9eea2c213ba0da764fbe51ee15
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671804"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

O ExcelApi 1.5 adiciona partes XML personalizadas. Eles são acessíveis por meio da [coleção de](/javascript/api/excel/excel.workbook#customxmlparts) partes XML personalizadas no objeto da workbook.

## <a name="custom-xml-part"></a>Parte XML personalizada

* Obter partes XML personalizadas usando sua ID.
* Obtenção de um novo conjunto com escopo de partes XML personalizadas cujos namespaces correspondam ao namespace especificado.
* Obter uma cadeia de caracteres XML associada a uma parte.
* Forneça a ID e o namespace de uma parte.
* Adicione uma nova parte XML personalizada à workbook.
* Definir uma parte XML inteira.
* Exclua uma parte XML personalizada.
* Exclua um atributo com o nome especificado do elemento identificado por xpath.
* Consulte o conteúdo XML por xpath.
* Inserir, atualizar e excluir atributos.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no Excel de requisitos da API JavaScript 1.5. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.5 ou anterior, consulte Excel APIs no conjunto de requisitos [1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete__)|Exclui a parte XML personalizada.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getXml__)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|A ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceUri)|URI do namespace da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#setXml_xml_)|Define o conteúdo XML completo da parte XML personalizada.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add_xml_)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getByNamespace_namespaceUri_)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getCount__)|Obtém o número de partes XML personalizadas na coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getCount__)|Obtém o número de partes CustomXML nesta coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItem__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID da tabela dinâmica.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)||
|[Runtime](/javascript/api/excel/excel.runtime)|||
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customXmlParts)|Representa a coleção de partes XML personalizadas contidas nesta workbook.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNext_visibleOnly_)|Obtém a planilha que segue esta.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNextOrNullObject_visibleOnly_)|Obtém a planilha que segue esta.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPrevious_visibleOnly_)|Obtém a planilha que precede essa.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPreviousOrNullObject_visibleOnly_)|Obtém a planilha que precede essa.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getFirst_visibleOnly_)|Obtém a primeira planilha na coleção.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getLast_visibleOnly_)|Obtém a última planilha na coleção.|

## <a name="see-also"></a>Confira também

* [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
