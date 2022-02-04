---
title: Excel conjunto de requisitos da API JavaScript 1.5
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

O ExcelApi 1.5 adiciona partes XML personalizadas. Eles são acessíveis por meio da [coleção de partes XML](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) personalizadas no objeto da workbook.

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

A tabela a seguir lista as APIs no Excel de requisitos da API JavaScript 1.5. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.5 ou anterior, consulte Excel APIs no conjunto de requisitos [1.5 ou anterior](/javascript/api/excel?view=excel-js-1.5&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-delete-member(1))|Exclui a parte XML personalizada.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-getxml-member(1))|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-id-member)|A ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-namespaceuri-member)|URI do namespace da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-setxml-member(1))|Define o conteúdo XML completo da parte XML personalizada.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-add-member(1))|Adiciona uma nova parte XML personalizada à pasta de trabalho.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getbynamespace-member(1))|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getcount-member(1))|Obtém o número de partes XML personalizadas na coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitem-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitemornullobject-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getcount-member(1))|Obtém o número de partes CustomXML nesta coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitem-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitemornullobject-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitem-member(1))|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-id-member)|ID da tabela dinâmica.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#excel-excel-requestcontext-runtime-member)||
|[Tempo de execução](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member)|Representa a coleção de partes XML personalizadas contidas nesta workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnext-member(1))|Obtém a planilha que segue esta.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnextornullobject-member(1))|Obtém a planilha que segue esta.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getprevious-member(1))|Obtém a planilha que precede essa.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getpreviousornullobject-member(1))|Obtém a planilha que precede essa.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getfirst-member(1))|Obtém a primeira planilha na coleção.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getlast-member(1))|Obtém a última planilha na coleção.|

## <a name="see-also"></a>Confira também

* [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
