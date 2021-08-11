---
title: APIs de visualização javascript do Word
description: Detalhes sobre as APIs JavaScript do Word futuras
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 1a3871ca4445e595620112bb5176fe2b7ab39015228a1602c119c06730cc90ab
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097842"
---
# <a name="word-javascript-preview-apis"></a>APIs de visualização javascript do Word

As novas APIs JavaScript do Word são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são adquiridos.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs JavaScript do Word atualmente em visualização. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Ocorre quando os dados dentro do controle de conteúdo são alterados.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Ocorre quando o controle de conteúdo é excluído.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Ocorre quando a seleção dentro do controle de conteúdo é alterada.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|O objeto que gerou o evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|O tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Exclui a parte XML personalizada.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Exclui um atributo com o nome dado do elemento identificado pelo xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Exclui o elemento identificado pelo xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Insere um atributo com o nome e o valor determinados ao elemento identificado pelo xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Insere o XML determinado no elemento pai identificado pelo xpath no índice de posição filho.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Consulta o conteúdo XML da parte XML personalizada.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtém a ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Obtém o URI do namespace da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Define o conteúdo XML completo da parte XML personalizada.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Atualiza o valor de um atributo com o nome dado do elemento identificado pelo xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Atualiza o XML do elemento identificado pelo xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Adiciona uma nova parte XML personalizada ao documento.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deletebookmark-name-)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Obtém as partes XML personalizadas no documento.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Ocorre quando um controle de conteúdo é adicionado.|
||[configurações](/javascript/api/word/word.document#settings)|Obtém as configurações do complemento no documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Obtém as partes XML personalizadas no documento.|
||[configurações](/javascript/api/word/word.documentcreated#settings)|Obtém as configurações do complemento no documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Obtém o formato da imagem em linha.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getlevelfont-level-)|Obtém a fonte do marcador, número ou imagem no nível especificado na lista.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|Obtém a representação de cadeia de caracteres codificada base64 da imagem no nível especificado na lista.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Redefine a fonte do marcador, número ou imagem no nível especificado na lista.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Define a imagem no nível especificado na lista.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Obtém os nomes de todos os indicadores ou sobrepostos ao intervalo.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertbookmark-name-)|Insere um indicador no intervalo.|
|[Configuração](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Exclui a configuração.|
||[key](/javascript/api/word/word.setting#key)|Obtém a chave da configuração.|
||[value](/javascript/api/word/word.setting#value)|Obtém ou define o valor da configuração.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Cria uma nova configuração ou define uma configuração existente.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteall--)|Exclui todas as configurações deste add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Obtém a contagem de configurações.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Mescla as células delimitadas inclusive por uma primeira e última célula.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Divide a célula no número especificado de linhas e colunas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Insere um controle de conteúdo na linha.|
||[merge()](/javascript/api/word/word.tablerow#merge--)|Mescla a linha em uma célula.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
