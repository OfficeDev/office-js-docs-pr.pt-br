---
title: APIs de visualização JavaScript do Word
description: Detalhes sobre as APIs JavaScript do Word futuro
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 6a3b67e65c4ced3f1b89d98afe45d5d6c33f63b6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996400"
---
# <a name="word-javascript-preview-apis"></a>APIs de visualização JavaScript do Word

Novas APIs JavaScript do Word são primeiro introduzidas em "Preview" e mais tarde se tornam parte de um conjunto de requisitos específico e numerado após o teste suficiente e o feedback do usuário é adquirido.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Word atualmente em versão prévia. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs previamente lançadas), confira [todas as APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Ocorre quando os dados no controle de conteúdo são alterados.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Ocorre quando o controle de conteúdo é excluído.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Ocorre quando a seleção no controle de conteúdo é alterada.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|O objeto que disparou o evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|O tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Exclui a parte XML personalizada.|
||[DeleteAttribute (XPath: cadeia de caracteres, namespaceMappings: any, Name: String)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Exclui um atributo com o nome fornecido do elemento identificado por XPath.|
||[deleteelement (XPath: cadeia de caracteres, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Exclui o elemento identificado por XPath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[InsertAttribute (XPath: String, namespaceMappings: any, Name: String, value: String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Insere um atributo com o nome e o valor fornecidos para o elemento identificado por XPath.|
||[insertelement (XPath: String, XML: String, namespaceMappings: any, index?: Number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Insere o XML especificado no elemento pai identificado pelo XPath no índice de posição de filho.|
||[consulta (XPath: cadeia de caracteres, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Consulta o conteúdo XML da parte XML personalizada.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtém a ID da parte XML personalizada.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Obtém o URI do namespace da parte XML personalizada.|
||[setXml (XML: String)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Define o conteúdo XML completo da parte XML personalizada.|
||[UpdateAttribute (XPath: String, namespaceMappings: any, Name: String, value: String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Atualiza o valor de um atributo com o nome fornecido do elemento identificado por XPath.|
||[updateElement (XPath: String, XML: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Atualiza o XML do elemento identificado pelo XPath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML: String)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Adiciona uma nova parte XML personalizada ao documento.|
||[getByNamespace (namespaceUri: cadeia de caracteres)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
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
|[Document](/javascript/api/word/word.document)|[deleteBookmark (Name: String)](/javascript/api/word/word.document#deletebookmark-name-)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.document#getbookmarkrange-name-)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Obtém as partes XML personalizadas no documento.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Ocorre quando um controle de conteúdo é adicionado.|
||[configurações](/javascript/api/word/word.document#settings)|Obtém as configurações do suplemento no documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (Name: String)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Obtém o intervalo de um indicador.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Obtém as partes XML personalizadas no documento.|
||[configurações](/javascript/api/word/word.documentcreated#settings)|Obtém as configurações do suplemento no documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Obtém o formato da imagem embutida.|
|[List](/javascript/api/word/word.list)|[getLevelFont (Level: Number)](/javascript/api/word/word.list#getlevelfont-level-)|Obtém a fonte do marcador, o número ou a imagem no nível especificado na lista.|
||[getLevelPicture (Level: Number)](/javascript/api/word/word.list#getlevelpicture-level-)|Obtém a representação de cadeia de caracteres codificada em base64 da imagem no nível especificado na lista.|
||[resetLevelFont (Level: Number, resetFontName?: Boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Redefine a fonte do marcador, o número ou a imagem no nível especificado na lista.|
||[setLevelPicture (Level: Number, base64EncodedImage?: String)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Define a imagem no nível especificado na lista.|
|[Range](/javascript/api/word/word.range)|[getbookmarks (includeHidden?: Boolean, includeAdjacent?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Obtém os nomes de todos os indicadores ou sobrepondo o intervalo.|
||[insertBookmark (Name: String)](/javascript/api/word/word.range#insertbookmark-name-)|Insere um indicador no intervalo.|
|[Configuração](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Exclui a configuração.|
||[key](/javascript/api/word/word.setting#key)|Obtém a chave da configuração.|
||[value](/javascript/api/word/word.setting#value)|Obtém ou define o valor da configuração.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key: String, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Cria uma nova configuração ou define uma configuração existente.|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Exclui todas as configurações deste suplemento.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Obtém a contagem de configurações.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Obtém um objeto Setting por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Obtém um objeto Setting por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow: Number, firstCell: Number, bottomRow: Number, lastCell: Number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Mescla as células delimitadas por inclusivo pela primeira e última célula.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (rowCount: Number, columnCount: Number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Divide a célula no número especificado de linhas e colunas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Insere um controle de conteúdo na linha.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Mescla a linha em uma célula.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
