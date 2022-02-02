---
title: APIs de visualização javascript do Word
description: Detalhes sobre as FUTURAS APIs JavaScript do Word.
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 4ef8bd9897689b354fa7c19ba0d7be7f8fb92be9
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320155"
---
# <a name="word-javascript-preview-apis"></a>APIs de visualização javascript do Word

As novas APIs JavaScript do Word são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são adquiridos.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs JavaScript do Word atualmente em visualização, exceto as que estão disponíveis apenas [em](#web-only-api-list) Word na Web. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs lançadas [anteriormente), consulte todas as APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Ocorre quando os dados dentro do controle de conteúdo são alterados.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Ocorre quando o controle de conteúdo é excluído.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Ocorre quando a seleção dentro do controle de conteúdo é alterada.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|O objeto que gerou o evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|O tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Exclui a parte XML personalizada.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Exclui um atributo com o nome dado do elemento identificado pelo xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Exclui o elemento identificado pelo xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtém a ID da parte XML personalizada.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Insere um atributo com o nome e o valor determinados ao elemento identificado pelo xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Insere o XML determinado no elemento pai identificado pelo xpath no índice de posição filho.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Obtém o URI do namespace da parte XML personalizada.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Consulta o conteúdo XML da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Define o conteúdo XML completo da parte XML personalizada.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Atualiza o valor de um atributo com o nome dado do elemento identificado pelo xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Atualiza o XML do elemento identificado pelo xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Adiciona uma nova parte XML personalizada ao documento.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Obtém as partes XML personalizadas no documento.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Obtém o intervalo de um indicador.|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Ocorre quando um controle de conteúdo é adicionado.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#search_searchText__searchOptions_)|Executa uma pesquisa com as opções de pesquisa especificadas no escopo de todo o documento.|
||[configurações](/javascript/api/word/word.document#settings)|Obtém as configurações do complemento no documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Obtém as partes XML personalizadas no documento.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Obtém o intervalo de um indicador.|
||[configurações](/javascript/api/word/word.documentcreated#settings)|Obtém as configurações do complemento no documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Obtém o formato da imagem em linha.|
|[Lista](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|Obtém a fonte do marcador, número ou imagem no nível especificado na lista.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Obtém a representação de cadeia de caracteres codificada base64 da imagem no nível especificado na lista.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Redefine a fonte do marcador, número ou imagem no nível especificado na lista.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Define a imagem no nível especificado na lista.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Obtém os nomes de todos os indicadores ou sobrepostos ao intervalo.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|Insere um indicador no intervalo.|
|[Configuração](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Exclui a configuração.|
||[key](/javascript/api/word/word.setting#key)|Obtém a chave da configuração.|
||[value](/javascript/api/word/word.setting#value)|Obtém ou define o valor da configuração.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Cria uma nova configuração ou define uma configuração existente.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Exclui todas as configurações deste add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Obtém a contagem de configurações.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Mescla as células delimitadas inclusive por uma primeira e última célula.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Divide a célula no número especificado de linhas e colunas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Insere um controle de conteúdo na linha.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Mescla a linha em uma célula.|

## <a name="web-only-api-list"></a>Lista de API somente na Web

A tabela a seguir lista as APIs JavaScript do Word atualmente em visualização apenas Word na Web. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs lançadas [anteriormente), consulte todas as APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#endnotes)|Obtém a coleção de notas de fim no corpo.|
||[notas de rodapé](/javascript/api/word/word.body#footnotes)|Obtém a coleção de notas de rodapé no corpo.|
||[getComments()](/javascript/api/word/word.body#getComments__)|Obtém comentários associados ao corpo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#getReviewedText_changeTrackingVersion_)|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
||[tipo](/javascript/api/word/word.body#type)|Obtém o tipo do corpo.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#authorEmail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/word/word.comment#authorName)|Obtém o nome do autor do comentário.|
||[content](/javascript/api/word/word.comment#content)|Obtém ou define o conteúdo do comentário como texto sem texto.|
||[contentRange](/javascript/api/word/word.comment#contentRange)|Obtém ou define o status do thread de comentário.|
||[creationDate](/javascript/api/word/word.comment#creationDate)|Obtém a data de criação do comentário.|
||[delete()](/javascript/api/word/word.comment#delete__)|Exclui o comentário e suas respostas.|
||[getRange()](/javascript/api/word/word.comment#getRange__)|Obtém o intervalo no documento principal em que o comentário está.|
||[id](/javascript/api/word/word.comment#id)|ID|
||[replies](/javascript/api/word/word.comment#replies)|Obtém a coleção de objetos de resposta associados ao comentário.|
||[reply(replyText: string)](/javascript/api/word/word.comment#reply_replyText_)|Adiciona uma nova resposta ao final do thread de comentário.|
||[resolvido](/javascript/api/word/word.comment#resolved)|Obtém ou define o status do thread de comentário.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|Obtém o primeiro comentário na coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|Obtém o primeiro comentário na coleção.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#getItem_index_)|Obtém um objeto comment por seu índice na coleção.|
||[items](/javascript/api/word/word.commentcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#bold)|Obtém ou define um valor que indica se o texto do comentário está em negrito.|
||[hiperlink](/javascript/api/word/word.commentcontentrange#hyperlink)|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#insertText_text__insertLocation_)|Insere texto no local especificado.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#isEmpty)|Verifica se o comprimento do intervalo é zero.|
||[italic](/javascript/api/word/word.commentcontentrange#italic)|Obtém ou define um valor que indica se o texto do comentário é itálico.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#strikeThrough)|Obtém ou define um valor que indica se o texto do comentário tem um tachado.|
||[texto](/javascript/api/word/word.commentcontentrange#text)|Obtém o texto do intervalo de comentários.|
||[underline](/javascript/api/word/word.commentcontentrange#underline)|Obtém ou define um valor que indica o tipo de sublinhado do texto de comentário.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/word/word.commentreply#authorName)|Obtém o nome do autor da resposta do comentário.|
||[content](/javascript/api/word/word.commentreply#content)|Obtém ou define o conteúdo da resposta do comentário.|
||[contentRange](/javascript/api/word/word.commentreply#contentRange)|Obtém ou define o intervalo de conteúdo do commentReply.|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|Obtém a data de criação da resposta de comentário.|
||[delete()](/javascript/api/word/word.commentreply#delete__)|Exclui a resposta do comentário. |
||[id](/javascript/api/word/word.commentreply#id)|ID|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|Obtém o comentário pai desta resposta.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|Obtém a primeira resposta de comentário na coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|Obtém a primeira resposta de comentário na coleção.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#getItem_index_)|Obtém um objeto de resposta de comentário pelo índice na coleção.|
||[items](/javascript/api/word/word.commentreplycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#endnotes)|Obtém a coleção de notas de fim no controle de conteúdo.|
||[notas de rodapé](/javascript/api/word/word.contentcontrol#footnotes)|Obtém a coleção de notas de rodapé no controle de conteúdo.|
||[getComments()](/javascript/api/word/word.contentcontrol#getComments__)|Obtém comentários associados ao corpo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#getReviewedText_changeTrackingVersion_)|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#changeTrackingMode)|Obtém ou define o modo ChangeTracking.|
||[getEndnoteBody()](/javascript/api/word/word.document#getEndnoteBody__)|Obtém as notas de fim do documento em um único corpo.|
||[getFootnoteBody()](/javascript/api/word/word.document#getFootnoteBody__)|Obtém as notas de rodapé do documento em um único corpo.|
|[Item de nota](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#body)|Representa o objeto body do item de anotação.|
||[delete()](/javascript/api/word/word.noteitem#delete__)|Exclui o item de anotação.|
||[getNext()](/javascript/api/word/word.noteitem#getNext__)|Obtém o próximo item de anotação do mesmo tipo.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#getNextOrNullObject__)|Obtém o próximo item de anotação do mesmo tipo.|
||[reference](/javascript/api/word/word.noteitem#reference)|Representa uma referência de nota de rodapé ou nota de fim no documento principal.|
||[tipo](/javascript/api/word/word.noteitem#type)|Representa o tipo de item de nota: nota de rodapé ou nota de fim.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|Obtém o primeiro item de anotação nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|Obtém o primeiro item de anotação nesta coleção.|
||[items](/javascript/api/word/word.noteitemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#endnotes)|Obtém a coleção de notas de fim no parágrafo.|
||[notas de rodapé](/javascript/api/word/word.paragraph#footnotes)|Obtém a coleção de notas de rodapé no parágrafo.|
||[getComments()](/javascript/api/word/word.paragraph#getComments__)|Obtém comentários associados ao parágrafo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#getReviewedText_changeTrackingVersion_)|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#endnotes)|Obtém a coleção de notas de fim no intervalo.|
||[notas de rodapé](/javascript/api/word/word.range#footnotes)|Obtém a coleção de notas de rodapé no intervalo.|
||[getComments()](/javascript/api/word/word.range#getComments__)|Obtém comentários associados ao intervalo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#getReviewedText_changeTrackingVersion_)|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#insertComment_commentText_)|Insira um comentário no intervalo.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#insertEndnote_insertText_)|Insere uma nota de fim.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#insertFootnote_insertText_)|Insere uma nota de rodapé.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#endnotes)|Obtém a coleção de notas de fim na tabela.|
||[notas de rodapé](/javascript/api/word/word.table#footnotes)|Obtém a coleção de notas de rodapé na tabela.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#endnotes)|Obtém a coleção de notas de fim na linha de tabela.|
||[notas de rodapé](/javascript/api/word/word.tablerow#footnotes)|Obtém a coleção de notas de rodapé na linha da tabela.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
