---
title: APIs de visualização javascript do Word
description: Detalhes sobre as FUTURAS APIs JavaScript do Word.
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
---

# <a name="word-javascript-preview-apis"></a>APIs de visualização javascript do Word

As novas APIs JavaScript do Word são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são adquiridos.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs JavaScript do Word atualmente em visualização, exceto as que estão disponíveis apenas [em](#web-only-api-list) Word na Web. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs lançadas [anteriormente), consulte todas as APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Ocorre quando os dados dentro do controle de conteúdo são alterados.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Ocorre quando o controle de conteúdo é excluído.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Ocorre quando a seleção dentro do controle de conteúdo é alterada.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|O objeto que gerou o evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|O tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Exclui a parte XML personalizada.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Exclui um atributo com o nome dado do elemento identificado pelo xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Exclui o elemento identificado pelo xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|Obtém o conteúdo XML completo da parte XML personalizada.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Obtém a ID da parte XML personalizada.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Insere um atributo com o nome e o valor determinados ao elemento identificado pelo xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Insere o XML determinado no elemento pai identificado pelo xpath no índice de posição filho.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|Obtém o URI do namespace da parte XML personalizada.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Consulta o conteúdo XML da parte XML personalizada.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|Define o conteúdo XML completo da parte XML personalizada.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Atualiza o valor de um atributo com o nome dado do elemento identificado pelo xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Atualiza o XML do elemento identificado pelo xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Adiciona uma nova parte XML personalizada ao documento.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|Obtém o número de itens na coleção.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|Obtém uma parte XML personalizada com base em sua ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Se o conjunto contiver exatamente um item, esse método o retornará.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|Obtém as partes XML personalizadas no documento.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|Obtém o intervalo de um indicador.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Ocorre quando um controle de conteúdo é adicionado.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|Executa uma pesquisa com as opções de pesquisa especificadas no escopo de todo o documento.|
||[configurações](/javascript/api/word/word.document#word-word-document-settings-member)|Obtém as configurações do complemento no documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|Obtém as partes XML personalizadas no documento.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|Exclui um indicador, se existir, do documento.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|Obtém o intervalo de um indicador.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|Obtém o intervalo de um indicador.|
||[configurações](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|Obtém as configurações do complemento no documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Obtém o formato da imagem em linha.|
|[Lista](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Obtém a fonte do marcador, número ou imagem no nível especificado na lista.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Obtém a representação de cadeia de caracteres codificada base64 da imagem no nível especificado na lista.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Redefine a fonte do marcador, número ou imagem no nível especificado na lista.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Define a imagem no nível especificado na lista.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|Obtém os nomes de todos os indicadores ou sobrepostos ao intervalo.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|Insere um indicador no intervalo.|
|[Configuração](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Exclui a configuração.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Obtém a chave da configuração.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Obtém ou define o valor da configuração.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Cria uma nova configuração ou define uma configuração existente.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|Exclui todas as configurações deste add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|Obtém a contagem de configurações.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|Obtém um objeto de configuração por sua chave, que é sensível a minúsculas.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|Mescla as células delimitadas inclusive por uma primeira e última célula.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Divide a célula no número especificado de linhas e colunas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Insere um controle de conteúdo na linha.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Mescla a linha em uma célula.|

## <a name="web-only-api-list"></a>Lista de API somente na Web

A tabela a seguir lista as APIs JavaScript do Word atualmente em visualização apenas Word na Web. Para ver uma lista completa de todas as APIs JavaScript do Word (incluindo APIs de visualização e APIs lançadas [anteriormente), consulte todas as APIs JavaScript do Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Obtém a coleção de notas de fim no corpo.|
||[notas de rodapé](/javascript/api/word/word.body#word-word-body-footnotes-member)|Obtém a coleção de notas de rodapé no corpo.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Obtém comentários associados ao corpo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
||[tipo](/javascript/api/word/word.body#word-word-body-type-member)|Obtém o tipo do corpo.|
|[Comentário](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Obtém o nome do autor do comentário.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Obtém ou define o conteúdo do comentário como texto sem texto.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Obtém ou define o status do thread de comentário.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Obtém a data de criação do comentário.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Exclui o comentário e suas respostas.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Obtém o intervalo no documento principal em que o comentário está.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Obtém a coleção de objetos de resposta associados ao comentário.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Adiciona uma nova resposta ao final do thread de comentário.|
||[resolvido](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Obtém ou define o status do thread de comentário.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Obtém o primeiro comentário na coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Obtém o primeiro comentário na coleção.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|Obtém um objeto comment por seu índice na coleção.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Obtém ou define um valor que indica se o texto do comentário está em negrito.|
||[hiperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Insere texto no local especificado.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Verifica se o comprimento do intervalo é zero.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Obtém ou define um valor que indica se o texto do comentário é itálico.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Obtém ou define um valor que indica se o texto do comentário tem um tachado.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Obtém o texto do intervalo de comentários.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Obtém ou define um valor que indica o tipo de sublinhado do texto de comentário.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Obtém o nome do autor da resposta do comentário.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Obtém ou define o conteúdo da resposta do comentário.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Obtém ou define o intervalo de conteúdo do commentReply.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Obtém a data de criação da resposta de comentário.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Exclui a resposta do comentário. |
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Obtém o comentário pai desta resposta.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Obtém a primeira resposta de comentário na coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Obtém a primeira resposta de comentário na coleção.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|Obtém um objeto de resposta de comentário pelo índice na coleção.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Obtém a coleção de notas de fim no controle de conteúdo.|
||[notas de rodapé](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Obtém a coleção de notas de rodapé no controle de conteúdo.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Obtém comentários associados ao corpo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Obtém ou define o modo ChangeTracking.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Obtém as notas de fim do documento em um único corpo.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Obtém as notas de rodapé do documento em um único corpo.|
|[Item de nota](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Representa o objeto body do item de anotação.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Exclui o item de anotação.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Obtém o próximo item de anotação do mesmo tipo.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Obtém o próximo item de anotação do mesmo tipo.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Representa uma referência de nota de rodapé ou nota de fim no documento principal.|
||[tipo](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Representa o tipo de item de nota: nota de rodapé ou nota de fim.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Obtém o primeiro item de anotação nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Obtém o primeiro item de anotação nesta coleção.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Obtém a coleção de notas de fim no parágrafo.|
||[notas de rodapé](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Obtém a coleção de notas de rodapé no parágrafo.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Obtém comentários associados ao parágrafo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Obtém a coleção de notas de fim no intervalo.|
||[notas de rodapé](/javascript/api/word/word.range#word-word-range-footnotes-member)|Obtém a coleção de notas de rodapé no intervalo.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Obtém comentários associados ao intervalo.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Obtém texto revisado com base na seleção ChangeTrackingVersion.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insira um comentário no intervalo.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Insere uma nota de fim.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Insere uma nota de rodapé.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Obtém a coleção de notas de fim na tabela.|
||[notas de rodapé](/javascript/api/word/word.table#word-word-table-footnotes-member)|Obtém a coleção de notas de rodapé na tabela.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Obtém a coleção de notas de fim na linha de tabela.|
||[notas de rodapé](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Obtém a coleção de notas de rodapé na linha da tabela.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
