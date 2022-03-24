---
title: Conjunto de requisitos da API JavaScript do Word 1.3
description: Detalhes sobre o conjunto de requisitos do WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: d9e0d450b601845d4e11e0fd74652c4e167f802c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746029"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Quais são as novidades na API JavaScript do Word 1.3

O WordApi 1.3 adicionou mais suporte para controles de conteúdo e configurações no nível do documento.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.3. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.3 ou anterior, consulte APIs do Word no conjunto de requisitos [1.3 ou anterior](/javascript/api/word?view=word-js-1.3&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|Cria um novo documento usando um arquivo .docx base64 opcional.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|Obtém o corpo todo, ou então, os pontos inicial ou final do corpo, como um intervalo.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[listas](/javascript/api/word/word.body#word-word-body-lists-member)|Obtém a coleção de listas de objetos no corpo.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|Obtém o corpo pai do corpo.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|Obtém o corpo pai do corpo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que inclui o corpo.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|Obtém a seção pai do corpo.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|Obtém a seção pai do corpo.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Obtém ou define o nome do estilo interno para o corpo.|
||[tables](/javascript/api/word/word.body#word-word-body-tables-member)|Obtém a coleção de tabelas de objetos no corpo.|
||[tipo](/javascript/api/word/word.body#word-word-body-type-member)|Obtém o tipo do corpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|Obtém o controle de todo o conteúdo, ou então, os pontos inicial ou final do controle de conteúdo, como um intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|Obtém os intervalos de texto no controle de conteúdo usando marcas de pontuação e/ou outras marcas finais.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|Insere uma tabela com a quantidade especificada de linhas e colunas dentro ou próxima do controle de conteúdo.|
||[listas](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Obtém a coleção de listas de objetos no controle de conteúdo.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|Obtém o corpo pai do controle de conteúdo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que inclui o controle de conteúdo.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|Obtém a tabela que contém o controle de conteúdo.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|Obtém a célula de tabela que contém o controle de conteúdo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|Obtém a célula de tabela que contém o controle de conteúdo.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|Obtém a tabela que contém o controle de conteúdo.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Divide o controle de conteúdo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Obtém ou define o nome do estilo interno para o controle de conteúdo.|
||[subtipo](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Obtém o subtipo de controle de conteúdo.|
||[tables](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Obtém a coleção de objetos de tabela no controle de conteúdo.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|Obtém um controle de conteúdo pelo respectivo identificador.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|Obtém os controles de conteúdo que têm os tipos especificados e/ou subtipos.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|Obtém o primeiro controle de conteúdo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|Obtém o primeiro controle de conteúdo nesta coleção.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Exclui a propriedade personalizada.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Obtém a chave da propriedade personalizada.|
||[tipo](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Obtém o tipo de valor da propriedade personalizada.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Obtém as propriedades do documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Obtém o objeto body do documento.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Obtém a coleção de objetos de controle de conteúdo no documento.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Abre o documento.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Obtém as propriedades do documento.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Salva o documento.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indica se as alterações do documento foram salvas.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Obtém a coleção de objetos de seção no documento.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|Obtém o nome do aplicativo do documento.|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Obtém ou define o autor do documento.|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Obtém ou define a categoria do documento.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Obtém ou define os comentários do documento.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Obtém ou define a empresa do documento.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|Obtém a data de criação do documento.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|Obtém a coleção de propriedades personalizadas do documento.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Obtém ou define o formato do documento.|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Obtém ou define as palavras-chave do documento.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|Obtém o último autor do documento.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|Obtém a data de impressão do documento.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|Obtém a hora em que o documento foi salvo pela última vez.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Obtém ou define o gerenciador do documento.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|Obtém o número de revisão do documento.|
||[security](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Obtém configurações de segurança do documento.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Obtém ou define o assunto do documento.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Obtém o modelo do documento.|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Obtém ou define o título do documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|Obtém a próxima imagem embutida.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|Obtém a próxima imagem embutida.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|Obtém a imagem, ou então, os pontos inicial ou final da imagem, como um intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que inclui a imagem embutida.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|Obtém a tabela que contém a imagem embutida.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|Obtém a célula de tabela que contém a imagem embutida.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|Obtém a célula de tabela que contém a imagem embutida.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|Obtém a tabela que contém a imagem embutida.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|Obtém a primeira imagem embutida nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|Obtém a primeira imagem embutida nesta coleção.|
|[Lista](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|Obtém os parágrafos que ocorrem no nível especificado na lista.|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|Obtém o marcador, o número ou a imagem no nível especificado como uma cadeia de caracteres.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Obtém a id da lista.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|Verifica se cada um dos 9 níveis existe na lista.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|Obtém todos os tipos de nível 9 na lista.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Obtém parágrafos na lista.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|Define o alinhamento do marcador, número ou imagem no nível especificado na lista.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|Define o formato de marcador no nível especificado na lista.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|Define os dois recuos do nível especificado na lista.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|Define o formato de numeração no nível especificado na lista.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|Define o número inicial no nível especificado na lista.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|Obtém uma lista por seu identificador.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|Obtém uma lista por seu identificador.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|Obtém a primeira lista nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|Obtém a primeira lista nesta coleção.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|Obtém um objeto de lista por seu índice na coleção.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|Obtém todos os itens de lista descendentes do item de lista.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Obtém ou define o nível do item na lista.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|Obtém o marcador de item de lista, número ou imagem como uma cadeia de caracteres.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|Obtém o número da ordem de item de lista em relação a seus irmãos.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|Permite que o parágrafo ingresse em uma lista existente no nível especificado.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|Move este parágrafo para fora de sua lista, caso o parágrafo seja um item da lista.|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|Obtém o próximo parágrafo.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|Obtém o próximo parágrafo.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|Obtém o parágrafo anterior.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|Obtém o parágrafo anterior.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|Obtém o parágrafo inteiro, ou então, os pontos inicial ou final do parágrafo, como um intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|Obtém os intervalos de texto no parágrafo usando marcas de pontuação e/ou outras marcas finais.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|Indica que o parágrafo é o último dentro do corpo do pai.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|Verifica se o parágrafo é um item da lista.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Obtém a lista à qual pertence esse parágrafo.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|Obtém o ListItem para o parágrafo.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|Obtém o ListItem para o parágrafo.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|Obtém a lista à qual pertence esse parágrafo.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|Obtém o corpo pai do parágrafo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que inclui o parágrafo.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|Obtém a tabela que contém o parágrafo.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|Obtém a célula de tabela que contém o parágrafo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|Obtém a célula de tabela que contém o parágrafo.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|Obtém a tabela que contém o parágrafo.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Divide o parágrafo em intervalos filho usando delimitadores.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|Inicia uma nova lista com este parágrafo.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Obtém ou define o nome do estilo interno para o parágrafo.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|Obtém o nível da tabela do parágrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|Obtém o primeiro parágrafo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|Obtém o primeiro parágrafo nesta coleção.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|Obtém o último parágrafo nesta coleção.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|Obtém o último parágrafo nesta coleção.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|Compara o local deste intervalo com a localização de outro intervalo.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|Obtém intervalos filho de hiperlink dentro do intervalo.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|Obtém o próximo intervalo de texto usando marcas de pontuação e/ou outras marcas finais.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|Obtém o próximo intervalo de texto usando marcas de pontuação e/ou outras marcas finais.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|Clona o intervalo, ou então, obtém os pontos inicial ou final do intervalo como um novo intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|Obtém os intervalos filho de texto no intervalo usando marcas de pontuação e/ou outras marcas finais.|
||[hiperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|Verifica se o comprimento do intervalo é zero.|
||[listas](/javascript/api/word/word.range#word-word-range-lists-member)|Obtém a coleção de listas de objetos no intervalo.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|Obtém o corpo pai do intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que inclui o intervalo.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|Obtém a tabela que contém o intervalo.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|Obtém a célula de tabela que contém o intervalo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|Obtém a célula de tabela que contém o intervalo.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|Obtém a tabela que contém o intervalo.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Divide o intervalo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Obtém ou define o nome do estilo interno para o intervalo.|
||[tables](/javascript/api/word/word.range#word-word-range-tables-member)|Obtém a coleção de tabelas de objetos no intervalo.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|Obtém o primeiro intervalo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|Obtém o primeiro intervalo nesta coleção.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Conjunto de api: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|Obtém a próxima seção.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|Obtém a próxima seção.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|Obtém a primeira seção nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|Obtém a primeira seção nesta coleção.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|Adiciona colunas ao início ou no final da tabela, usando a primeira ou última coluna existente como um modelo.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|Adiciona linhas ao início ou no final da tabela, usando a primeira ou última linha existente como um modelo.|
||[alignment](/javascript/api/word/word.table#word-word-table-alignment-member)|Obtém ou define o alinhamento da tabela em relação à coluna de página.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|Autoajusta as colunas da tabela para a largura da janela.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Limpa o conteúdo da tabela.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Exclui toda a tabela.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|Exclui colunas específicas.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|Exclui linha específicas.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|Distribui uniformemente a largura das colunas.|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|Obtém a fonte.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|Obtém o estilo de borda para a borda especificada.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|Obtém a célula da tabela em uma linha e coluna especificada.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|Obtém a célula da tabela em uma linha e coluna especificada.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|Obtém a próxima tabela.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|Obtém a próxima tabela.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|Obtém o parágrafo após a tabela.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|Obtém o parágrafo após a tabela.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|Obtém o parágrafo antes da tabela.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|Obtém o parágrafo antes da tabela.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|Obtém o intervalo que contém esta tabela, ou o intervalo no início ou no final da tabela.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|Obtém e define o número de linhas de cabeçalho.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|Obtém e define o alinhamento horizontal de cada célula na tabela.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|Insere um controle de conteúdo na tabela.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|Indica se todas as linhas de tabela são uniformes.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|Obtém o nível de aninhamento da tabela.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|Obtém o corpo pai da tabela.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|Obtém o controle de conteúdo que contém a tabela.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|Obtém o controle de conteúdo que contém a tabela.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|Obtém a tabela que contém esta tabela.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|Obtém a célula de tabela que contém esta tabela.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|Obtém a célula de tabela que contém esta tabela.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|Obtém a tabela que contém esta tabela.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|Obtém a quantidade de linhas na tabela.|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|Obtém todas as linhas da tabela.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#word-word-table-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto table.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Seleciona a tabela, ou então, a posição no início ou no final da tabela e navega na interface do usuário do Word até ela.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|Obtém e define a cor de sombreamento.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Obtém ou define o nome do estilo usado para a tabela.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|Obtém e define se a tabela tem colunas em tiras.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|Obtém e define se a tabela tem linhas em tiras.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Obtém ou define o nome do estilo interno para a tabela.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|Obtém e define se a tabela tem uma primeira coluna com um estilo especial.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|Obtém e define se a tabela tem uma última coluna com um estilo especial.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|Obtém e define se a tabela tem uma (última) linha total com um estilo especial.|
||[tables](/javascript/api/word/word.table#word-word-table-tables-member)|Obtém as tabelas filho aninhadas em um nível mais profundo.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Obtém e define os valores de texto na tabela, como uma matriz de Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|Obtém e define o alinhamento vertical de cada célula na tabela.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Obtém e define a largura da tabela em pontos.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Obtém ou define a cor da borda da tabela.|
||[tipo](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Obtém ou define o tipo de borda da tabela.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Obtém ou define a largura, em pontos, da borda da tabela.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Obtém o objeto do corpo da célula.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|Obtém o índice da célula em sua linha.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|Obtém e define a largura da coluna da célula em pontos.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|Exclui a coluna que contém essa célula.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|Exclui a linha que contém essa célula.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|Obtém o estilo de borda para a borda especificada.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|Obtém a próxima célula.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|Obtém a próxima célula.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|Obtém e define o alinhamento horizontal da célula.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|Adiciona colunas à esquerda ou à direita da célula, usando a coluna da célula como um modelo.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|Insere linhas acima ou abaixo da célula, usando a linha da célula como um modelo.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|Obtém a linha pai da célula.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|Obtém a tabela pai da célula.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|Obtém o índice da linha da célula na tabela.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|Obtém ou define a cor de sombreamento da célula.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Obtém e define o texto da célula.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|Obtém e define o alinhamento vertical da célula.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Obtém a largura da célula em pontos.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|Obtém a primeira célula da tabela nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|Obtém a primeira célula da tabela nesta coleção.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|Obtém a primeira tabela nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|Obtém a primeira tabela nesta coleção.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|Obtém a quantidade de células na linha.|
||[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Obtém células.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Limpa o conteúdo da linha.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Exclui toda a linha.|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Obtém a fonte.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|Obtém o estilo de borda das células na linha.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|Obtém a próxima linha.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|Obtém a próxima linha.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|Obtém e define o alinhamento horizontal de cada célula na linha.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|Insere linhas usando esta linha como um modelo.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|Verifica se a linha é uma linha de cabeçalho.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|Obtém uma tabela pai.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|Obtém e define a altura da linha preferencial em pontos.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|Obtém o índice da linha em sua tabela pai.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo da linha.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Seleciona a linha e navega na interface do usuário do Word até ele.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|Obtém e define a cor de sombreamento.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Obtém e define os valores de texto na linha, como uma matriz Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|Obtém e define o alinhamento vertical das células na linha.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|Obtém a primeira linha nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|Obtém a primeira linha nesta coleção.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
