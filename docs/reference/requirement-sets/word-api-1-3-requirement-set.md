---
title: Conjunto de requisitos da API JavaScript do Word 1.3
description: Detalhes sobre o conjunto de requisitos do WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 0291a8a96e0ae38bf9d1061a09dac3d3b9cc3ddb
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178101"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Quais são as novidades na API JavaScript do Word 1.3

O WordApi 1.3 adicionou mais suporte para controles de conteúdo e configurações no nível do documento.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.3. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.3 ou anterior, consulte APIs do Word no conjunto de requisitos [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Cria um novo documento usando um arquivo .docx codificado base64 opcional.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Obtém o corpo todo, ou então, os pontos inicial ou final do corpo, como um intervalo.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[lists](/javascript/api/word/word.body#lists)|Obtém a coleção de listas de objetos no corpo.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Obtém o corpo pai do corpo.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Obtém o corpo pai do corpo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o corpo.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Obtém a seção pai do corpo.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Obtém a seção pai do corpo.|
||[tables](/javascript/api/word/word.body#tables)|Obtém a coleção de tabelas de objetos no corpo.|
||[type](/javascript/api/word/word.body#type)|Obtém o tipo do corpo.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Obtém ou define o nome do estilo interno para o corpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Obtém o controle de todo o conteúdo, ou então, os pontos inicial ou final do controle de conteúdo, como um intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos de texto no controle de conteúdo usando marcas de pontuação e/ou outras marcas finais.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas dentro ou próxima do controle de conteúdo.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Obtém a coleção de listas de objetos no controle de conteúdo.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Obtém o corpo pai do controle de conteúdo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o controle de conteúdo.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Obtém a tabela que contém o controle de conteúdo.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Obtém a célula de tabela que contém o controle de conteúdo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Obtém a célula de tabela que contém o controle de conteúdo.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Obtém a tabela que contém o controle de conteúdo.|
||[subtipo](/javascript/api/word/word.contentcontrol#subtype)|Obtém o subtipo de controle de conteúdo.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Obtém a coleção de objetos de tabela no controle de conteúdo.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide o controle de conteúdo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Obtém ou define o nome do estilo interno para o controle de conteúdo.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Obtém um controle de conteúdo pelo respectivo identificador.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Obtém os controles de conteúdo que têm os tipos especificados e/ou subtipos.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Obtém o primeiro controle de conteúdo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Obtém o primeiro controle de conteúdo nesta coleção.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/word/word.customproperty#key)|Obtém a chave da propriedade personalizada.|
||[type](/javascript/api/word/word.customproperty#type)|Obtém o tipo de valor da propriedade personalizada.|
||[value](/javascript/api/word/word.customproperty#value)|Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtém as propriedades do documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Abre o documento.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtém o objeto body do documento.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no documento.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtém as propriedades do documento.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indica se as alterações do documento foram salvas.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtém a coleção de objetos de seção no documento.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Salva o documento.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Obtém ou define o autor do documento.|
||[category](/javascript/api/word/word.documentproperties#category)|Obtém ou define a categoria do documento.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtém ou define os comentários do documento.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtém ou define a empresa do documento.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtém ou define o formato do documento.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Obtém ou define as palavras-chave do documento.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Obtém ou define o gerenciador do documento.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Obtém o nome do aplicativo do documento.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Obtém a data de criação do documento.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Obtém a coleção de propriedades personalizadas do documento.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Obtém o último autor do documento.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Obtém a data de impressão do documento.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Obtém a hora em que o documento foi salvo pela última vez.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Obtém o número de revisão do documento.|
||[security](/javascript/api/word/word.documentproperties#security)|Obtém configurações de segurança do documento.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtém o modelo do documento.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtém ou define o assunto do documento.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtém ou define o título do documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Obtém a próxima imagem embutida.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Obtém a próxima imagem embutida.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Obtém a imagem, ou então, os pontos inicial ou final da imagem, como um intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui a imagem embutida.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Obtém a tabela que contém a imagem embutida.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Obtém a célula de tabela que contém a imagem embutida.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Obtém a célula de tabela que contém a imagem embutida.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Obtém a tabela que contém a imagem embutida.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Obtém a primeira imagem embutida nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Obtém a primeira imagem embutida nesta coleção.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Obtém os parágrafos que ocorrem no nível especificado na lista.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|Obtém o marcador, o número ou a imagem no nível especificado como uma cadeia de caracteres.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado.|
||[id](/javascript/api/word/word.list#id)|Obtém a id da lista.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Verifica se cada um dos 9 níveis existe na lista.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Obtém todos os tipos de nível 9 na lista.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtém parágrafos na lista.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Define o alinhamento do marcador, número ou imagem no nível especificado na lista.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Define o formato de marcador no nível especificado na lista.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Define os dois recuos do nível especificado na lista.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Define o formato de numeração no nível especificado na lista.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Define o número inicial no nível especificado na lista.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Obtém uma lista por seu identificador.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Obtém uma lista por seu identificador.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Obtém a primeira lista nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Obtém a primeira lista nesta coleção.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Obtém um objeto de lista por seu índice na coleção.|
||[items](/javascript/api/word/word.listcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Obtém todos os itens de lista descendentes do item de lista.|
||[level](/javascript/api/word/word.listitem#level)|Obtém ou define o nível do item na lista.|
||[listString](/javascript/api/word/word.listitem#liststring)|Obtém o marcador de item de lista, número ou imagem como uma cadeia de caracteres.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Obtém o número da ordem de item de lista em relação a seus irmãos.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Permite que o parágrafo ingresse em uma lista existente no nível especificado.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Move este parágrafo para fora de sua lista, caso o parágrafo seja um item da lista.|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Obtém o próximo parágrafo.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Obtém o próximo parágrafo.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Obtém o parágrafo anterior.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Obtém o parágrafo anterior.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Obtém o parágrafo inteiro, ou então, os pontos inicial ou final do parágrafo, como um intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos de texto no parágrafo usando marcas de pontuação e/ou outras marcas finais.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indica que o parágrafo é o último dentro do corpo do pai.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Verifica se o parágrafo é um item da lista.|
||[list](/javascript/api/word/word.paragraph#list)|Obtém a lista à qual pertence esse parágrafo.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Obtém o ListItem para o parágrafo.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Obtém o ListItem para o parágrafo.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Obtém a lista à qual pertence esse parágrafo.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Obtém o corpo pai do parágrafo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o parágrafo.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Obtém a tabela que contém o parágrafo.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Obtém a célula de tabela que contém o parágrafo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Obtém a célula de tabela que contém o parágrafo.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Obtém a tabela que contém o parágrafo.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Obtém o nível da tabela do parágrafo.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Divide o parágrafo em intervalos filho usando delimitadores.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Inicia uma nova lista com este parágrafo.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Obtém ou define o nome do estilo interno para o parágrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Obtém o primeiro parágrafo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Obtém o primeiro parágrafo nesta coleção.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Obtém o último parágrafo nesta coleção.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Obtém o último parágrafo nesta coleção.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compara o local deste intervalo com a localização de outro intervalo.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Obtém intervalos filho de hiperlink dentro do intervalo.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Obtém o próximo intervalo de texto usando marcas de pontuação e/ou outras marcas finais.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Obtém o próximo intervalo de texto usando marcas de pontuação e/ou outras marcas finais.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clona o intervalo, ou então, obtém os pontos inicial ou final do intervalo como um novo intervalo.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos filho de texto no intervalo usando marcas de pontuação e/ou outras marcas finais.|
||[hiperlink](/javascript/api/word/word.range#hyperlink)|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Verifica se o comprimento do intervalo é zero.|
||[lists](/javascript/api/word/word.range#lists)|Obtém a coleção de listas de objetos no intervalo.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Obtém o corpo pai do intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o intervalo.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Obtém a tabela que contém o intervalo.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Obtém a célula de tabela que contém o intervalo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Obtém a célula de tabela que contém o intervalo.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Obtém a tabela que contém o intervalo.|
||[tables](/javascript/api/word/word.range#tables)|Obtém a coleção de tabelas de objetos no intervalo.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide o intervalo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Obtém ou define o nome do estilo interno para o intervalo.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Obtém o primeiro intervalo nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Obtém o primeiro intervalo nesta coleção.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[aplicativo](/javascript/api/word/word.requestcontext#application)|[Conjunto de api: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Obtém a próxima seção.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Obtém a próxima seção.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Obtém a primeira seção nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Obtém a primeira seção nesta coleção.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Adiciona colunas ao início ou no final da tabela, usando a primeira ou última coluna existente como um modelo.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Adiciona linhas ao início ou no final da tabela, usando a primeira ou última linha existente como um modelo.|
||[alignment](/javascript/api/word/word.table#alignment)|Obtém ou define o alinhamento da tabela em relação à coluna de página.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Autoajusta as colunas da tabela para a largura da janela.|
||[clear()](/javascript/api/word/word.table#clear--)|Limpa o conteúdo da tabela.|
||[delete()](/javascript/api/word/word.table#delete--)|Exclui toda a tabela.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Exclui colunas específicas.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Exclui linha específicas.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Distribui uniformemente a largura das colunas.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Obtém o estilo de borda para a borda especificada.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Obtém a célula da tabela em uma linha e coluna especificada.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Obtém a célula da tabela em uma linha e coluna especificada.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Obtém a próxima tabela.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Obtém a próxima tabela.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Obtém o parágrafo após a tabela.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Obtém o parágrafo após a tabela.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Obtém o parágrafo antes da tabela.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Obtém o parágrafo antes da tabela.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Obtém o intervalo que contém esta tabela, ou o intervalo no início ou no final da tabela.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Obtém e define o número de linhas de cabeçalho.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Obtém e define o alinhamento horizontal de cada célula na tabela.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Insere um controle de conteúdo na tabela.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Obtém a fonte.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indica se todas as linhas de tabela são uniformes.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Obtém o nível de aninhamento da tabela.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Obtém o corpo pai da tabela.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Obtém o controle de conteúdo que contém a tabela.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que contém a tabela.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Obtém a tabela que contém esta tabela.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Obtém a célula de tabela que contém esta tabela.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Obtém a célula de tabela que contém esta tabela.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Obtém a tabela que contém esta tabela.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Obtém a quantidade de linhas na tabela.|
||[rows](/javascript/api/word/word.table#rows)|Obtém todas as linhas da tabela.|
||[tables](/javascript/api/word/word.table#tables)|Obtém as tabelas filho aninhadas em um nível mais profundo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto table.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Seleciona a tabela, ou então, a posição no início ou no final da tabela e navega na interface do usuário do Word até ela.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Obtém e define a cor de sombreamento.|
||[style](/javascript/api/word/word.table#style)|Obtém ou define o nome do estilo usado para a tabela.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Obtém e define se a tabela tem colunas em tiras.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Obtém e define se a tabela tem linhas em tiras.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Obtém ou define o nome do estilo interno para a tabela.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Obtém e define se a tabela tem uma primeira coluna com um estilo especial.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Obtém e define se a tabela tem uma última coluna com um estilo especial.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Obtém e define se a tabela tem uma (última) linha total com um estilo especial.|
||[values](/javascript/api/word/word.table#values)|Obtém e define os valores de texto na tabela, como uma matriz de Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Obtém e define o alinhamento vertical de cada célula na tabela.|
||[width](/javascript/api/word/word.table#width)|Obtém e define a largura da tabela em pontos.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtém ou define a cor da borda da tabela.|
||[type](/javascript/api/word/word.tableborder#type)|Obtém ou define o tipo de borda da tabela.|
||[width](/javascript/api/word/word.tableborder#width)|Obtém ou define a largura, em pontos, da borda da tabela.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Obtém e define a largura da coluna da célula em pontos.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Exclui a coluna que contém essa célula.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleterow--)|Exclui a linha que contém essa célula.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Obtém o estilo de borda para a borda especificada.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.tablecell#getnext--)|Obtém a próxima célula.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Obtém a próxima célula.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Obtém e define o alinhamento horizontal da célula.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Adiciona colunas à esquerda ou à direita da célula, usando a coluna da célula como um modelo.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Insere linhas acima ou abaixo da célula, usando a linha da célula como um modelo.|
||[body](/javascript/api/word/word.tablecell#body)|Obtém o objeto do corpo da célula.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Obtém o índice da célula em sua linha.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Obtém a linha pai da célula.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Obtém a tabela pai da célula.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Obtém o índice da linha da célula na tabela.|
||[width](/javascript/api/word/word.tablecell#width)|Obtém a largura da célula em pontos.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Obtém ou define a cor de sombreamento da célula.|
||[value](/javascript/api/word/word.tablecell#value)|Obtém e define o texto da célula.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Obtém e define o alinhamento vertical da célula.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Obtém a primeira célula da tabela nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Obtém a primeira célula da tabela nesta coleção.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Obtém a primeira tabela nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Obtém a primeira tabela nesta coleção.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Limpa o conteúdo da linha.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Exclui toda a linha.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Obtém o estilo de borda das células na linha.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext()](/javascript/api/word/word.tablerow#getnext--)|Obtém a próxima linha.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Obtém a próxima linha.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Obtém e define o alinhamento horizontal de cada célula na linha.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Insere linhas usando esta linha como um modelo.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Obtém e define a altura da linha preferencial em pontos.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Obtém a quantidade de células na linha.|
||[cells](/javascript/api/word/word.tablerow#cells)|Obtém células.|
||[font](/javascript/api/word/word.tablerow#font)|Obtém a fonte.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Verifica se a linha é uma linha de cabeçalho.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Obtém uma tabela pai.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Obtém o índice da linha em sua tabela pai.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Executa uma pesquisa com as SearchOptions especificadas no escopo da linha.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Seleciona a linha e navega na interface do usuário do Word até ele.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Obtém e define a cor de sombreamento.|
||[values](/javascript/api/word/word.tablerow#values)|Obtém e define os valores de texto na linha, como uma matriz Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Obtém e define o alinhamento vertical das células na linha.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Obtém a primeira linha nesta coleção.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Obtém a primeira linha nesta coleção.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
