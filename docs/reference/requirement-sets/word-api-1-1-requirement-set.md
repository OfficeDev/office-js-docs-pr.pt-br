---
title: Conjunto de requisitos da API JavaScript do Word 1.1
description: Detalhes sobre o conjunto de requisitos do WordApi 1.1.
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: dfcb1954cd9522de6165130cc115fddbb5f3ec45
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744211"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Novidades na API JavaScript do Word 1.1

O WordApi 1.1 é o primeiro conjunto de requisitos da API JavaScript do Word. É o único conjunto de requisitos de API do Word com suporte Word 2016.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.1. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.1, consulte APIs do Word no conjunto de requisitos [1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|Limpa o conteúdo do objeto Body.|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|Obtém a coleção de objetos rich text content control no corpo.|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|Obtém o formato de texto do corpo.|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|Obtém uma representação HTML do objeto body.|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|Obtém a representação OOXML (Office Open XML) do objeto Body.|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|Obtém a coleção de objetos InlinePicture no corpo.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Quebra o objeto Body com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|Insere um documento no corpo, no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|Insere HTML no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|Insere texto no corpo, no local especificado.|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|Obtém a coleção de objetos de parágrafo no corpo.|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|Obtém o controle de conteúdo que inclui o corpo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#word-word-body-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto body.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|Seleciona o corpo e navega na interface do usuário do Word até ele.|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|Obtém ou define o nome de estilo do corpo.|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|Obtém o texto do corpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|Obtém ou define a aparência do controle de conteúdo.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|Obtém ou define um valor que indica se o usuário pode excluir o controle de conteúdo.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|Obtém ou define um valor que indica se o usuário pode editar o conteúdo do controle de conteúdo.|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|Limpa o conteúdo do controle de conteúdo.|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|Obtém ou define a cor do controle de conteúdo.|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|Obtém a coleção de objetos de controle de conteúdo no controle de conteúdo.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|Exclui o controle de conteúdo e o respectivo conteúdo.|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|Obtém o formato de texto do controle de conteúdo.|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|Obtém uma representação HTML do objeto de controle de conteúdo.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|Obtém a representação OOXML (Office Open XML) do objeto do controle de conteúdo.|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|Obtém um número inteiro que representa o identificador do controle de conteúdo.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|Obtém a coleção de objetos inlinePicture no controle de conteúdo.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|Insere uma quebra no local especificado no documento principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|Insere um documento no controle de conteúdo no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|Insere HTML no local especificado dentro do controle de conteúdo.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|Insere o OOXML no controle de conteúdo no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|Insere texto no local especificado dentro do controle de conteúdo.|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|Obtém a coleção de objetos de parágrafo no controle de conteúdo.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|Obtém o controle de conteúdo que inclui o controle de conteúdo.|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|Obtém ou define o texto do espaço reservado do controle de conteúdo.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|Obtém ou define um valor que determina quando o controle de conteúdo é removido após a edição.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto de controle de conteúdo.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|Seleciona o controle de conteúdo.|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|Obtém ou define o nome de estilo do controle de conteúdo.|
||[marcar](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|Obtém ou define uma marca para identificar um controle de conteúdo.|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|Obtém o texto do controle de conteúdo.|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|Obtém ou define o título do controle de conteúdo.|
||[tipo](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|Obtém o tipo de controle de conteúdo.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|Obtém um controle de conteúdo pelo respectivo identificador.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|Obtém os controles de conteúdo com a marca especificada.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|Obtém os controles de conteúdo com o título especificado.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|Obtém um controle de conteúdo pelo índice na coleção.|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|Obtém o objeto body do documento principal.|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|Obtém a coleção de objetos de controle de conteúdo no documento.|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getselection-member(1))|Obtém a seleção atual do documento.|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|Salva o documento.|
||[saved](/javascript/api/word/word.document#word-word-document-saved-member)|Indica se as alterações do documento foram salvas.|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|Obtém a coleção de objetos de seção no documento.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|Obtém ou define um valor que indica se a fonte está em negrito.|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|Obtém ou define a cor da fonte especificada.|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|Obtém ou define um valor que indica se a fonte tem um tachado duplo.|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|Obtém ou define a cor de realçada.|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|Obtém ou define um valor que indica se a fonte está em itálico.|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|Obtém ou define um valor que representa o nome da fonte.|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|Obtém ou define um valor que representa o tamanho da fonte em pontos.|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|Obtém ou define um valor que indica se a fonte tem um tachado.|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|Obtém ou define um valor que indica se a fonte é um subscrito.|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|Obtém ou define um valor que indica se a fonte é um sobrescrito.|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|Obtém ou define um valor que indica o tipo de sublinhado da fonte.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|Obtém ou define uma cadeia de caracteres que representa o texto alternativo associado à imagem em linha.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|Obtém ou define uma cadeia de caracteres que inclui o título da imagem embutida.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|Obtém a representação de cadeia de caracteres codificada base64 da imagem embutda.|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|Obtém ou define um número que descreve a altura da imagem embutida.|
||[hiperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|Obtém ou define um hiperlink na imagem.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|Quebra a imagem embutida com um controle de conteúdo de rich text.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|Obtém ou define um valor que indica se a imagem embutida mantém as respectivas proporções originais quando você a redimensiona.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrol-member)|Obtém o controle de conteúdo que inclui a imagem embutida.|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|Obtém ou define um número que descreve a largura da imagem embutida.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|Obtém ou define o alinhamento de um parágrafo.|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|Limpa o conteúdo do objeto Paragraph.|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|Obtém a coleção de objetos de controle de conteúdo no parágrafo.|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|Exclui o parágrafo e o respectivo conteúdo do documento.|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|Retorna ou define o valor, em pontos, para um recuo deslocado ou da primeira linha.|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|Obtém o formato de texto do parágrafo.|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|Obtém uma representação HTML do objeto paragraph.|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|Obtém a representação OOXML (Office Open XML) do objeto Paragraph.|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|Obtém a coleção de objetos InlinePicture no parágrafo.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Quebra o objeto Paragraph com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|Insere um documento no parágrafo no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|Insere HTML no local especificado dentro do parágrafo.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|Insere uma imagem no local especificado dentro do parágrafo.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|Insere o OOXML no parágrafo no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|Insere texto no local especificado dentro do parágrafo.|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|Obtém ou define o valor de recuo à esquerda, em pontos, para o parágrafo.|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|Obtém ou define o espaçamento entre linhas, em pontos, para o parágrafo especificado.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|Obtém ou define a quantidade de espaçamento, em linhas de grade, após o parágrafo.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|Obtém ou define a quantidade de espaçamento, em linhas de grade, antes do parágrafo.|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|Obtém ou define o nível de estrutura de tópicos para o parágrafo.|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|Obtém o controle de conteúdo que inclui o parágrafo.|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|Obtém ou define o valor de recuo à direita, em pontos, para o parágrafo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto paragraph.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|Seleciona e navega na interface do usuário do Word até o parágrafo.|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|Obtém ou define o espaçamento, em pontos, após o parágrafo.|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|Obtém ou define o espaçamento, em pontos, antes o parágrafo.|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|Obtém ou define o nome de estilo do parágrafo.|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|Obtém o texto do parágrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|Limpa o conteúdo do objeto Range.|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|Obtém a coleção de objetos de controle de conteúdo no intervalo.|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|Exclui o intervalo e o respectivo conteúdo do documento.|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|Obtém o formato de texto do intervalo.|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|Obtém uma representação HTML do objeto range.|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|Obtém a representação OOXML do objeto Range.|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Quebra o objeto Range com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|Insere um documento no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|Insere HTML no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|Insere um texto no local especificado.|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|Obtém a coleção de objetos de parágrafo no intervalo.|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|Obtém o controle de conteúdo que inclui o intervalo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#word-word-range-search-member(1))|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto range.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|Seleciona e navega na interface do usuário do Word até o intervalo.|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|Obtém ou define o nome de estilo do intervalo.|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|Obtém o texto do intervalo.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|Obtém ou define um valor que determina quando ignorar todos os caracteres de pontuação entre as palavras.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|Obtém ou define um valor que indica se deve ignorar todo o espaço em branco entre palavras.|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|Obtém ou define um valor que determina quando realizar uma pesquisa que diferencia maiúsculas de minúsculas.|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que começam com a cadeia de caracteres da pesquisa.|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que terminam com a cadeia de caracteres da pesquisa.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|Obtém ou define um valor que determina quando a operação Localizar encontra apenas palavras inteiras, e não o texto que faz parte de uma palavra maior.|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|Obtém o objeto body da seção.|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|Obtém um dos rodapés da seção.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getheader-member(1))|Obtém um dos cabeçalhos da seção.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
