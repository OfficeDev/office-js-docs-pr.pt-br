---
title: Conjunto de requisitos da API JavaScript do Word 1.1
description: Detalhes sobre o conjunto de requisitos do WordApi 1.1
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: c7b1adfa1af76f9994ced793dfddcf457cf733858fd27ba0ef763a67c35611c2
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092448"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Novidades na API JavaScript do Word 1.1

O WordApi 1.1 é o primeiro conjunto de requisitos da API JavaScript do Word. É o único conjunto de requisitos de API do Word com suporte Word 2016.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.1. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.1, consulte APIs do Word no conjunto de requisitos [1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|Limpa o conteúdo do objeto Body.|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|Obtém uma representação HTML do objeto body.|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|Obtém a representação OOXML (Office Open XML) do objeto Body.|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|Quebra o objeto Body com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|Insere um documento no corpo, no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|Insere HTML no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertText_text__insertLocation_)|Insere texto no corpo, no local especificado.|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[contentControls](/javascript/api/word/word.body#contentControls)|Obtém a coleção de objetos rich text content control no corpo.|
||[font](/javascript/api/word/word.body#font)|Obtém o formato de texto do corpo.|
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|Obtém a coleção de objetos InlinePicture no corpo.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtém a coleção de objetos de parágrafo no corpo.|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|Obtém o controle de conteúdo que inclui o corpo.|
||[text](/javascript/api/word/word.body#text)|Obtém o texto do corpo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto body.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select_selectionMode_)|Seleciona o corpo e navega na interface do usuário do Word até ele.|
||[style](/javascript/api/word/word.body#style)|Obtém ou define o nome de estilo do corpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#appearance)|Obtém ou define a aparência do controle de conteúdo.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|Obtém ou define um valor que indica se o usuário pode excluir o controle de conteúdo.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|Obtém ou define um valor que indica se o usuário pode editar o conteúdo do controle de conteúdo.|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|Limpa o conteúdo do controle de conteúdo.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtém ou define a cor do controle de conteúdo.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete_keepContent_)|Exclui o controle de conteúdo e o respectivo conteúdo.|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|Obtém uma representação HTML do objeto de controle de conteúdo.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|Obtém a representação OOXML (Office Open XML) do objeto do controle de conteúdo.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|Insere uma quebra no local especificado no documento principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|Insere um documento no controle de conteúdo no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|Insere HTML no local especificado dentro do controle de conteúdo.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|Insere o OOXML no controle de conteúdo no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|Insere texto no local especificado dentro do controle de conteúdo.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|Obtém ou define o texto do espaço reservado do controle de conteúdo.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|Obtém a coleção de objetos de controle de conteúdo no controle de conteúdo.|
||[font](/javascript/api/word/word.contentcontrol#font)|Obtém o formato de texto do controle de conteúdo.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtém um número inteiro que representa o identificador do controle de conteúdo.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|Obtém a coleção de objetos inlinePicture no controle de conteúdo.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtém a coleção de objetos Paragraph no controle de conteúdo.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|Obtém o controle de conteúdo que inclui o controle de conteúdo.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtém o texto do controle de conteúdo.|
||[type](/javascript/api/word/word.contentcontrol#type)|Obtém o tipo de controle de conteúdo.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|Obtém ou define um valor que determina quando o controle de conteúdo é removido após a edição.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto de controle de conteúdo.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select_selectionMode_)|Seleciona o controle de conteúdo.|
||[style](/javascript/api/word/word.contentcontrol#style)|Obtém ou define o nome de estilo do controle de conteúdo.|
||[marcar](/javascript/api/word/word.contentcontrol#tag)|Obtém ou define uma marca para identificar um controle de conteúdo.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtém ou define o título do controle de conteúdo.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|Obtém um controle de conteúdo pelo respectivo identificador.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|Obtém os controles de conteúdo com a marca especificada.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|Obtém os controles de conteúdo com o título especificado.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|Obtém um controle de conteúdo pelo índice na coleção.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[getSelection()](/javascript/api/word/word.document#getSelection__)|Obtém a seleção atual do documento.|
||[body](/javascript/api/word/word.document#body)|Obtém o objeto body do documento.|
||[contentControls](/javascript/api/word/word.document#contentControls)|Obtém a coleção de objetos de controle de conteúdo no documento.|
||[saved](/javascript/api/word/word.document#saved)|Indica se as alterações do documento foram salvas.|
||[sections](/javascript/api/word/word.document#sections)|Obtém a coleção de objetos de seção no documento.|
||[save()](/javascript/api/word/word.document#save__)|Salva o documento.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtém ou define um valor que indica se a fonte está em negrito.|
||[color](/javascript/api/word/word.font#color)|Obtém ou define a cor da fonte especificada.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|Obtém ou define um valor que indica se a fonte tem um tachado duplo.|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|Obtém ou define a cor de realçada.|
||[italic](/javascript/api/word/word.font#italic)|Obtém ou define um valor que indica se a fonte está em itálico.|
||[name](/javascript/api/word/word.font#name)|Obtém ou define um valor que representa o nome da fonte.|
||[size](/javascript/api/word/word.font#size)|Obtém ou define um valor que representa o tamanho da fonte em pontos.|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|Obtém ou define um valor que indica se a fonte tem um tachado.|
||[subscript](/javascript/api/word/word.font#subscript)|Obtém ou define um valor que indica se a fonte é um subscrito.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtém ou define um valor que indica se a fonte é um sobrescrito.|
||[underline](/javascript/api/word/word.font#underline)|Obtém ou define um valor que indica o tipo de sublinhado da fonte.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|Obtém ou define uma cadeia de caracteres que representa o texto alternativo associado à imagem em linha.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|Obtém ou define uma cadeia de caracteres que inclui o título da imagem embutida.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|Obtém a representação de cadeia de caracteres codificada base64 da imagem embutda.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtém ou define um número que descreve a altura da imagem embutida.|
||[hiperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Obtém ou define um hiperlink na imagem.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|Quebra a imagem embutida com um controle de conteúdo de rich text.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|Obtém ou define um valor que indica se a imagem embutida mantém as respectivas proporções originais quando você a redimensiona.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|Obtém o controle de conteúdo que inclui a imagem embutida.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtém ou define um número que descreve a largura da imagem embutida.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#alignment)|Obtém ou define o alinhamento de um parágrafo.|
||[clear()](/javascript/api/word/word.paragraph#clear__)|Limpa o conteúdo do objeto Paragraph.|
||[delete()](/javascript/api/word/word.paragraph#delete__)|Exclui o parágrafo e o respectivo conteúdo do documento.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|Retorna ou define o valor, em pontos, para um recuo deslocado ou da primeira linha.|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|Obtém uma representação HTML do objeto paragraph.|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|Obtém a representação OOXML (Office Open XML) do objeto Paragraph.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|Quebra o objeto Paragraph com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|Insere um documento no parágrafo no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|Insere HTML no local especificado dentro do parágrafo.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insere uma imagem no local especificado dentro do parágrafo.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|Insere o OOXML no parágrafo no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|Insere texto no local especificado dentro do parágrafo.|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|Obtém ou define o valor de recuo à esquerda, em pontos, para o parágrafo.|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|Obtém ou define o espaçamento entre linhas, em pontos, para o parágrafo especificado.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|Obtém ou define a quantidade de espaçamento, em linhas de grade, após o parágrafo.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|Obtém ou define a quantidade de espaçamento, em linhas de grade, antes do parágrafo.|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|Obtém ou define o nível de estrutura de tópicos para o parágrafo.|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|Obtém a coleção de objetos de controle de conteúdo no parágrafo.|
||[font](/javascript/api/word/word.paragraph#font)|Obtém o formato de texto do parágrafo.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|Obtém a coleção de objetos InlinePicture no parágrafo.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|Obtém o controle de conteúdo que inclui o parágrafo.|
||[text](/javascript/api/word/word.paragraph#text)|Obtém o texto do parágrafo.|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|Obtém ou define o valor de recuo à direita, em pontos, para o parágrafo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto paragraph.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select_selectionMode_)|Seleciona e navega na interface do usuário do Word até o parágrafo.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|Obtém ou define o espaçamento, em pontos, após o parágrafo.|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|Obtém ou define o espaçamento, em pontos, antes o parágrafo.|
||[style](/javascript/api/word/word.paragraph#style)|Obtém ou define o nome de estilo do parágrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|Limpa o conteúdo do objeto Range.|
||[delete()](/javascript/api/word/word.range#delete__)|Exclui o intervalo e o respectivo conteúdo do documento.|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|Obtém uma representação HTML do objeto range.|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|Obtém a representação OOXML do objeto Range.|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|Insere uma quebra no local especificado no documento principal.|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|Quebra o objeto Range com um controle de conteúdo de rich text.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|Insere um documento no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|Insere HTML no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertText_text__insertLocation_)|Insere um texto no local especificado.|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[contentControls](/javascript/api/word/word.range#contentControls)|Obtém a coleção de objetos de controle de conteúdo no intervalo.|
||[font](/javascript/api/word/word.range#font)|Obtém o formato de texto do intervalo.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtém a coleção de objetos de parágrafo no intervalo.|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|Obtém o controle de conteúdo que inclui o intervalo.|
||[text](/javascript/api/word/word.range#text)|Obtém o texto do intervalo.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Executa uma pesquisa com as SearchOptions especificadas no escopo do objeto range.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select_selectionMode_)|Seleciona e navega na interface do usuário do Word até o intervalo.|
||[style](/javascript/api/word/word.range#style)|Obtém ou define o nome de estilo do intervalo.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|Obtém ou define um valor que determina quando ignorar todos os caracteres de pontuação entre as palavras.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|Obtém ou define um valor que indica se deve ignorar todo o espaço em branco entre palavras.|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|Obtém ou define um valor que determina quando realizar uma pesquisa que diferencia maiúsculas de minúsculas.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que começam com a cadeia de caracteres da pesquisa.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que terminam com a cadeia de caracteres da pesquisa.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|Obtém ou define um valor que determina quando a operação Localizar encontra apenas palavras inteiras, e não o texto que faz parte de uma palavra maior.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais.|
|[Section](/javascript/api/word/word.section)|[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getFooter_type_)|Obtém um dos rodapés da seção.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getHeader_type_)|Obtém um dos cabeçalhos da seção.|
||[body](/javascript/api/word/word.section#body)|Obtém o objeto body da seção.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
