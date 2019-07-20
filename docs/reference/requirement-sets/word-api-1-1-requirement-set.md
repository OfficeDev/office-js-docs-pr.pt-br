---
title: Conjunto de requisitos de API JavaScript do Word 1,1
description: Detalhes sobre o conjunto de requisitos WordApi 1,1
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 7c9ecbb8edaf1134b9f8801a6ade77b1b30e332f
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35805288"
---
# <a name="whats-new-in-word-javascript-api-11"></a>O que há de novo no Word JavaScript API 1,1

WordApi 1,1 é o primeiro conjunto de requisitos da API JavaScript do Word. É o único conjunto de requisitos da API do Word suportado pelo Word 2016.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs adicionadas como parte do conjunto de requisitos do WordApi 1,1.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Limpa o conteúdo do objeto Body. O usuário pode executar a operação de desfazer no conteúdo limpo.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Obtém uma representação HTML do objeto Body. Quando renderizado em uma página da Web ou em um visualizador de HTML, a formatação será uma correspondência próxima, mas não exata, à formatação do documento. Este método não retorna o mesmo HTML para o mesmo documento em diferentes plataformas (Windows, Mac, etc.). Se você precisar de fidelidade exata ou consistência entre plataformas, use `Body.getOoxml()` e converta o XML RETORNADO em HTML.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Obtém a representação OOXML (Office Open XML) do objeto Body.|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal. O valor de insertLocation pode ser 'Start' ou 'End'.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Quebra o objeto Body com um controle de conteúdo de rich text.|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no corpo, no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Insere HTML no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Insere um formato OOXML no local especificado.  O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Start' ou 'End'.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Insere texto no corpo, no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo de Rich Text no corpo. Somente leitura.|
||[font](/javascript/api/word/word.body#font)|Obtém o formato de texto do corpo. Use isso para obter e definir o nome, tamanho, cor e outras propriedades da fonte. Somente leitura.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Obtém a coleção de objetos InlinePicture no corpo. A coleção não inclui imagens flutuantes. Somente leitura.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtém a coleção de objetos Paragraph no corpo. Somente leitura.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Obtém o controle de conteúdo que inclui o corpo. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[text](/javascript/api/word/word.body#text)|Obtém o texto do corpo. Usa o método insertText para inserir texto. Somente leitura.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions)](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza uma pesquisa com o Searchoptions especificado no escopo do objeto Body. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Seleciona o corpo e navega na interface do usuário do Word até ele.|
||[style](/javascript/api/word/word.body#style)|Obtém ou define o nome do estilo usado para o corpo. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[apresentação](/javascript/api/word/word.contentcontrol#appearance)|Obtém ou define a aparência do controle de conteúdo. O valor pode ser "BoundingBox", "tags" ou "Hidden".|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|Obtém ou define um valor que indica se o usuário pode excluir o controle de conteúdo. Mutuamente exclusivo com a propriedade removeWhenEdited.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Obtém ou define um valor que indica se o usuário pode editar o conteúdo do controle de conteúdo.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Limpa o conteúdo do controle de conteúdo. O usuário pode executar a operação de desfazer no conteúdo limpo.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtém ou define a cor do controle de conteúdo. A cor é especificada no formato ' #RRGGBB ' ou usando o nome da cor.|
||[excluir (keepContent: Boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Exclui o controle de conteúdo e o respectivo conteúdo. Quando keepContent é definido como verdadeiro, o conteúdo não é excluído.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Obtém uma representação HTML do objeto de controle de conteúdo. Quando renderizado em uma página da Web ou em um visualizador de HTML, a formatação será uma correspondência próxima, mas não exata, à formatação do documento. Este método não retorna o mesmo HTML para o mesmo documento em diferentes plataformas (Windows, Mac, etc.). Se você precisar de fidelidade exata ou consistência entre plataformas, use `ContentControl.getOoxml()` e converta o XML RETORNADO em HTML.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Obtém a representação OOXML (Office Open XML) do objeto do controle de conteúdo.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal. O valor insertLocation pode ser ' Start ', ' End ', ' before ' ou ' after '. Este método não pode ser usado com os controles de conteúdo "RichTextTable", "RichTextTableRow" e "RichTextTableCell".|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no controle de conteúdo no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Insere HTML no local especificado dentro do controle de conteúdo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Insere OOXML no controle de conteúdo no local especificado.  O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor insertLocation pode ser ' Start ', ' End ', ' before ' ou ' after '.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Insere texto no local especificado dentro do controle de conteúdo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|Obtém ou define o texto do espaço reservado do controle de conteúdo. Quando o controle de conteúdo está vazio, o sistema exibe o texto esmaecido.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no controle de conteúdo. Somente leitura.|
||[font](/javascript/api/word/word.contentcontrol#font)|Obtém o formato de texto do controle de conteúdo. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtém um número inteiro que representa o identificador do controle de conteúdo. Somente leitura.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Obtém a coleção de objetos inlinePicture no controle de conteúdo. A coleção não inclui imagens flutuantes. Somente leitura.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtém a coleção de objetos Paragraph no controle de conteúdo. Somente leitura.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Obtém o controle de conteúdo que inclui o controle de conteúdo. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtém o texto do controle de conteúdo. Somente leitura.|
||[tipo](/javascript/api/word/word.contentcontrol#type)|Obtém o tipo de controle de conteúdo. Atualmente, temos suporte apenas para controles de conteúdo de rich text. Somente leitura.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Obtém ou define um valor que determina quando o controle de conteúdo é removido após a edição. Mutuamente exclusivo com a propriedade cannotDelete.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions)](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza uma pesquisa com o Searchoptions especificado no escopo do objeto de controle de conteúdo. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Seleciona o controle de conteúdo. Isso faz com que o Word role até a seleção.|
||[style](/javascript/api/word/word.contentcontrol#style)|Obtém ou define o nome do estilo do controle de conteúdo. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|
||[identificador](/javascript/api/word/word.contentcontrol#tag)|Obtém ou define uma marca para identificar um controle de conteúdo.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtém ou define o título do controle de conteúdo.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Obtém um controle de conteúdo pelo respectivo identificador. Gera se não há um controle de conteúdo com o identificador nesta coleção.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Obtém os controles de conteúdo com a marca especificada.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Obtém os controles de conteúdo com o título especificado.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Obtém um controle de conteúdo por seu índice na coleção.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[GetSelection ()](/javascript/api/word/word.document#getselection--)|Obtém a seleção atual do documento. Não há suporte para várias seleções.|
||[body](/javascript/api/word/word.document#body)|Obtém o objeto Body do documento. O corpo é o texto que exclui cabeçalhos, rodapés, notas de rodapé, caixas de texto, etc. Somente leitura.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no documento. Isso inclui controles de conteúdo no corpo do documento, cabeçalhos, rodapés, caixas de texto, etc. Somente leitura.|
||[salvo](/javascript/api/word/word.document#saved)|Indica se as alterações do documento foram salvas. Um valor true indica que o documento não foi alterado desde que foi salvo. Somente leitura.|
||[sections](/javascript/api/word/word.document#sections)|Obtém a coleção de objetos section no documento. Somente leitura.|
||[save()](/javascript/api/word/word.document#save--)|Salva o documento. Caso o documento não tenha sido salvo, ele usa a convenção de nomenclatura de arquivo padrão do Word.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtém ou define um valor que indica se a fonte está em negrito. True quando a fonte é formatada como negrito; caso contrário, false.|
||[color](/javascript/api/word/word.font#color)|Obtém ou define a cor da fonte especificada. Você pode fornecer o valor no formato ' #RRGGBB ' ou o nome da cor.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Obtém ou define um valor que indica se a fonte tem um tachado duplo. True quando a fonte é formatada como texto tachado duplo; caso contrário, false.|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|Obtém ou define a cor de realce. Para defini-lo, use um valor no formato ' #RRGGBB ' ou o nome da cor. Para remover a cor de realce, defina-a como NULL. A cor de realce retornada pode estar no formato ' #RRGGBB ', uma cadeia de caracteres vazia para cores de realce mistas ou NULL para nenhuma cor de realce.|
||[italic](/javascript/api/word/word.font#italic)|Obtém ou define um valor que indica se a fonte está em itálico. True quando a fonte está em itálico; caso contrário, false.|
||[name](/javascript/api/word/word.font#name)|Obtém ou define um valor que representa o nome da fonte.|
||[size](/javascript/api/word/word.font#size)|Obtém ou define um valor que representa o tamanho da fonte em pontos.|
||[Tachado](/javascript/api/word/word.font#strikethrough)|Obtém ou define um valor que indica se a fonte tem um tachado. True quando a fonte é formatada como texto tachado; caso contrário, false.|
||[subscript](/javascript/api/word/word.font#subscript)|Obtém ou define um valor que indica se a fonte é um subscrito. True quando a fonte é formatada como subscrito; caso contrário, false.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtém ou define um valor que indica se a fonte é um sobrescrito. True quando a fonte é formatada como sobrescrito; caso contrário, false.|
||[underline](/javascript/api/word/word.font#underline)|Obtém ou define um valor que indica o tipo de sublinhado da fonte. ' Nenhum ' se a fonte não estiver sublinhada.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Obtém ou define uma cadeia de caracteres que representa o texto alternativo associado à imagem embutida.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Obtém ou define uma cadeia de caracteres que inclui o título da imagem embutida.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Obtém a representação de cadeia de caracteres codificada em base64 da imagem embutida.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtém ou define um número que descreve a altura da imagem embutida.|
||[hiperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Obtém ou define um hiperlink na imagem. Use um ' # ' para separar a parte de endereço da parte de local opcional.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Quebra a imagem embutida com um controle de conteúdo de rich text.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Obtém ou define um valor que indica se a imagem embutida mantém as respectivas proporções originais quando você a redimensiona.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Obtém o controle de conteúdo que inclui a imagem embutida. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtém ou define um número que descreve a largura da imagem embutida.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Paragraph](/javascript/api/word/word.paragraph)|[Alignment](/javascript/api/word/word.paragraph#alignment)|Obtém ou define o alinhamento de um parágrafo. O valor pode ser 'left', 'centered', 'right' ou 'justified'.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Limpa o conteúdo do objeto Paragraph. O usuário pode executar a operação de desfazer no conteúdo limpo.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Exclui o parágrafo e o respectivo conteúdo do documento.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|Retorna ou define o valor, em pontos, para um recuo deslocado ou da primeira linha. Usa um valor positivo para definir um recuo da primeira linha e um valor negativo para definir um recuo deslocado.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Obtém uma representação HTML do objeto Paragraph. Quando renderizado em uma página da Web ou em um visualizador de HTML, a formatação será uma correspondência próxima, mas não exata, à formatação do documento. Este método não retorna o mesmo HTML para o mesmo documento em diferentes plataformas (Windows, Mac, etc.). Se você precisar de fidelidade exata ou consistência entre plataformas, use `Paragraph.getOoxml()` e converta o XML RETORNADO em HTML.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Obtém a representação OOXML (Office Open XML) do objeto Paragraph.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Quebra o objeto Paragraph com um controle de conteúdo de rich text.|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no parágrafo no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Insere HTML no local especificado dentro do parágrafo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem no local especificado dentro do parágrafo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Insere OOXML no parágrafo no local especificado. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Insere texto no local especificado dentro do parágrafo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|Obtém ou define o valor de recuo à esquerda, em pontos, para o parágrafo.|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|Obtém ou define o espaçamento entre linhas, em pontos, para o parágrafo especificado. Na interface do usuário do Word, esse valor é dividido por 12.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|Obtém ou define a quantidade de espaçamento, em linhas de grade, após o parágrafo.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|Obtém ou define a quantidade de espaçamento, em linhas de grade, antes do parágrafo.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|Obtém ou define o nível de estrutura de tópicos para o parágrafo.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no parágrafo. Somente leitura.|
||[font](/javascript/api/word/word.paragraph#font)|Obtém o formato de texto do parágrafo. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Obtém a coleção de objetos InlinePicture no parágrafo. A coleção não inclui imagens flutuantes. Somente leitura.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Obtém o controle de conteúdo que inclui o parágrafo. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[text](/javascript/api/word/word.paragraph#text)|Obtém o texto do parágrafo. Somente leitura.|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|Obtém ou define o valor de recuo à direita, em pontos, para o parágrafo.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza uma pesquisa com o Searchoptions especificado no escopo do objeto Paragraph. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Seleciona e navega na interface do usuário do Word até o parágrafo.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|Obtém ou define o espaçamento, em pontos, após o parágrafo.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Obtém ou define o espaçamento, em pontos, antes o parágrafo.|
||[style](/javascript/api/word/word.paragraph#style)|Obtém ou define o nome do estilo para o parágrafo. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Limpa o conteúdo do objeto Range. O usuário pode executar a operação de desfazer no conteúdo limpo.|
||[delete()](/javascript/api/word/word.range#delete--)|Exclui o intervalo e o respectivo conteúdo do documento.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Obtém uma representação HTML do objeto Range. Quando renderizado em uma página da Web ou em um visualizador de HTML, a formatação será uma correspondência próxima, mas não exata, à formatação do documento. Este método não retorna o mesmo HTML para o mesmo documento em diferentes plataformas (Windows, Mac, etc.). Se você precisar de fidelidade exata ou consistência entre plataformas, use `Range.getOoxml()` e converta o XML RETORNADO em HTML.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Obtém a representação OOXML do objeto Range.|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Quebra o objeto Range com um controle de conteúdo de rich text.|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no local especificado. O valor insertLocation pode ser ' replace ', ' Start ', ' End ', ' before ' ou ' after '.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Insere HTML no local especificado. O valor insertLocation pode ser ' replace ', ' Start ', ' End ', ' before ' ou ' after '.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Insere um formato OOXML no local especificado.  O valor insertLocation pode ser ' replace ', ' Start ', ' End ', ' before ' ou ' after '.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Insere um texto no local especificado. O valor insertLocation pode ser ' replace ', ' Start ', ' End ', ' before ' ou ' after '.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no intervalo. Somente leitura.|
||[font](/javascript/api/word/word.range#font)|Obtém o formato de texto do intervalo. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtém a coleção de objetos Paragraph no intervalo. Somente leitura.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Obtém o controle de conteúdo que inclui o intervalo. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[text](/javascript/api/word/word.range#text)|Obtém o texto do intervalo. Somente leitura.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions)](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza uma pesquisa com o Searchoptions especificado no escopo do objeto Range. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Seleciona e navega na interface do usuário do Word até o intervalo.|
||[style](/javascript/api/word/word.range#style)|Obtém ou define o nome do estilo para o intervalo. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|Obtém ou define um valor que determina quando ignorar todos os caracteres de pontuação entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de pontuação", na caixa de diálogo "Localizar e substituir".|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|Obtém ou define um valor que indica se deve ignorar todos os espaços em branco entre as palavras. Corresponde à caixa de seleção ignorar caracteres de espaço em branco na caixa de diálogo Localizar e substituir.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Obtém ou define um valor que determina quando realizar uma pesquisa que diferencia maiúsculas de minúsculas. Corresponde à caixa de seleção diferenciar maiúsculas de minúsculas da caixa de diálogo Localizar e substituir.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que começam com a cadeia de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir prefixo", na caixa de diálogo "Localizar e substituir".|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Obtém ou define um valor que determina quando fazer correspondência com as palavras que terminam com a cadeia de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir sufixo", na caixa de diálogo "Localizar e substituir".|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Obtém ou define um valor que determina quando a operação Localizar encontra apenas palavras inteiras, e não o texto que faz parte de uma palavra maior. Corresponde à caixa de seleção "Localizar apenas palavras inteiras", na caixa de diálogo "Localizar e substituir".|
||[matchWildCards](/javascript/api/word/word.searchoptions#matchwildcards)||
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais. Corresponde à caixa de seleção "Usar caracteres curinga", na caixa de diálogo "Localizar e substituir".|
|[Section](/javascript/api/word/word.section)|[getrodapé (tipo: Word. HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Obtém um dos rodapés da seção.|
||[GetHeader (tipo: Word. HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Obtém um dos cabeçalhos da seção.|
||[body](/javascript/api/word/word.section#body)|Obtém o objeto Body da seção. Isso não inclui o cabeçalho/rodapé e outros metadados da seção. Somente leitura.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
