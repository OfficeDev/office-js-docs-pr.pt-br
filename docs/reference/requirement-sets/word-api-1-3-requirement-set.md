---
title: Conjunto de requisitos de API JavaScript do Word 1,3
description: Detalhes sobre o conjunto de requisitos WordApi 1,3
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: fe72a3047fdbdd719fd115858e4010fbc2c639e5
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268555"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Quais são as novidades na API JavaScript do Word 1.3

WordApi 1,3 adicionou mais suporte para controles de conteúdo, XML personalizado e configurações de nível de documento.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos de API JavaScript do Word, 1,3. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Word 1,3 ou anterior, confira [APIs do Word no conjunto de requisitos 1,3 ou anterior](/javascript/api/word?view=word-js-1.3).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/word/word.application)|[CreateDocument (base64file?: cadeia de caracteres)](/javascript/api/word/word.application#createdocument-base64file-)|Cria um novo documento usando um arquivo. docx codificado em base64 opcional.|
|[Body](/javascript/api/word/word.body)|[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Obtém o corpo todo, ou então, os pontos inicial ou final do corpo, como um intervalo.|
||[InsertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Start' ou 'End'.|
||[listas](/javascript/api/word/word.body#lists)|Obtém a coleção de listas de objetos no corpo. Somente leitura.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Obtém o corpo pai do corpo. Por exemplo, o corpo pai do corpo de uma célula de tabela poderia ser um cabeçalho. Gera uma exceção se não há um corpo pai. Somente leitura.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Obtém o corpo pai do corpo. Por exemplo, o corpo pai do corpo de uma célula de tabela poderia ser um cabeçalho. Retorna um objeto nulo se não há um corpo pai. Somente leitura.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o corpo. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Obtém a seção pai do corpo. Gera se não há uma seção pai. Somente leitura.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Obtém a seção pai do corpo. Retorna um objeto NULL se não houver uma seção pai. Somente leitura.|
||[tables](/javascript/api/word/word.body#tables)|Obtém a coleção de tabelas de objetos no corpo. Somente leitura.|
||[tipo](/javascript/api/word/word.body#type)|Obtém o tipo do corpo. O tipo pode ser 'MainDoc', 'Section', 'Header', 'Footer' ou 'TableCell'. Somente leitura.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Obtém ou define o nome do estilo interno para o corpo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Obtém o controle de todo o conteúdo, ou então, os pontos inicial ou final do controle de conteúdo, como um intervalo.|
||[gettextranges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos de texto no controle de conteúdo usando marcas de Pontuação e/ou outras marcas de fim.|
||[InsertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas dentro ou próxima do controle de conteúdo. O valor insertLocation pode ser ' Start ', ' End ', ' before ' ou ' after '.|
||[listas](/javascript/api/word/word.contentcontrol#lists)|Obtém a coleção de listas de objetos no controle de conteúdo. Somente leitura.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Obtém o corpo pai do controle de conteúdo. Somente leitura.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o controle de conteúdo. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Obtém a tabela que contém o controle de conteúdo. Lança se não está contida em uma tabela. Somente leitura.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Obtém a célula de tabela que contém o controle de conteúdo. Lança se não está contida em uma célula de tabela. Somente leitura.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Obtém a célula de tabela que contém o controle de conteúdo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Obtém a tabela que contém o controle de conteúdo. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|
||[subtipo](/javascript/api/word/word.contentcontrol#subtype)|Obtém o subtipo de controle de conteúdo. O subtipo pode ser 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' e 'RichTextTable' para controles de conteúdo em rich text. Somente leitura.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Obtém a coleção de objetos de tabela no controle de conteúdo. Somente leitura.|
||[Split (Delimiters: String [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide o controle de conteúdo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Obtém ou define o nome do estilo interno para o controle de conteúdo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Obtém um controle de conteúdo pelo respectivo identificador. Retorna um objeto NULL se não houver um controle de conteúdo com o identificador nessa coleção.|
||[Métodogetbytypes (tipos: Word. ContentControltype [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Obtém os controles de conteúdo que têm os tipos e/ou subtipos especificados.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Obtém o primeiro controle de conteúdo nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Obtém o primeiro controle de conteúdo nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/word/word.customproperty#key)|Obtém a chave da propriedade personalizada. Somente leitura.|
||[tipo](/javascript/api/word/word.customproperty#type)|Obtém o tipo de valor da propriedade personalizada. Os valores possíveis são: String, Number, Date, Boolean. Somente leitura.|
||[value](/javascript/api/word/word.customproperty#value)|Obtém ou define o valor da propriedade personalizada. Observe que, embora o Word na Web e o formato de arquivo DOCX permitam que essas propriedades sejam arbitrariamente Long, a versão de área de trabalho do Word truncará valores de cadeia de caracteres para caracteres de 255 16 bits (possivelmente criando um Unicode inválido ao quebrar um par substituto).|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[Add (Key: String, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Lança se a propriedade personalizada não existe.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto NULL se a propriedade personalizada não existir.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtém as propriedades do documento. Somente leitura.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Abre o documento.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtém o objeto Body do documento. O corpo é o texto que exclui cabeçalhos, rodapés, notas de rodapé, caixas de texto, etc. Somente leitura.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Obtém a coleção de objetos de controle de conteúdo no documento. Isso inclui controles de conteúdo no corpo do documento, cabeçalhos, rodapés, caixas de texto, etc. Somente leitura.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtém as propriedades do documento. Somente leitura.|
||[salvo](/javascript/api/word/word.documentcreated#saved)|Indica se as alterações do documento foram salvas. Um valor true indica que o documento não foi alterado desde que foi salvo. Somente leitura.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtém a coleção de objetos section no documento. Somente leitura.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Salva o documento. Caso o documento não tenha sido salvo, ele usa a convenção de nomenclatura de arquivo padrão do Word.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[autor](/javascript/api/word/word.documentproperties#author)|Obtém ou define o autor do documento.|
||[Categorias](/javascript/api/word/word.documentproperties#category)|Obtém ou define a categoria do documento.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtém ou define os comentários do documento.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtém ou define a empresa do documento.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtém ou define o formato do documento.|
||[Palavras-chave](/javascript/api/word/word.documentproperties#keywords)|Obtém ou define as palavras-chave do documento.|
||[Gerenciador](/javascript/api/word/word.documentproperties#manager)|Obtém ou define o gerenciador do documento.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Obtém o nome do aplicativo do documento. Somente leitura.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Obtém a data de criação do documento. Somente leitura.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Obtém a coleção de propriedades personalizadas do documento. Somente leitura.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Obtém o último autor do documento. Somente leitura.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Obtém a data de impressão do documento. Somente leitura.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Obtém a hora em que o documento foi salvo pela última vez. Somente leitura.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Obtém o número de revisão do documento. Somente leitura.|
||[segurança](/javascript/api/word/word.documentproperties#security)|Obtém a segurança do documento. Somente leitura.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtém o modelo do documento. Somente leitura.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtém ou define o assunto do documento.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtém ou define o título do documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Obtém a próxima imagem embutida. Lança se esta imagem embutida é a última.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Obtém a próxima imagem embutida. Retorna um objeto NULL se esta imagem embutida for a última.|
||[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Obtém a imagem, ou então, os pontos inicial ou final da imagem, como um intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui a imagem embutida. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Obtém a tabela que contém a imagem embutida. Lança se não está contida em uma tabela. Somente leitura.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Obtém a célula de tabela que contém a imagem embutida. Lança se não está contida em uma célula de tabela. Somente leitura.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Obtém a célula de tabela que contém a imagem embutida. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Obtém a tabela que contém a imagem embutida. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Obtém a primeira imagem embutida nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Obtém a primeira imagem embutida nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
|[List](/javascript/api/word/word.list)|[Métodogetlevelparagraphs (Level: Number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Obtém os parágrafos que ocorrem no nível especificado na lista.|
||[getlevelstring (Level: Number)](/javascript/api/word/word.list#getlevelstring-level-)|Obtém o marcador, o número ou a imagem no nível especificado como uma cadeia de caracteres.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor insertLocation pode ser ' Start ', ' End ', ' before ' ou ' after '.|
||[id](/javascript/api/word/word.list#id)|Obtém a ID da lista.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Verifica se cada um dos 9 níveis existe na lista. Um valor true indica que o nível existe, o que significa que há pelo menos um item de lista nesse nível. Somente leitura.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Obtém todos os tipos de nível 9 na lista. Cada tipo pode ser ' marcador ', ' número ' ou ' imagem '. Somente leitura.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtém parágrafos na lista. Somente leitura.|
||[setLevelAlignment (Level: Number, Alignment: Word. Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Define o alinhamento do marcador, o número ou a imagem no nível especificado na lista.|
||[setLevelBullet (Level: Number, listBullet: Word. ListBullet, charco?: Number, NomeDaFonte?: String)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Define o formato de marcador no nível especificado na lista. Se o marcador é 'Custom', o charCode é necessário.|
||[setLevelIndents (Level: Number, TextIndent: Number, bulletNumberPictureIndent: Number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Define os dois recuos do nível especificado na lista.|
||[setLevelNumbering (Level: Number, listNumbering: Word. ListNumbering, formatString?: matriz<número \| de cadeia de caracteres>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Define o formato de numeração no nível especificado na lista.|
||[setLevelStartingNumber (Level: Number, startingNumber: Number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Define o número inicial no nível especificado na lista. O valor padrão é 1.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Obtém uma lista por seu identificador. Lança se não há uma lista com o identificador nessa coleção.|
||[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Obtém uma lista por seu identificador. Retorna um objeto NULL se não houver uma lista com o identificador nessa coleção.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Obtém a primeira lista nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Obtém a primeira lista nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Obtém um objeto de lista por seu índice na coleção.|
||[items](/javascript/api/word/word.listcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ListItem](/javascript/api/word/word.listitem)|[GetAncestor (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir. Lança se o item de lista não tem ancestral.|
||[getAncestorOrNullObject (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir. Retorna um objeto NULL se o item de lista não tiver ancestral.|
||[getdescendants (directChildrenOnly?: Boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Obtém todos os itens de lista descendentes do item de lista.|
||[level](/javascript/api/word/word.listitem#level)|Obtém ou define o nível do item na lista.|
||[listString](/javascript/api/word/word.listitem#liststring)|Obtém o marcador, o número ou a imagem do item de lista como uma cadeia de caracteres. Somente leitura.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Obtém o número da ordem de item de lista em relação a seus irmãos. Somente leitura.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (ListId: Number, Level: Number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Permite que o parágrafo ingresse em uma lista existente no nível especificado. Falhará se o parágrafo não puder ingressar na lista ou se o parágrafo já for um item da lista.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Move este parágrafo para fora de sua lista, caso o parágrafo seja um item da lista.|
||[getNext ()](/javascript/api/word/word.paragraph#getnext--)|Obtém o próximo parágrafo. Lança se o parágrafo é o último.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Obtém o próximo parágrafo. Retorna um objeto NULL se o parágrafo for o último.|
||[getprevious ()](/javascript/api/word/word.paragraph#getprevious--)|Obtém o parágrafo anterior. Lança se o parágrafo é o primeiro.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Obtém o parágrafo anterior. Retorna um objeto NULL se o parágrafo for o primeiro.|
||[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Obtém o parágrafo inteiro, ou então, os pontos inicial ou final do parágrafo, como um intervalo.|
||[gettextranges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos de texto no parágrafo usando marcas de Pontuação e/ou outras marcas de fim.|
||[InsertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indica que o parágrafo é o último dentro do corpo do pai. Somente leitura.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Verifica se o parágrafo é um item da lista. Somente leitura.|
||[list](/javascript/api/word/word.paragraph#list)|Obtém a lista à qual pertence esse parágrafo. Lança se o parágrafo não está em uma lista. Somente leitura.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Obtém o ListItem para o parágrafo. Lança se o parágrafo não faz parte de uma lista. Somente leitura.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Obtém o ListItem para o parágrafo. Retorna um objeto nulo se o parágrafo não fizer parte de uma lista. Somente leitura.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Obtém a lista à qual pertence esse parágrafo. Retorna um objeto nulo se o parágrafo não estiver em uma lista. Somente leitura.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Obtém o corpo pai do parágrafo. Somente leitura.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o parágrafo. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Obtém a tabela que contém o parágrafo. Lança se não está contida em uma tabela. Somente leitura.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Obtém a célula de tabela que contém o parágrafo. Lança se não está contida em uma célula de tabela. Somente leitura.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Obtém a célula de tabela que contém o parágrafo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Obtém a tabela que contém o parágrafo. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Obtém o nível da tabela do parágrafo. Retorna 0 se o parágrafo não estiver em uma tabela. Somente leitura.|
||[Split (Delimiters: String [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Divide o parágrafo em intervalos filho usando delimitadores.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Inicia uma nova lista com este parágrafo. Falhará se o parágrafo já for um item da lista.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Obtém ou define o nome do estilo interno para o parágrafo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Obtém o primeiro parágrafo nesta coleção. Lança se a coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Obtém o primeiro parágrafo nesta coleção. Retorna um objeto NULL se a coleção estiver vazia.|
||[GetLast ()](/javascript/api/word/word.paragraphcollection#getlast--)|Obtém o último parágrafo nesta coleção. Lança se a coleção está vazia.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Obtém o último parágrafo nesta coleção. Retorna um objeto NULL se a coleção estiver vazia.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compara o local deste intervalo com a localização de outro intervalo.|
||[expandto (intervalo: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo. Este intervalo não é alterado. Gera se os dois intervalos não têm uma União.|
||[expandToOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo. Este intervalo não é alterado. Retorna um objeto NULL se os dois intervalos não tiverem uma União.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Obtém intervalos filho de hiperlink dentro do intervalo.|
||[getNextTextRange (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Obtém o próximo intervalo de texto usando marcas de Pontuação e/ou outras marcas de fim. Lança se este intervalo de texto é o último.|
||[getNextTextRangeOrNullObject (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Obtém o próximo intervalo de texto usando marcas de Pontuação e/ou outras marcas de fim. Retorna um objeto NULL se o intervalo de texto for o último.|
||[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clona o intervalo, ou então, obtém os pontos inicial ou final do intervalo como um novo intervalo.|
||[gettextranges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Obtém os intervalos de texto filhos no intervalo usando marcas de Pontuação e/ou outras marcas de fim.|
||[hiperlink](/javascript/api/word/word.range#hyperlink)|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo. Todos os hiperlinks no intervalo são excluídos quando você configura um novo hiperlink no intervalo. Use um ' # ' para separar a parte de endereço da parte de local opcional.|
||[InsertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[intersectWith (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo. Este intervalo não é alterado. Lança se os dois intervalos não são sobrepostos ou adjacentes.|
||[intersectWithOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo. Este intervalo não é alterado. Retorna um objeto NULL se os dois intervalos não forem sobrepostos ou adjacentes.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Verifica se o comprimento do intervalo é zero. Somente leitura.|
||[listas](/javascript/api/word/word.range#lists)|Obtém a coleção de listas de objetos no intervalo. Somente leitura.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Obtém o corpo pai do intervalo. Somente leitura.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que inclui o intervalo. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Obtém a tabela que contém o intervalo. Lança se não está contida em uma tabela. Somente leitura.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Obtém a célula de tabela que contém o intervalo. Lança se não está contida em uma célula de tabela. Somente leitura.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Obtém a célula de tabela que contém o intervalo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Obtém a tabela que contém o intervalo. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|
||[tables](/javascript/api/word/word.range#tables)|Obtém a coleção de tabelas de objetos no intervalo. Somente leitura.|
||[Split (Delimiters: String [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide o intervalo em intervalos filho usando delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Obtém ou define o nome do estilo interno para o intervalo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Obtém o primeiro intervalo nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Obtém o primeiro intervalo nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
|[Section](/javascript/api/word/word.section)|[getNext ()](/javascript/api/word/word.section#getnext--)|Obtém a próxima seção. Lança se esta seção é a última.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Obtém a próxima seção. Retorna um objeto NULL se esta seção for a última.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Obtém a primeira seção nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Obtém a primeira seção nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
|[Table](/javascript/api/word/word.table)|[AddColumns (insertLocation: Word. InsertLocation, columnCount: Number, Values?: String [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Adiciona colunas ao início ou no final da tabela, usando a primeira ou última coluna existente como um modelo. Isto é aplicável às tabelas uniformes. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|
||[AddRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Adiciona linhas ao início ou no final da tabela, usando a primeira ou última linha existente como um modelo. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|
||[Alignment](/javascript/api/word/word.table#alignment)|Obtém ou define o alinhamento da tabela em relação à coluna da página. O valor pode ser ' left ', ' centered ' ou ' right '.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Autoajusta as colunas da tabela para a largura da janela.|
||[clear()](/javascript/api/word/word.table#clear--)|Limpa o conteúdo da tabela.|
||[delete()](/javascript/api/word/word.table#delete--)|Exclui toda a tabela.|
||[deleteColumns (columnIndex: Number, columnCount?: Number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Exclui colunas específicas. Isto é aplicável às tabelas uniformes.|
||[deleteRows (rowIndex: Number, rowCount?: Number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Exclui linha específicas.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Distribui uniformemente a largura das colunas. Isto é aplicável às tabelas uniformes.|
||[GetBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Obtém o estilo de borda para a borda especificada.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Obtém a célula da tabela em uma linha e coluna especificada. Lança se a célula da tabela especificada não existe.|
||[getCellOrNullObject (rowIndex: Number, cellIndex: Number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Obtém a célula da tabela em uma linha e coluna especificada. Retorna um objeto NULL se a célula da tabela especificada não existir.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext ()](/javascript/api/word/word.table#getnext--)|Obtém a próxima tabela. Lança se esta tabela é a última.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Obtém a próxima tabela. Retorna um objeto NULL se esta tabela for a última.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Obtém o parágrafo após a tabela. Gera se não há parágrafo após a tabela.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Obtém o parágrafo após a tabela. Retorna um objeto NULL se não houver um parágrafo após a tabela.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Obtém o parágrafo antes da tabela. Gera se não há parágrafo antes da tabela.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Obtém o parágrafo antes da tabela. Retorna um objeto NULL se não houver um parágrafo antes da tabela.|
||[GetRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Obtém o intervalo que contém esta tabela, ou o intervalo no início ou no final da tabela.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Obtém e define o número de linhas de cabeçalho.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Obtém e define o alinhamento horizontal de cada célula na tabela. O valor pode ser ' left ', ' centered ', ' right ' ou ' justificado '.|
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Insere um controle de conteúdo na tabela.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[InsertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[font](/javascript/api/word/word.table#font)|Obtém a fonte. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indica se todas as linhas de tabela são uniformes. Somente leitura.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Obtém o nível de aninhamento da tabela. Tabelas de nível superior têm o nível 1. Somente leitura.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Obtém o corpo pai da tabela. Somente leitura.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Obtém o controle de conteúdo que contém a tabela. Gera se não há um controle de conteúdo pai. Somente leitura.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Obtém o controle de conteúdo que contém a tabela. Retorna um objeto NULL se não houver um controle de conteúdo pai. Somente leitura.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Obtém a tabela que contém esta tabela. Lança se não está contida em uma tabela. Somente leitura.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Obtém a célula de tabela que contém esta tabela. Lança se não está contida em uma célula de tabela. Somente leitura.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Obtém a célula de tabela que contém esta tabela. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Obtém a tabela que contém esta tabela. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|
||[Validação](/javascript/api/word/word.table#rowcount)|Obtém a quantidade de linhas na tabela. Somente leitura.|
||[rows](/javascript/api/word/word.table#rows)|Obtém todas as linhas da tabela. Somente leitura.|
||[tables](/javascript/api/word/word.table#tables)|Obtém as tabelas filho aninhadas em um nível mais profundo. Somente leitura.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions](/javascript/api/word/word.table#search-searchtext--searchoptions-)|Realiza uma pesquisa com o Searchoptions especificado no escopo do objeto Table. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Seleciona a tabela, ou então, a posição no início ou no final da tabela e navega na interface do usuário do Word até ela.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Obtém e define a cor de sombreamento. Você pode definir a cor no formato "#RRGGBB" ou usando o nome da cor.|
||[style](/javascript/api/word/word.table#style)|Obtém ou define o nome do estilo usado para a tabela. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Obtém e define se a tabela tem colunas em tiras.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Obtém e define se a tabela tem linhas em tiras.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Obtém ou define o nome do estilo interno para a tabela. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Obtém e define se a tabela tem uma primeira coluna com um estilo especial.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Obtém e define se a tabela tem uma última coluna com um estilo especial.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Obtém e define se a tabela tem uma (última) linha total com um estilo especial.|
||[values](/javascript/api/word/word.table#values)|Obtém e define os valores de texto na tabela, como uma matriz de Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Obtém e define o alinhamento vertical de cada célula na tabela. O valor pode ser "Top", "Center" ou "Bottom".|
||[width](/javascript/api/word/word.table#width)|Obtém e define a largura da tabela em pontos.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtém ou define a cor da borda da tabela.|
||[tipo](/javascript/api/word/word.tableborder#type)|Obtém ou define o tipo de borda da tabela.|
||[width](/javascript/api/word/word.tableborder#width)|Obtém ou define a largura, em pontos, da borda da tabela. Não aplicável a tipos de borda de tabela que têm larguras fixas.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Obtém e define a largura da coluna da célula em pontos. Isto é aplicável às tabelas uniformes.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Exclui a coluna que contém essa célula. Isto é aplicável às tabelas uniformes.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Exclui a linha que contém essa célula.|
||[GetBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Obtém o estilo de borda para a borda especificada.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext ()](/javascript/api/word/word.tablecell#getnext--)|Obtém a próxima célula. Lança se esta célula é a última.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Obtém a próxima célula. Retorna um objeto NULL se esta célula for a última.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Obtém e define o alinhamento horizontal da célula. O valor pode ser ' left ', ' centered ', ' right ' ou ' justificado '.|
||[insertColumns (insertLocation: Word. InsertLocation, columnCount: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Adiciona colunas à esquerda ou à direita da célula, usando a coluna da célula como um modelo. Isto é aplicável às tabelas uniformes. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Insere linhas acima ou abaixo da célula, usando a linha da célula como um modelo. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|
||[body](/javascript/api/word/word.tablecell#body)|Obtém o objeto do corpo da célula. Somente leitura.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Obtém o índice da célula em sua linha. Somente leitura.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Obtém a linha pai da célula. Somente leitura.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Obtém a tabela pai da célula. Somente leitura.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Obtém o índice da linha da célula na tabela. Somente leitura.|
||[width](/javascript/api/word/word.tablecell#width)|Obtém a largura da célula em pontos. Somente leitura.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Obtém ou define a cor de sombreamento da célula. Você pode definir a cor no formato "#RRGGBB" ou usando o nome da cor.|
||[value](/javascript/api/word/word.tablecell#value)|Obtém e define o texto da célula.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Obtém e define o alinhamento vertical da célula. O valor pode ser "Top", "Center" ou "Bottom".|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Obtém a primeira célula da tabela nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Obtém a primeira célula da tabela nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Obtém a primeira tabela nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Obtém a primeira tabela nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Limpa o conteúdo da linha.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Exclui toda a linha.|
||[GetBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Obtém o estilo de borda das células na linha.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Obtém o preenchimento de célula em pontos.|
||[getNext ()](/javascript/api/word/word.tablerow#getnext--)|Obtém a próxima linha. Lança se esta linha é a última.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Obtém a próxima linha. Retorna um objeto NULL se esta linha for a última.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Obtém e define o alinhamento horizontal de cada célula na linha. O valor pode ser ' left ', ' centered ', ' right ' ou ' justificado '.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Insere linhas usando esta linha como um modelo. Se os valores forem especificados, insere os valores para as novas linhas.|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Obtém e define a altura da linha preferencial em pontos.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Obtém a quantidade de células na linha. Somente leitura.|
||[nas](/javascript/api/word/word.tablerow#cells)|Obtém células. Somente leitura.|
||[font](/javascript/api/word/word.tablerow#font)|Obtém a fonte. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Verifica se a linha é uma linha de cabeçalho. Somente leitura. Para definir o número de linhas de cabeçalho, use HeaderRowCount no objeto de tabela.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Obtém uma tabela pai. Somente leitura.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Obtém o índice da linha em sua tabela pai. Somente leitura.|
||[Search (ProcurarTexto: String, searchoptions?: Word. Searchoptions)](/javascript/api/word/word.tablerow#search-searchtext--searchoptions-)|Realiza uma pesquisa com o Searchoptions especificado no escopo da linha. Os resultados da pesquisa são uma coleção de objetos Range.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Seleciona a linha e navega na interface do usuário do Word até ele.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Define o preenchimento de célula em pontos.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Obtém e define a cor de sombreamento. Você pode definir a cor no formato "#RRGGBB" ou usando o nome da cor.|
||[values](/javascript/api/word/word.tablerow#values)|Obtém e define os valores de texto na linha, como uma matriz JavaScript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Obtém e define o alinhamento vertical das células na linha. O valor pode ser "Top", "Center" ou "Bottom".|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Obtém a primeira linha nesta coleção. Lança se esta coleção está vazia.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Obtém a primeira linha nesta coleção. Retorna um objeto NULL se essa coleção estiver vazia.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
