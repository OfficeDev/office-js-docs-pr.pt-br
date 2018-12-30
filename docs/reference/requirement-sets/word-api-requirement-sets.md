---
title: Conjuntos de requisitos da API JavaScript do Word
description: ''
ms.date: 11/14/2018
ms.openlocfilehash: 92678945f504e39bcf712558736145e948a8c628
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458045"
---
# <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Word são executados em várias versões do Office, incluindo Office 2016 ou posterior para Windows, Office para iPad, Office para Mac e Office Online. A tabela a seguir lista os conjuntos de requisitos do Word, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build desses aplicativos.

> [!NOTE]
> Para os conjuntos de requisitos marcados como Beta, use a versão especificada (ou posterior) do software do Office e a biblioteca Beta da CDN:https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Entradas não listadas como Beta estão geralmente disponíveis, e você pode continuar a usar a biblioteca Produção da CDN:https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Conjunto de requisitos  |   Office 365 para Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Versão 1612 (Compilação 7668.1000) ou posterior| Março de 2017, 2.22 ou posterior | Março de 2017, 15.32 ou posterior| Março de 2017 ||
| WordApi 1.2  | Atualização de dezembro de 2015, versão 1601 (build 6568.1000) ou posterior | Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 | |
| WordApi 1.1  | Versão 1509 (build 4266.1001) ou posterior| Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 | |

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém o conjunto de requisitos 1.1 de WordApi.

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="whats-new-in-word-javascript-api-13"></a>Quais são as novidades na API JavaScript do Word 1.3 

A seguir estão as novas adições às APIs JavaScript do Word no conjunto de requisitos 1.3. 

|Objeto| Novidades| Descrição|Conjunto de requisitos| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Método_ > createDocument(base64File: string) | Cria um novo documento usando um arquivo .docx codificado em base64. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > lists|Obtém a coleção de listas de objetos no corpo. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > parentBody|Obtém o corpo pai do corpo. Por exemplo, o corpo pai do corpo de uma célula de tabela poderia ser um cabeçalho. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > parentSection|Obtém a seção pai do corpo. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > styleBuiltIn|Obtém ou define o nome do estilo interno para o corpo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > tables|Obtém a coleção de tabelas de objetos no corpo. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Relação_ > type|Obtém o tipo do corpo. O tipo pode ser 'MainDoc', 'Section', 'Header', 'Footer' ou 'TableCell'. Somente leitura.|1.3|
|[body](/javascript/api/word/word.body)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtém o corpo todo, ou então, os pontos inicial ou final do corpo, como um intervalo.|1.3|
|[body](/javascript/api/word/word.body)|_Método_>insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Start' ou 'End'.|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Relação_ > quebras|Especifica a forma de uma quebra de: linha, a página ou tipo de seção. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > lists|Obtém a coleção de listas de objetos no controle de conteúdo. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > parentBody|Obtém o corpo pai do controle de conteúdo. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > parentTable|Obtém a tabela que contém o controle de conteúdo. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > parentTableCell|Obtém a célula de tabela que contém o controle de conteúdo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > styleBuiltIn|Obtém ou define o nome do estilo interno para o controle de conteúdo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > subtype|Obtém o subtipo de controle de conteúdo. O subtipo pode ser 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' e 'RichTextTable' para controles de conteúdo em rich text. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relação_ > tables|Obtém a coleção de objetos de tabela no controle de conteúdo. Somente leitura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtém o controle de todo o conteúdo, ou então, os pontos inicial ou final do controle de conteúdo, como um intervalo.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtém os intervalos de texto no controle de conteúdo usando marcas de pontuação e/ou outras marcas finais.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_>insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Insere uma tabela com a quantidade especificada de linhas e colunas dentro ou próxima do controle de conteúdo. O valor de insertLocation pode ser 'Start', 'End', 'Before' ou 'After'.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_> split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Divide o controle de conteúdo em intervalos filho usando delimitadores.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Método_getByTypes(types: ContentControlType)|Obtém os controles de conteúdo com os tipos e/ou subtipos especificados.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Método_ > getFirst()|Obtém o primeiro controle de conteúdo nesta coleção.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propriedade_ > key|Obtém a chave da propriedade personalizada. Somente leitura. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propriedade_ > value|Obtém ou define o valor da propriedade personalizada.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Relação_ > type|Obtém o tipo de valor da propriedade personalizada. Somente leitura.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Método_ > Delete()|Exclui a propriedade personalizada.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Propriedade_ > itens|Uma coleção de objetos customProperty. Somente leitura.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > deleteAll()|Exclui todas as propriedades personalizadas nesta coleção.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > getCount()|Obtém a contagem das propriedades personalizadas.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > getItem(key: string)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > set(key: string, value: object)|Cria ou define uma propriedade personalizada.|1.3|
|[document](/javascript/api/word/word.document)|_Relação_ > properties|Obtém as propriedades do documento atual. Somente leitura.|1.3|
|[documentCreated](/javascript/api/word/word.documentcreated)|_Método_ > open()|Abra o documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > applicationName|Obtém o nome do aplicativo do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > author|Obtém ou define o autor do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > category|Obtém ou define a categoria do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > comments|Obtém ou define os comentários do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > company|Obtém ou define a empresa do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > format|Obtém ou define o formato do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > keywords|Obtém ou define as palavras-chave do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > lastAuthor|Obtém ou define o último autor do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > manager|Obtém ou define o gerenciador do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > revisionNumber|Obtém o número de revisão do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > security|Obtém a segurança do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > subject|Obtém ou define o assunto do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > template|Obtém o modelo do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriedade_ > title|Obtém ou define o título do documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relação_ > creationDate|Obtém a data de criação do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relação_ > customProperties|Obtém a coleção de propriedades personalizadas do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relação_ > lastPrintDate|Obtém a data de impressão do documento. Somente leitura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relação_ > lastSaveTime|Obtém a hora em que o documento foi salvo pela última vez. Somente leitura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relação_ > parentTable|Obtém a tabela que contém a imagem embutida. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relação_ > parentTableCell|Obtém a célula de tabela que contém a imagem embutida. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > getNext()|Obtém a próxima imagem embutida.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtém a imagem, ou então, os pontos inicial ou final da imagem, como um intervalo.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Método_ > getFirst()|Obtém a primeira imagem embutida nesta coleção.|1.3|
|[list](/javascript/api/word/word.list)|_Propriedade_ > id|Obtém a id da lista. Somente leitura.|1.3|
|[list](/javascript/api/word/word.list)|_Propriedade_ > levelExistences|Verifica se cada um dos 9 níveis existe na lista. Um valor true indica que o nível existe, o que significa que há pelo menos um item de lista nesse nível. Somente leitura.|1.3|
|[list](/javascript/api/word/word.list)|_Relação_ > levelTypes|Obtém todos os tipos de nível 9 na lista. Cada tipo pode ser 'Bullet', 'Number' ou 'Picture'. Somente leitura.|1.3|
|[list](/javascript/api/word/word.list)|_Relação_ > paragraphs|Obtém parágrafos na lista. Somente leitura.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_getLevelParagraphs(level: number)|Obtém os parágrafos que ocorrem no nível especificado na lista.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_> getLevelString(level: number)|Obtém o marcador, o número ou a imagem no nível especificado como uma cadeia de caracteres.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_>insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Start', 'End', 'Before' ou 'After'.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_> setLevelAlignment(level: number, alignment: Alignment)|Define o alinhamento do marcador, o número ou a imagem no nível especificado na lista.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_> setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|Define o formato de marcador no nível especificado na lista. Se o marcador é 'Custom', o charCode é necessário.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|Define os dois recuos do nível especificado na lista.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object[])|Define o formato de numeração no nível especificado na lista.|1.3|
|[lista](/javascript/api/word/word.list)|_Método_> setLevelStartingNumber(level: number, startingNumber: number)|Define o número inicial no nível especificado na lista. O valor padrão é 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Propriedade_ > itens|Uma coleção de objetos de lista. Somente leitura.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_> getById(id: number)|Obtém uma lista por seu identificador.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_ > getFirst()|Obtém a primeira lista nesta coleção.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_ > getItem(index: number)|Obtém um objeto de lista por seu índice na coleção.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriedade_ > level|Obtém ou define o nível do item na lista.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriedade_ > listString|Obtém o marcador de item de lista, o número ou a imagem como uma cadeia de caracteres. Somente leitura.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriedade_ > siblingIndex|Obtém o número da ordem de item de lista em relação a seus irmãos. Somente leitura.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Método_ > getAncestor(parentOnly: bool)|Obtém o pai do item de lista ou o ancestral mais próximo se o pai não existir.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Método_ > getDescendants(directChildrenOnly: bool)|Obtém todos os itens de lista descendentes do item de lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriedade_ > isLastParagraph|Indica que o parágrafo é o último dentro do corpo do pai. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriedade_ > isListItem|Verifica se o parágrafo é um item da lista. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriedade_ > tableNestingLevel|Obtém o nível da tabela do parágrafo. Retorna 0 se o parágrafo não estiver em uma tabela. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > list|Obtém a lista à qual pertence esse parágrafo. Retorna um objeto nulo se o parágrafo não estiver em uma lista. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > listItem|Obtém o ListItem para o parágrafo. Retorna um objeto nulo se o parágrafo não fizer parte de uma lista. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > parentBody|Obtém o corpo pai do parágrafo. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > parentTable|Obtém a tabela que contém o parágrafo. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > parentTableCell|Obtém a célula de tabela que contém o parágrafo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relação_ > styleBuiltIn|Obtém ou define o nome do estilo interno para o parágrafo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > attachToList (listId: número nível: número)|Permite que o parágrafo ingresse em uma lista existente no nível especificado. Falhará se o parágrafo não puder ingressar na lista ou se o parágrafo já for um item da lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > detachFromList()|Move este parágrafo para fora de sua lista, caso o parágrafo seja um item da lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getNext()|Obtém o próximo parágrafo.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getPrevious()|Obtém o parágrafo anterior.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtém o parágrafo inteiro, ou então, os pontos inicial ou final do parágrafo, como um intervalo.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtém os intervalos de texto no parágrafo usando marcas de pontuação e/ou outras marcas finais.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_>insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)|Divide o parágrafo em intervalos filho usando delimitadores.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > startNewList()|Inicia uma nova lista com este parágrafo. Falhará se o parágrafo já for um item da lista.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Método_ > getFirst()|Obtém o primeiro parágrafo nesta coleção.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Método_ > getLast()|Obtém o último parágrafo nesta coleção.|1.3|
|[range](/javascript/api/word/word.range)|_Propriedade_ > hiperlink|Obtém o primeiro hiperlink no intervalo ou define um hiperlink no intervalo. Todos os hiperlinks no intervalo são excluídos quando você configura um novo hiperlink no intervalo. Use um caractere newline ('\n') para separar a parte de endereço da parte de local opcional.|1.3|
|[range](/javascript/api/word/word.range)|_Propriedade_ > isEmpty|Verifica se o comprimento do intervalo é zero. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > lists|Obtém a coleção de listas de objetos no intervalo. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > parentBody|Obtém o corpo pai do intervalo. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > parentTable|Obtém a tabela que contém o intervalo. Retorna nulo se não estiver contido em uma tabela. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > parentTableCell|Obtém a célula de tabela que contém o intervalo. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > styleBuiltIn|Obtém ou define o nome do estilo interno para o intervalo. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|1.3|
|[range](/javascript/api/word/word.range)|_Relação_ > tables|Obtém a coleção de tabelas de objetos no intervalo. Somente leitura.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > compareLocationWith(range: Range)|Compara o local deste intervalo com a localização de outro intervalo.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > expandTo(range: Range)|Retorna um novo intervalo que se estende a partir deste intervalo em qualquer direção para cobrir outro intervalo. Este intervalo não é alterado.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getHyperlinkRanges()|Obtém intervalos filho de hiperlink dentro do intervalo.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getNextTextRange (endingMarks: cadeia de caracteres, trimSpacing: bool)|Obtém o próximo intervalo de texto usando marcas de pontuação e/ou outras marcas finais.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getRange(rangeLocation: RangeLocation)|Clona o intervalo, ou então, obtém os pontos inicial ou final do intervalo como um novo intervalo.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtém os intervalos filho do texto no parágrafo usando marcas de pontuação e/ou outras marcas finais.|1.3|
|[range](/javascript/api/word/word.range)|_Método_>insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > intersectWith(range: Range)|Retorna um novo intervalo como ponto de interseção deste intervalo com outro intervalo. Este intervalo não é alterado.|1.3|
|[range](/javascript/api/word/word.range)|_Método_> split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Divide o intervalo em intervalos filho usando delimitadores.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Propriedade_ > itens|Uma coleção de objetos de intervalo. Somente leitura.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Método_ > getFirst()|Obtém o primeiro intervalo nesta coleção.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Método_ > getItem(index: number)|Obtém um objeto de intervalo por seu índice na coleção.|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Método_ > carregar (objeto: objeto, a opção: objeto)|Preenche o objeto proxy criado na camada JavaScript com a propriedade e as opções especificadas no parâmetro. |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Método_ > sync()|Envia a fila de solicitações para o Word e retorna um objeto Promise, que pode ser usado para o encadeamento de mais ações.|1.3|
|[section](/javascript/api/word/word.section)|_Método_ > getNext()|Obtém a próxima seção.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Método_ > getFirst()|Obtém a primeira seção nesta coleção.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > headerRowCount|Obtém e define o número de linhas de cabeçalho.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > height|Obtém a altura da tabela em pontos. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > isUniform|Indica se todas as linhas de tabela são uniformes. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > nestingLevel|Obtém o nível de aninhamento da tabela. Tabelas de nível superior têm o nível 1. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > rowCount|Obtém a quantidade de linhas na tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > shadingColor|Obtém e define a cor de sombreamento.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > style|Obtém ou define o nome do estilo usado para a tabela. Use esta propriedade de estilos personalizados e nomes de estilo localizados. Para usar os estilos internos que são portáteis entre localidades, confira a propriedade "styleBuiltIn".|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > styleBandedColumns|Obtém e define se a tabela tem colunas em tiras.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > styleBandedRows|Obtém e define se a tabela tem linhas em tiras.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > styleFirstColumn|Obtém e define se a tabela tem uma primeira coluna com um estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > styleLastColumn|Obtém e define se a tabela tem uma última coluna com um estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > styleTotalRow|Obtém e define se a tabela tem uma (última) linha total com um estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > values|Obtém e define os valores de texto na tabela, como uma matriz de Javascript 2D.|1.3|
|[table](/javascript/api/word/word.table)|_Propriedade_ > width|Obtém e define a largura da tabela em pontos.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > font|Obtém a fonte. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > horizontalAlignment|Obtém e define o alinhamento horizontal de cada célula na tabela. O valor pode ser 'left', 'centered', 'right' ou 'justified'.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > paragraphAfter|Obtém o parágrafo após a tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > paragraphBefore|Obtém o parágrafo antes da tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > parentBody|Obtém o corpo pai da tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > parentContentControl|Obtém o controle de conteúdo que contém a tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > parentTable|Obtém a tabela que contém esta tabela. Retorna um objeto nulo se não estiver contido em uma tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > parentTableCell|Obtém a célula de tabela que contém esta tabela. Retorna um objeto nulo se não estiver contido em uma célula de tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > rows|Obtém todas as linhas da tabela. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > styleBuiltIn|Obtém ou define o nome do estilo interno para a tabela. Use esta propriedade para estilos internos que são portáteis entre localidades. Para usar estilos personalizados ou nomes de estilo localizados, confira a propriedade "estilo".|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > tables|Obtém as tabelas filho aninhadas em um nível mais profundo. Somente leitura.|1.3|
|[table](/javascript/api/word/word.table)|_Relação_ > verticalAlignment|Obtém e define o alinhamento vertical de cada célula na tabela. O valor pode ser 'top', 'center' ou 'bottom'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_>addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Adiciona colunas ao início ou no final da tabela, usando a primeira ou última coluna existente como um modelo. Isto é aplicável às tabelas uniformes. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_> addRows(insertLocation: InsertLocation, rowCount: number, values: string)|Adiciona linhas ao início ou no final da tabela, usando a primeira ou última linha existente como um modelo. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > autoFitContents()|Autoajusta as colunas da tabela para a largura do seu conteúdo.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > autoFitWindow()|Autoajusta as colunas da tabela para a largura da janela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > clear()|Limpa o conteúdo da tabela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > Delete()|Exclui toda a tabela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > deleteColumns (columnIndex: número NúmeroDeColunas: número)|Exclui colunas específicas. Isto é aplicável às tabelas uniformes.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > deleteRows (rowIndex: número rowCount: número)|Exclui linha específicas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > distributeColumns()|Distribui uniformemente a largura das colunas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > distributeRows()|Distribui uniformemente a altura das linhas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_>getBorder(borderLocation: BorderLocation)|Obtém o estilo de borda para a borda especificada.|1.3|
|[table](/javascript/api/word/word.table)|_Método_getCell(rowIndex: number, cellIndex: number)|Obtém a célula da tabela em uma linha e coluna especificada.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtém o preenchimento de célula em pontos.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getNext()|Obtém a próxima tabela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtém o intervalo que contém esta tabela, ou o intervalo no início ou no final da tabela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > insertContentControl()|Insere um controle de conteúdo na tabela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_>insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_>insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Insere uma tabela com a quantidade especificada de linhas e colunas. O valor de insertLocation pode ser 'Before' ou 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Executa uma pesquisa com os searchOptions especificados no escopo do objeto de tabela. Os resultados da pesquisa são uma coleção de objetos de intervalo.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > select(selectionMode: SelectionMode)|Seleciona a tabela, ou então, a posição no início ou no final da tabela e navega na interface do usuário do Word até ela.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Define o preenchimento de célula em pontos.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propriedade_ > color|Obtém ou define a cor da borda da tabela, como um valor hexadecimal ou nome.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propriedade_ > width|Obtém ou define a largura, em pontos, da borda da tabela. Não aplicável a tipos de borda de tabela que têm larguras fixas.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Relação_ > type|Obtém ou define o tipo de borda da tabela.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > cellIndex|Obtém o índice da célula em sua linha. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > columnWidth|Obtém e define a largura da coluna da célula em pontos. Isto é aplicável às tabelas uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > rowIndex|Obtém o índice da linha da célula na tabela. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > shadingColor|Obtém ou define a cor de sombreamento da célula. Você pode definir a cor no formato "#RRGGBB" ou usando o nome da cor.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > value|Obtém e define o texto da célula.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriedade_ > width|Obtém a largura da célula em pontos. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relação_ > body|Obtém o objeto do corpo da célula. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relação_ > horizontalAlignment|Obtém e define o alinhamento horizontal da célula. O valor pode ser 'left', 'centered', 'right' ou 'justified'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relação_ > parentRow|Obtém a linha pai da célula. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relação_ > parentTable|Obtém a tabela pai da célula. Somente leitura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relação_ > verticalAlignment|Obtém e define o alinhamento vertical da célula. O valor pode ser 'top', 'center' ou 'bottom'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > deleteColumn()|Exclui a coluna que contém essa célula. Isto é aplicável às tabelas uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > deleteRow()|Exclui a linha que contém essa célula.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_>getBorder(borderLocation: BorderLocation)|Obtém o estilo de borda para a borda especificada.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtém o preenchimento de célula em pontos.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > getNext()|Obtém a próxima célula.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > insertColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Adiciona colunas à esquerda ou à direita da célula, usando a coluna da célula como um modelo. Isto é aplicável às tabelas uniformes. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_> insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|Insere linhas acima ou abaixo da célula, usando a linha da célula como um modelo. Os valores de cadeia de caracteres, se especificado, são definidos nas linhas recém-inseridas.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Define o preenchimento de célula em pontos.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Propriedade_ > itens|Uma coleção de objetos TableCell. Somente leitura.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Método_ > getFirst()|Obtém a primeira célula da tabela nesta coleção.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Método_ > getItem(index: number)|Obtém um objeto de célula de tabela pelo índice na coleção.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Propriedade_ > itens|Uma coleção de objetos de tabela. Somente leitura.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Método_ > getFirst()|Obtém a primeira tabela nesta coleção.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Método_ > getItem(index: number)|Obtém um objeto de tabela pelo índice na coleção.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > cellCount|Obtém a quantidade de células na linha. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > isHeader|Verifica se a linha é uma linha de cabeçalho. Somente leitura. Para definir o número de linhas de cabeçalho, use HeaderRowCount no objeto de tabela. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > preferredHeight|Obtém e define a altura da linha preferencial em pontos.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > rowIndex|Obtém o índice da linha em sua tabela pai. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > shadingColor|Obtém e define a cor de sombreamento.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriedade_ > values|Obtém e define os valores de texto na linha, como uma matriz de Javascript 1D.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relação_ > cells|Obtém células. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relação_ > font|Obtém a fonte. Use isto para obter e definir o nome, o tamanho e a cor da fonte, além de outras propriedades. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relação_ > horizontalAlignment|Obtém e define o alinhamento horizontal de cada célula na linha. O valor pode ser 'left', 'centered', 'right' ou 'justified'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relação_ > parentTable|Obtém uma tabela pai. Somente leitura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relação_ > verticalAlignment|Obtém e define o alinhamento vertical das células na linha. O valor pode ser 'top', 'center' ou 'bottom'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > clear()|Limpa o conteúdo da linha.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > Delete()|Exclui toda a linha.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_>getBorder(borderLocation: BorderLocation)|Obtém o estilo de borda das células na linha.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtém o preenchimento de célula em pontos.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > getNext()|Obtém a próxima linha.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_> insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|Insere linhas usando esta linha como um modelo. Se os valores forem especificados, insere os valores para as novas linhas.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Executa uma pesquisa com os searchOptions especificados no escopo da linha. Os resultados da pesquisa são uma coleção de objetos de intervalo.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > select(selectionMode: SelectionMode)|Seleciona a linha e navega na interface do usuário do Word até ele.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Define o preenchimento de célula em pontos.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Propriedade_ > itens|Uma coleção de objetos TableRow. Somente leitura.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Método_ > getFirst()|Obtém a primeira linha nesta coleção.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Método_ > getItem(index: number)|Obtém um objeto de linha de tabela por seu índice na coleção.|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Quais são as novidades na API JavaScript do Word 1.2

A seguir estão as novas adições às APIs JavaScript do Word no conjunto de requisitos 1.2. 

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Insere uma imagem embutida no local especificado dentro do controle de conteúdo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relação_ > paragraph|Obtém o parágrafo pai que inclui a imagem embutida. Somente leitura.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > Delete()|Exclui a imagem embutida do documento.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertBreak(breakType: BreakType, insertLocation: InsertLocation)|Insere uma quebra no local especificado no documento principal. O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|Insere um documento no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertHtml(html: string, insertLocation: InsertLocation)|Insere HTML no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Insere uma imagem embutida no local especificado. O valor de insertLocation pode ser 'Replace', 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertOoxml(ooxml: string, insertLocation: InsertLocation)|Insere um formato OOXML no local especificado.  O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_>insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_insertText(text: string, insertLocation: InsertLocation)|Insere um texto no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > select(selectionMode: SelectionMode)|Seleciona a imagem embutida. Isso faz com que o Word role até a seleção.|1.2|
|[range](/javascript/api/word/word.range)|_Relação_ > inlinePictures|Obtém a coleção de objetos de imagem embutida presentes no intervalo. Somente leitura.|1.2|
|[range](/javascript/api/word/word.range)|_Método_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Insere uma imagem no local especificado. O valor de insertLocation pode ser 'Replace', 'Start', 'End', 'Before' ou 'After'.|1.2|

## <a name="word-javascript-api-11"></a>API JavaScript do Word 1.1

A API JavaScript do Word 1.1 é a primeira versão da API. Para saber mais sobre a API, confira os tópicos de referência [API JavaScript do Word](/javascript/api/word). 

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
