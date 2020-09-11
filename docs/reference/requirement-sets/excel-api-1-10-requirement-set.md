---
title: Conjunto de requisitos de API JavaScript do Excel 1,10
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,10
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 890d198f238e29d39744d87d754381543ebcaf6a
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431231"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>O que há de novo na API JavaScript do Excel 1,10

O ExcelApi 1,10 introduziu os principais recursos, como comentários, contornos e Segmentações de tópicos. Ele também adicionou suporte a eventos para clicar e classificar em nível de planilha.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Comments](../../excel/excel-add-ins-comments.md) | Adicione, edite e exclua comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Descreve](../../excel/excel-add-ins-ranges-advanced.md#group-data-for-an-outline) | Agrupar linhas e colunas para formar contornos recolhíveis. | [Intervalo](/javascript/api/excel/excel.range), [planilha](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#slicers) | Insira e configure as segmentações de dados em tabelas e Tabelas dinâmicas. | [Segmentação de dados](/javascript/api/excel/excel.slicer) |
| [Mais eventos de planilha](../../excel/excel-add-ins-events.md) | Ouvir eventos Click e Sort na planilha. | [Planilha (eventos)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,10. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,10 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,10 ou anterior](/javascript/api/excel?view=excel-js-1.10&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Exclui o comentário e todas as respostas conectadas.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtém a célula em que este comentário está localizado.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtém o nome do autor do comentário.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtém o horário de criação do comentário. Retorna null se o comentário foi convertido de uma nota, pois o comentário não possui uma data de criação.|
||[id](/javascript/api/excel/excel.comment#id)|Representa o identificador de comentário. Somente leitura.|
||[replies](/javascript/api/excel/excel.comment#replies)|Representa uma coleção de objetos de resposta associados ao comentário. Somente leitura.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress: \| String de intervalo, Content: \| cadeia de caracteres CommentRichContent, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Cria um novo comentário com o conteúdo fornecido na célula especificada. Um `InvalidArgument` erro será acionado se o intervalo fornecido for maior que uma célula.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtém o número de comentários na coleção.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtém um comentário da coleção com base em seu ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtém um comentário da coleção com base em sua posição.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtém o comentário da célula especificada.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtém o comentário ao qual a resposta fornecida está conectada.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtém ou define o conteúdo da resposta do comentário. A cadeia de caracteres é de texto sem formatação.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Exclui a resposta do comentário. |
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtém a célula em que esta resposta de comentário está localizada.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtém o comentário pai desta resposta.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtém o nome do autor da resposta do comentário.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtém o horário de criação da resposta do comentário.|
||[id](/javascript/api/excel/excel.commentreply#id)|Representa o identificador de resposta do comentário. Somente leitura.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content: CommentRichContent \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtém o número de respostas de comentários na coleção.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Retorna uma resposta de comentário identificada pela respectiva ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtém uma resposta de comentário com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)||[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Exclui a Tabela Dinâmica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Cria uma duplicata desta Tabela Dinâmica com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtém o nome da Tabela Dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Especifica se este objeto PivotTable é somente leitura. Somente leitura.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Cria uma Tabela Dinâmica em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtém o número de estilos de PivotTable na coleção.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtém a Tabela Dinâmica padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Obtém um PivotTableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Obtém um PivotTableStyle por nome. Se PivotTableStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault (newDefaultStyle: PivotTableStyle \| cadeia de caracteres)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Define a Tabela Dinâmica padrão para uso no escopo do objeto pai.|
|[Range](/javascript/api/excel/excel.range)|[Grupo (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Agrupa colunas e linhas de uma estrutura de tópicos.|
||[hideGroupDetails (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Ocultar detalhes do grupo de linhas ou colunas.|
||[height](/javascript/api/excel/excel.range#height)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo. Somente leitura.|
||[left](/javascript/api/excel/excel.range#left)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha para a borda esquerda do intervalo. Somente leitura.|
||[top](/javascript/api/excel/excel.range#top)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo. Somente leitura.|
||[width](/javascript/api/excel/excel.range#width)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo. Somente leitura.|
||[showGroupDetails (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Mostrar detalhes do grupo de linhas ou colunas.|
||[Desagrupar (groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Desagrupa colunas e linhas de uma estrutura de tópicos.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copia e cola um objeto Forma.|
||[placement](/javascript/api/excel/excel.shape#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Representa a legenda da segmentação de dados.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Limpa todos os filtros aplicados à segmentação de dados no momento.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Exclui a segmentação de dados.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Retorna uma matriz de chaves de itens selecionados. Somente leitura.|
||[height](/javascript/api/excel/excel.slicer#height)|Representa a altura, em pontos, da segmentação de dados.|
||[left](/javascript/api/excel/excel.slicer#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicer#name)|Representa o nome da segmentação de dados.|
||[id](/javascript/api/excel/excel.slicer#id)|Representa a id exclusiva da segmentação de dados. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|Verdadeiro se todos os filtros atualmente aplicados na segmentação de dados estiverem desmarcados.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Representa a coleção de SlicerItems que faz parte da segmentação de dados. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Representa a planilha que contém a segmentação de dados. Somente leitura.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Seleciona itens de segmentação de itens com base em suas chaves. As seleções anteriores estão desmarcadas.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Os valores possíveis são: "DataSourceOrder", "crescente", "decrescente".|
||[style](/javascript/api/excel/excel.slicer#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicer#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicer#width)|Representa a largura, em pontos, da segmentação de dados.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adiciona uma nova segmentação de dados à pasta de trabalho.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Retorna o número de segmentações de dados na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtém um objeto de segmentação de dados usando seu respectivo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtém uma segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtém uma segmentação de dados usando seu nome ou id. Se a ela não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True se o item da segmentação de dados estiver selecionado.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True se o item de segmentação de dados tiver dados.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Representa o título exibido na interface do usuário.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Retorna o número de itens da segmentação de dados na segmentação de dados.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtém um objeto de item da segmentação de dados usando sua chave ou nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtém um item da segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtém um item da segmentação de dados usando sua chave ou nome. Se o item da segmentação de dados não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Exclui o SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Cria uma duplicata deste SlicerStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtém o nome o SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Especifica se este objeto SlicerStyle é somente leitura. Somente leitura.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Cria um SlicerStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtém o número de segmentação de estilos na coleção.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtém o padrão SlicerStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtém uma SlicerStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtém uma SlicerStyle por nome. Se o SlicerStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Define o padrão SlicerStyle para uso no escopo do objeto pai.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Cria uma duplicata deste TableStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Especifica se este objeto TableStyle é somente leitura. Somente leitura.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Cria um TableStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtém o número de estilos de tabelas na coleção.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtém o padrão TableStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Obtém um TableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Obtém um TableStyle por nome. Se o TableStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Define a TableStyle padrão para uso no escopo do objeto pai..|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Cria uma duplicata deste TimelineStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtém o nome do TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Especifica se este objeto timelinestyle é somente leitura. Somente leitura.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Cria um TimelineStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtém o número de estilos de linha do tempo na coleção.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtém o padrão TimelineStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtém uma TimelineStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtém uma TimelineStyle por nome. Se o TimelineStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Define o padrão TimelineStyle para uso no escopo do objeto pai.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver um slicer ativo, uma `ItemNotFound` exceção será lançada.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver segmentação de dados ativa, um objeto nulo será retornado.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Representa uma coleção de comentários associados à pasta de trabalho. Somente leitura.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Representa uma coleção de SlicerStyles associados à pasta de trabalho. Somente leitura.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Representa uma coleção de segmentações de dados associados à pasta de trabalho. Somente leitura.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Representa uma coleção de TableStyles associadas à pasta de trabalho. Somente leitura.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Representa uma coleção de TimelineStyles associados à pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Retorna um conjunto de todos os objetos Comments na planilha. Somente leitura.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Ocorre quando uma ou mais colunas são classificadas. Isso acontece como resultado de uma operação de classificação da esquerda para a direita.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Ocorre quando uma ou mais linhas são classificadas. Isso ocorre como resultado de uma operação de classificação de cima para baixo.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Ocorre quando uma ação com clique/tocado à esquerda acontece na planilha. Esse evento não será acionado quando você clicar nos seguintes casos:|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Retorna uma coleção de slicers que fazem parte da planilha. Somente leitura.|
||[showOutlineLevels (translevels: Number, columnLevels: Number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Mostra grupos de linhas ou colunas por seus níveis de estrutura de tópicos.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Ocorre quando uma ou mais colunas são classificadas. Isso acontece como resultado de uma operação de classificação da esquerda para a direita.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Ocorre quando uma ou mais linhas são classificadas. Isso ocorre como resultado de uma operação de classificação de cima para baixo.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Ocorre quando a operação com o botão esquerdo/tocado acontece na coleção de planilhas. Esse evento não será acionado quando você clicar nos seguintes casos:|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica. Somente colunas alteradas como resultado da operação de classificação são retornadas.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtém o id da planilha onde a classificação aconteceu.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica. Somente as linhas alteradas como resultado da operação de classificação são retornadas.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtém o id da planilha onde a classificação aconteceu.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtém o endereço que representa a célula que foi clicada/tocada para uma planilha específica.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|A distância, em pontos, do ponto de clique com o botão esquerdo/tocado à esquerda (ou à direita para idiomas da direita para a esquerda) da linha de grade da célula com clique à esquerda/tocado.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|A distância, em pontos, desde o ponto clicado/tocado com o botão esquerdo até a borda da linha de grade superior da célula clicada/tocada com o botão esquerdo.|
||[tipo](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtém o id da planilha na qual a célula foi clicada com o botão esquerdo/tocada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)