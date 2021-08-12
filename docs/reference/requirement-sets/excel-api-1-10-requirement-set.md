---
title: Excel Conjunto de requisitos da API JavaScript 1.10
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 34c21ad0e90593352ae4042c2be148e607c63164aac1845357e9f96371104f6f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087204"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Novidades na API JavaScript 1.10 Excel JavaScript

O ExcelApi 1.10 introduziu os principais recursos, como comentários, contornos e slicers. Ele também adicionou suporte a eventos para clique e classificação no nível da planilha.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Comments](../../excel/excel-add-ins-comments.md) | Adicione, edite e exclua comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Linhas e colunas de grupo para formar contornos retrálíveis. | [Intervalo,](/javascript/api/excel/excel.range) [Planilha](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insira e configure as segmentações de dados em tabelas e Tabelas dinâmicas. | [Segmentação de dados](/javascript/api/excel/excel.slicer) |
| [Mais eventos de planilha](../../excel/excel-add-ins-events.md) | Ouça clique e classificar eventos na planilha. | [Planilha (Eventos)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.10. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.10 ou anterior, consulte Excel APIs no conjunto de requisitos [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|O conteúdo do comentário.|
||[delete()](/javascript/api/excel/excel.comment#delete__)|Exclui o comentário e todas as respostas conectadas.|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|Obtém a célula onde este comentário está localizado.|
||[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.comment#authorName)|Obtém o nome do autor do comentário.|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|Obtém o horário de criação do comentário.|
||[id](/javascript/api/excel/excel.comment#id)|Especifica o identificador de comentário.|
||[replies](/javascript/api/excel/excel.comment#replies)|Representa uma coleção de objetos de resposta associados ao comentário.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Cadeia \| de caracteres de intervalo, conteúdo: cadeia de caracteres, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Cria um novo comentário com o conteúdo fornecido na célula especificada.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|Obtém o número de comentários na coleção.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|Obtém um comentário da coleção com base em seu ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|Obtém um comentário da coleção com base em sua posição.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|Obtém o comentário da célula especificada.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|Obtém o comentário ao qual a resposta dada está conectada.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|O conteúdo da resposta ao comentário.|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|Exclui a resposta do comentário. |
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|Obtém a célula onde esta resposta de comentário está localizada.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|Obtém o comentário pai desta resposta.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|Obtém o nome do autor da resposta do comentário.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|Obtém o horário de criação da resposta do comentário.|
||[id](/javascript/api/excel/excel.commentreply#id)|Especifica o identificador de resposta de comentário.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Cria uma resposta de comentário para um comentário.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|Obtém o número de respostas de comentários na coleção.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|Retorna uma resposta de comentário identificada pela respectiva ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|Obtém uma resposta de comentário com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|Exclui o estilo de tabela dinâmica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|Cria uma duplicata desse estilo de tabela dinâmica com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtém o nome do estilo de tabela dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|Especifica se esse `PivotTableStyle` objeto é somente leitura.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|Cria um em `PivotTableStyle` branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|Obtém o número de estilos de PivotTable na coleção.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|Obtém o estilo de tabela dinâmica padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|Obtém `PivotTableStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|Obtém `PivotTableStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault (newDefaultStyle: PivotTableStyle \| cadeia de caracteres)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|Define o estilo de tabela dinâmica padrão para uso no escopo do objeto pai.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#group_groupOption_)|Grupos colunas e linhas para um contorno.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|Oculta os detalhes da linha ou do grupo de colunas.|
||[height](/javascript/api/excel/excel.range#height)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo.|
||[left](/javascript/api/excel/excel.range#left)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha até a borda esquerda do intervalo.|
||[top](/javascript/api/excel/excel.range#top)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo.|
||[width](/javascript/api/excel/excel.range#width)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|Mostra os detalhes da linha ou do grupo de colunas.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup_groupOption_)|Desagrupa colunas e linhas para um contorno.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|Copia e colará um `Shape` objeto.|
||[placement](/javascript/api/excel/excel.shape#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Representa a legenda da slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|Limpa todos os filtros aplicados à segmentação de dados no momento.|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|Exclui a segmentação de dados.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|Retorna uma matriz de chaves de itens selecionados.|
||[height](/javascript/api/excel/excel.slicer#height)|Representa a altura, em pontos, da segmentação de dados.|
||[left](/javascript/api/excel/excel.slicer#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicer#name)|Representa o nome da slicer.|
||[id](/javascript/api/excel/excel.slicer#id)|Representa a ID exclusiva da slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|O valor `true` é se todos os filtros atualmente aplicados à slicer são limpos.|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|Representa a coleção de itens de slicer que fazem parte da slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Representa a planilha que contém a segmentação de dados.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|Seleciona itens de slicer com base em suas chaves.|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|Representa a ordem de classificação dos itens na segmentação de dados.|
||[style](/javascript/api/excel/excel.slicer#style)|Valor constante que representa o estilo da slicer.|
||[top](/javascript/api/excel/excel.slicer#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicer#width)|Representa a largura, em pontos, da segmentação de dados.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|Adiciona uma nova segmentação de dados à pasta de trabalho.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|Retorna o número de segmentações de dados na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|Obtém um objeto slicer usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|Obtém uma segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|Obtém uma slicer usando seu nome ou ID.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|O valor `true` será se o item da slicer estiver selecionado.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasData)|O valor `true` é se o item da slicer tiver dados.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Representa o título exibido na interface Excel interface do usuário.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|Retorna o número de itens da segmentação de dados na segmentação de dados.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|Obtém um objeto de item da segmentação de dados usando sua chave ou nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|Obtém um item da segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|Obtém um item da segmentação de dados usando sua chave ou nome.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|Exclui o estilo da slicer.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|Cria uma duplicata desse estilo de slicer com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtém o nome do estilo da slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|Especifica se esse `SlicerStyle` objeto é somente leitura.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|Cria um estilo de slicer em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|Obtém o número de segmentação de estilos na coleção.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|Obtém o `SlicerStyle` padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|Obtém `SlicerStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|Obtém `SlicerStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|Define o estilo de slicer padrão para uso no escopo do objeto pai.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|Exclui o estilo da tabela.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|Cria uma duplicata desse estilo de tabela com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtém o nome do estilo da tabela.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|Especifica se esse `TableStyle` objeto é somente leitura.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|Cria um em `TableStyle` branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|Obtém o número de estilos de tabelas na coleção.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|Obtém o estilo de tabela padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|Obtém `TableStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|Obtém `TableStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|Define o estilo de tabela padrão para uso no escopo do objeto pai.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|Exclui o estilo da tabela.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|Cria uma duplicata desse estilo de linha do tempo com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtém o nome do estilo da linha do tempo.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|Especifica se esse `TimelineStyle` objeto é somente leitura.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|Cria um em `TimelineStyle` branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|Obtém o número de estilos de linha do tempo na coleção.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|Obtém o estilo de linha do tempo padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|Obtém `TimelineStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|Obtém `TimelineStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|Define o estilo de linha do tempo padrão para uso no escopo do objeto pai.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|Obtém a segmentação de dados ativa no momento na pasta de trabalho.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|Obtém a segmentação de dados ativa no momento na pasta de trabalho.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Representa uma coleção de comentários associados à workbook.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|Representa uma coleção de SlicerStyles associados à pasta de trabalho.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Representa uma coleção de slicers associadas à workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|Representa uma coleção de TableStyles associadas à pasta de trabalho.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|Representa uma coleção de TimelineStyles associados à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Retorna um conjunto de todos os objetos Comments na planilha.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|Ocorre quando uma ou mais colunas são classificadas.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|Ocorre quando uma ou mais linhas são classificadas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|Ocorre quando uma ação clicada à esquerda/mapeada ocorre na planilha.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Retorna uma coleção de slicers que fazem parte da planilha.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|Mostra grupos de linhas ou colunas por seus níveis de contorno.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|Ocorre quando uma ou mais colunas são classificadas.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|Ocorre quando uma ou mais linhas são classificadas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|Ocorre quando a operação clicada à esquerda/mapeada ocorre na coleção de planilhas.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|Obtém a ID da planilha onde a classificação aconteceu.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|Obtém a ID da planilha onde a classificação aconteceu.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtém o endereço que representa a célula que foi clicada/tocada para uma planilha específica.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|A distância, em pontos, do ponto de grade clicado/mapeado para a esquerda (ou direita para idiomas da direita para a esquerda) da célula clicada/mapeada à esquerda.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|A distância, em pontos, desde o ponto clicado/tocado com o botão esquerdo até a borda da linha de grade superior da célula clicada/tocada com o botão esquerdo.|
||[tipo](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|Obtém a ID da planilha na qual a célula foi clicada à esquerda/tapped.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)