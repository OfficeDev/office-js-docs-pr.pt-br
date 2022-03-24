---
title: Excel conjunto de requisitos da API JavaScript 1.10
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 53cf0ec55a26f02a615a3c5eee0b718b818790d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746339"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Novidades na API JavaScript 1.10 Excel JavaScript

O ExcelApi 1.10 introduziu os principais recursos, como comentários, contornos e slicers. Ele também adicionou suporte a eventos para clique e classificação no nível da planilha.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Comments](../../excel/excel-add-ins-comments.md) | Adicione, edite e exclua comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Linhas e colunas de grupo para formar contornos retrálíveis. | [Intervalo](/javascript/api/excel/excel.range), [Planilha](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insira e configure as segmentações de dados em tabelas e Tabelas dinâmicas. | [Segmentação de dados](/javascript/api/excel/excel.slicer) |
| [Mais eventos de planilha](../../excel/excel-add-ins-events.md) | Ouça clique e classificar eventos na planilha. | [Planilha (Eventos)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.10. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.10 ou anterior, consulte Excel APIs no conjunto de requisitos [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true) ou anterior.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|Obtém o nome do autor do comentário.|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|O conteúdo do comentário.|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|Obtém o horário de criação do comentário.|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|Exclui o comentário e todas as respostas conectadas.|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|Obtém a célula onde este comentário está localizado.|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|Especifica o identificador de comentário.|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|Representa uma coleção de objetos de resposta associados ao comentário.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Cadeia de caracteres \| de intervalo, conteúdo: cadeia de caracteres, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Cria um novo comentário com o conteúdo fornecido na célula especificada.|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|Obtém o número de comentários na coleção.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|Obtém um comentário da coleção com base em seu ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|Obtém um comentário da coleção com base em sua posição.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|Obtém o comentário da célula especificada.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|Obtém o comentário ao qual a resposta dada está conectada.|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|Obtém o nome do autor da resposta do comentário.|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|O conteúdo da resposta ao comentário.|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|Obtém o horário de criação da resposta do comentário.|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|Exclui a resposta do comentário. |
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|Obtém a célula onde esta resposta de comentário está localizada.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|Obtém o comentário pai desta resposta.|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|Especifica o identificador de resposta de comentário.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Cria uma resposta de comentário para um comentário.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|Obtém o número de respostas de comentários na coleção.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|Retorna uma resposta de comentário identificada pela respectiva ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|Obtém uma resposta de comentário com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|Exclui o estilo de tabela dinâmica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|Cria uma duplicata desse estilo de tabela dinâmica com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|Obtém o nome do estilo de tabela dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|Especifica se esse objeto `PivotTableStyle` é somente leitura.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|Cria um em branco `PivotTableStyle` com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|Obtém o número de estilos de PivotTable na coleção.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|Obtém o estilo de tabela dinâmica padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|Obtém `PivotTableStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|Obtém `PivotTableStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault (newDefaultStyle: PivotTableStyle \| cadeia de caracteres)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|Define o estilo de tabela dinâmica padrão para uso no escopo do objeto pai.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|Grupos colunas e linhas para um contorno.|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|Oculta os detalhes da linha ou do grupo de colunas.|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha até a borda esquerda do intervalo.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|Mostra os detalhes da linha ou do grupo de colunas.|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|Desagrupa colunas e linhas para um contorno.|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|Copia e colará um `Shape` objeto.|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|Representa como o objeto é anexado às células abaixo dela.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|Representa a legenda da slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|Limpa todos os filtros aplicados à segmentação de dados no momento.|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|Exclui a segmentação de dados.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|Retorna uma matriz de chaves de itens selecionados.|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|Representa a altura, em pontos, da segmentação de dados.|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|Representa a ID exclusiva da slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|O valor é `true` se todos os filtros atualmente aplicados à slicer são limpos.|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|Representa o nome da slicer.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|Seleciona itens de slicer com base em suas chaves.|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|Representa a coleção de itens de slicer que fazem parte da slicer.|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|Representa a ordem de classificação dos itens na segmentação de dados.|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|Valor constante que representa o estilo da slicer.|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|Representa a largura, em pontos, da segmentação de dados.|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|Representa a planilha que contém a segmentação de dados.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|Adiciona uma nova segmentação de dados à pasta de trabalho.|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|Retorna o número de segmentações de dados na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|Obtém um objeto slicer usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|Obtém uma segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|Obtém uma slicer usando seu nome ou ID.|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|O valor é `true` se o item da slicer tiver dados.|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|O valor será `true` se o item da slicer estiver selecionado.|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|Representa o título exibido na interface Excel interface do usuário.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|Retorna o número de itens da segmentação de dados na segmentação de dados.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|Obtém um objeto de item da segmentação de dados usando sua chave ou nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|Obtém um item da segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|Obtém um item da segmentação de dados usando sua chave ou nome.|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|Exclui o estilo da slicer.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|Cria uma duplicata desse estilo de slicer com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|Obtém o nome do estilo da slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|Especifica se esse objeto `SlicerStyle` é somente leitura.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|Cria um estilo de slicer em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|Obtém o número de segmentação de estilos na coleção.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|Obtém o padrão `SlicerStyle` para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|Obtém `SlicerStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|Obtém `SlicerStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|Define o estilo de slicer padrão para uso no escopo do objeto pai.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|Exclui o estilo da tabela.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|Cria uma duplicata desse estilo de tabela com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|Obtém o nome do estilo da tabela.|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|Especifica se esse objeto `TableStyle` é somente leitura.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|Cria um em branco `TableStyle` com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|Obtém o número de estilos de tabelas na coleção.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|Obtém o estilo de tabela padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|Obtém `TableStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|Obtém `TableStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|Define o estilo de tabela padrão para uso no escopo do objeto pai.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|Exclui o estilo da tabela.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|Cria uma duplicata desse estilo de linha do tempo com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|Obtém o nome do estilo da linha do tempo.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|Especifica se esse objeto `TimelineStyle` é somente leitura.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|Cria um em branco `TimelineStyle` com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|Obtém o número de estilos de linha do tempo na coleção.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|Obtém o estilo de linha do tempo padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|Obtém `TimelineStyle` um pelo nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|Obtém `TimelineStyle` um pelo nome.|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|Define o estilo de linha do tempo padrão para uso no escopo do objeto pai.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|Representa uma coleção de comentários associados à workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|Obtém a segmentação de dados ativa no momento na pasta de trabalho.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|Obtém a segmentação de dados ativa no momento na pasta de trabalho.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho.|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|Representa uma coleção de SlicerStyles associados à pasta de trabalho.|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|Representa uma coleção de slicers associadas à workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|Representa uma coleção de TableStyles associadas à pasta de trabalho.|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|Representa uma coleção de TimelineStyles associados à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|Retorna um conjunto de todos os objetos Comments na planilha.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Ocorre quando uma ou mais colunas são classificadas.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|Ocorre quando uma ou mais linhas são classificadas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|Ocorre quando uma ação clicada à esquerda/mapeada ocorre na planilha.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|Mostra grupos de linhas ou colunas por seus níveis de contorno.|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|Retorna uma coleção de slicers que fazem parte da planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|Ocorre quando uma ou mais colunas são classificadas.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|Ocorre quando uma ou mais linhas são classificadas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|Ocorre quando a operação clicada à esquerda/mapeada ocorre na coleção de planilhas.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|Obtém a ID da planilha onde a classificação aconteceu.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|Obtém a ID da planilha onde a classificação aconteceu.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|Obtém o endereço que representa a célula que foi clicada/tocada para uma planilha específica.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|A distância, em pontos, do ponto de grade clicado/mapeado para a esquerda (ou direita para idiomas da direita para a esquerda) da célula clicada/mapeada à esquerda.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|A distância, em pontos, desde o ponto clicado/tocado com o botão esquerdo até a borda da linha de grade superior da célula clicada/tocada com o botão esquerdo.|
||[tipo](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a célula foi clicada à esquerda/tapped.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)