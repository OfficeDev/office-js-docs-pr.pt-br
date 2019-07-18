---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2199b7c115a1edd66bb7b1fef86eb3bc7bba473e
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771950"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

> [!NOTE]
> As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para usar as APIs de visualização, você deve referenciar a biblioteca **beta** no CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) e também pode ser necessário ingressar no programa Office Insider para obter uma compilação recente do Office.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Segmentação de dados](../../excel/excel-add-ins-pivottables.md#slicers-preview) | Insira e configure as segmentações de dados em tabelas e Tabelas dinâmicas. | [Segmentação de dados](/javascript/api/excel/excel.slicer) |
| [Comments](../../excel/excel-add-ins-workbooks.md#comments-preview) | Adicione, edite e exclua comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) a pasta de trabalho | Salve e feche a pasta de trabalho.  | [Workbook](/javascript/api/excel/excel.workbook) |
| [Inserir pasta de trabalho](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insira uma pasta de trabalho em outra.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Exclui o thread de comentários. |
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtém a célula em que este comentário está localizado.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtém o nome do autor do comentário.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtém o horário de criação do comentário. Retorna null se o comentário foi convertido de uma nota, pois o comentário não possui uma data de criação.|
||[id](/javascript/api/excel/excel.comment#id)|Representa o identificador de comentário. Somente leitura.|
||[replies](/javascript/api/excel/excel.comment#replies)|Representa uma coleção de objetos de resposta associados ao comentário. Somente leitura.|
||[Set (Propriedades: Excel. Comment)](/javascript/api/excel/excel.comment#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. CommentUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.comment#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Cria um novo comentário (thread de comentário) com o conteúdo fornecido na célula especificada. Um `InvalidArgument` erro será acionado se o intervalo fornecido for maior que uma célula.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Cria um novo comentário (thread de comentário) com o conteúdo fornecido na célula especificada. Um `InvalidArgument` erro será acionado se o intervalo fornecido for maior que uma célula.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtém o número de comentários na coleção.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtém um comentário da coleção com base em seu ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtém um comentário da coleção com base em sua posição.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtém o comentário da célula especificada.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtém um comentário relacionado à respectiva ID de resposta na coleção.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentCollectionData](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[CommentCollectionLoadOptions](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|Para cada ITEM na coleção: Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|Para cada ITEM na coleção: Obtém o nome do autor do comentário.|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|Para cada ITEM na coleção: Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|Para cada ITEM na coleção: Obtém a hora de criação do comentário. Retorna null se o comentário foi convertido de uma nota, pois o comentário não possui uma data de criação.|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|Para cada ITEM na coleção: representa o identificador de comentário. Somente leitura.|
|[CommentCollectionUpdateData](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[CommentData](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|Obtém o nome do autor do comentário.|
||[content](/javascript/api/excel/excel.commentdata#content)|Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|Obtém o horário de criação do comentário. Retorna null se o comentário foi convertido de uma nota, pois o comentário não possui uma data de criação.|
||[id](/javascript/api/excel/excel.commentdata#id)|Representa o identificador de comentário. Somente leitura.|
||[replies](/javascript/api/excel/excel.commentdata#replies)|Representa uma coleção de objetos de resposta associados ao comentário. Somente leitura.|
|[CommentLoadOptions](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|Obtém o email do autor do comentário.|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|Obtém o nome do autor do comentário.|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|Obtém o horário de criação do comentário. Retorna null se o comentário foi convertido de uma nota, pois o comentário não possui uma data de criação.|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|Representa o identificador de comentário. Somente leitura.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtém ou define o conteúdo da resposta do comentário. A cadeia de caracteres é de texto sem formatação.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Exclui a resposta do comentário. |
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtém a célula em que esta resposta de comentário está localizada.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtém o comentário pai desta resposta.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtém o nome do autor da resposta do comentário.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtém o horário de criação da resposta do comentário.|
||[id](/javascript/api/excel/excel.commentreply#id)|Representa o identificador de resposta do comentário. Somente leitura.|
||[Set (Propriedades: Excel. CommentReply)](/javascript/api/excel/excel.commentreply#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. CommentReplyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.commentreply#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtém o número de respostas de comentários na coleção.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Retorna uma resposta de comentário identificada pela respectiva ID. Somente leitura.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtém uma resposta de comentário com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CommentReplyCollectionData](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[CommentReplyCollectionLoadOptions](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|Para cada ITEM na coleção: Obtém o email do autor da resposta de comentário.|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|Para cada ITEM na coleção: Obtém o nome do autor da resposta de comentário.|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|Para cada ITEM na coleção: Obtém ou define o conteúdo da resposta de comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|Para cada ITEM na coleção: Obtém a hora de criação da resposta de comentário.|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|Para cada ITEM na coleção: representa o identificador de resposta de comentário. Somente leitura.|
|[CommentReplyCollectionUpdateData](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[CommentReplyData](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|Obtém o nome do autor da resposta do comentário.|
||[content](/javascript/api/excel/excel.commentreplydata#content)|Obtém ou define o conteúdo da resposta do comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|Obtém o horário de criação da resposta do comentário.|
||[id](/javascript/api/excel/excel.commentreplydata#id)|Representa o identificador de resposta do comentário. Somente leitura.|
|[CommentReplyLoadOptions](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|Obtém o email do autor da resposta do comentário.|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|Obtém o nome do autor da resposta do comentário.|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|Obtém ou define o conteúdo da resposta do comentário. A cadeia de caracteres é de texto sem formatação.|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|Obtém o horário de criação da resposta do comentário.|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|Representa o identificador de resposta do comentário. Somente leitura.|
|[CommentReplyUpdateData](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|Obtém ou define o conteúdo da resposta do comentário. A cadeia de caracteres é de texto sem formatação.|
|[CommentUpdateData](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|Obtém ou define o conteúdo do comentário. A cadeia de caracteres é de texto sem formatação.|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|Para cada ITEM na coleção: representa como o objeto é anexado às células abaixo dele.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias. A célula retornada é a interseção da linha e coluna fornecidas que contém os dados da hierarquia especificada. Esse método é o inverso de chamar getPivotItems e getDataHierarchy em uma célula específica.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|Especifica se a lista de campos pode ser mostrada na interface do usuário.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Exclui a Tabela Dinâmica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Cria uma duplicata desta Tabela Dinâmica com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtém o nome da Tabela Dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Especifica se este objeto PivotTableStyle é de somente leitura. Somente leitura.|
||[Set (Propriedades: Excel. PivotTable)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotTableStyleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Cria uma Tabela Dinâmica em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtém o número de estilos de PivotTable na coleção.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtém a Tabela Dinâmica padrão para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Obtém um PivotTableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Obtém um PivotTableStyle por nome. Se PivotTableStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault (newDefaultStyle: PivotTableStyle \| cadeia de caracteres)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Define a Tabela Dinâmica padrão para uso no escopo do objeto pai.|
|[PivotTableStyleCollectionData](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[PivotTableStyleCollectionLoadOptions](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|Para cada ITEM na coleção: Obtém o nome do PivotTable.|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|Para cada ITEM na coleção: especifica se este objeto PivotTable é somente leitura. Somente leitura.|
|[PivotTableStyleCollectionUpdateData](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[PivotTableStyleData](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|Obtém o nome da Tabela Dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|Especifica se este objeto PivotTableStyle é de somente leitura. Somente leitura.|
|[PivotTableStyleLoadOptions](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|Obtém o nome da Tabela Dinâmica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|Especifica se este objeto PivotTableStyle é de somente leitura. Somente leitura.|
|[PivotTableStyleUpdateData](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|Obtém o nome da Tabela Dinâmica.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtém o objeto range que contém a célula âncora para uma célula que recebe o despejo. Somente leitura.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Falha se aplicado a um intervalo com mais de uma célula. Somente leitura.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtém objeto range que contém o intervalo de despejo quando chamado em uma célula âncora. Somente leitura.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[height](/javascript/api/excel/excel.range#height)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo. Somente leitura.|
||[left](/javascript/api/excel/excel.range#left)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha para a borda esquerda do intervalo. Somente leitura.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
||[top](/javascript/api/excel/excel.range#top)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo. Somente leitura.|
||[width](/javascript/api/excel/excel.range#width)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo. Somente leitura.|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|Para cada ITEM na coleção: representa se todas as células têm uma borda de despejo.|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|Para cada ITEM na coleção: retorna a distância em pontos, para 100% de zoom, da borda superior do intervalo para a borda inferior do intervalo. Somente leitura.|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|Para cada ITEM na coleção: retorna a distância em pontos, para 100% de zoom, da borda esquerda da planilha para a borda esquerda do intervalo. Somente leitura.|
||[savedAsArray](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|Para cada ITEM na coleção: representa se todas as células seriam salvas como uma fórmula de matriz.|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|Para cada ITEM na coleção: retorna a distância em pontos, para 100% de zoom, da borda superior da planilha até a borda superior do intervalo. Somente leitura.|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|Para cada ITEM na coleção: retorna a distância em pontos, para 100% de zoom, da borda esquerda do intervalo à borda direita do intervalo. Somente leitura.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[height](/javascript/api/excel/excel.rangedata#height)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo. Somente leitura.|
||[left](/javascript/api/excel/excel.rangedata#left)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha para a borda esquerda do intervalo. Somente leitura.|
||[savedAsArray](/javascript/api/excel/excel.rangedata#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
||[top](/javascript/api/excel/excel.rangedata#top)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo. Somente leitura.|
||[width](/javascript/api/excel/excel.rangedata#width)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo. Somente leitura.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|Representa se todas as células têm uma borda de despejo.|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|Retorna a distância em pontos, para zoom de 100%, da borda superior do intervalo até a borda inferior do intervalo. Somente leitura.|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda da planilha para a borda esquerda do intervalo. Somente leitura.|
||[savedAsArray](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|Representa se todas as células seriam salvas como uma fórmula de matriz.|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|Retorna a distância em pontos, para zoom de 100%, da borda superior da planilha até a borda superior do intervalo. Somente leitura.|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|Retorna a distância em pontos, para zoom de 100%, da borda esquerda do intervalo até a borda direita do intervalo. Somente leitura.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copia e cola um objeto Forma.|
||[placement](/javascript/api/excel/excel.shape#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha. Retorna um objeto Shape que representa a nova imagem.|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|Para cada ITEM na coleção: representa como o objeto é anexado às células abaixo dele.|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|Representa como o objeto é anexado às células abaixo dela.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Representa a legenda da segmentação de dados.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Limpa todos os filtros aplicados à segmentação de dados no momento.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Exclui a segmentação de dados.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Retorna uma matriz de chaves de itens selecionados. Somente leitura.|
||[height](/javascript/api/excel/excel.slicer#height)|Representa a altura, em pontos, da segmentação de dados.|
||[left](/javascript/api/excel/excel.slicer#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicer#name)|Representa o nome da segmentação de dados.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[id](/javascript/api/excel/excel.slicer#id)|Representa a id exclusiva da segmentação de dados. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|Verdadeiro se todos os filtros atualmente aplicados na segmentação de dados estiverem desmarcados.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Representa a coleção de SlicerItems que faz parte da segmentação de dados. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Representa a planilha que contém a segmentação de dados. Somente leitura.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Seleciona os itens da segmentação de dados com base em suas chaves. A seleção anterior será limpa.|
||[Set (Propriedades: Excel. slicer)](/javascript/api/excel/excel.slicer#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. SlicerUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.slicer#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicer#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicer#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicer#width)|Representa a largura, em pontos, da segmentação de dados.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adiciona uma nova segmentação de dados à pasta de trabalho.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Retorna o número de segmentações de dados na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtém um objeto de segmentação de dados usando seu respectivo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtém uma segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtém uma segmentação de dados usando seu nome ou id. Se a ela não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerCollectionData](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[SlicerCollectionLoadOptions](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|Para cada ITEM na coleção: representa a legenda da segmentação de itens.|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|Para cada ITEM na coleção: representa a altura, em pontos, da segmentação de tópicos.|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|Para cada ITEM na coleção: representa a ID exclusiva de slicer. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|Para cada ITEM na coleção: true se todos os filtros aplicados no momento na segmentação de trabalho são limpos.|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|Para cada ITEM na coleção: representa a distância, em pontos, do lado esquerdo da segmentação de tópicos à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|Para cada ITEM na coleção: representa o nome da segmentação de itens.|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|Para cada ITEM na coleção: representa o nome da segmentação de itens usado na fórmula.|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|Para cada ITEM na coleção: representa a ordem de classificação dos itens na segmentação de,. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|Para cada ITEM da coleção: valor constante que representa o estilo de segmentação de itens. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|Para cada ITEM na coleção: representa a distância, em pontos, da borda superior da segmentação de itens à parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|Para cada ITEM na coleção: representa a largura, em pontos, da segmentação de tópicos.|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|Para cada ITEM na coleção: representa a planilha que contém a segmentação de conteúdo.|
|[SlicerCollectionUpdateData](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[SlicerData](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|Representa a legenda da segmentação de dados.|
||[height](/javascript/api/excel/excel.slicerdata#height)|Representa a altura, em pontos, da segmentação de dados.|
||[id](/javascript/api/excel/excel.slicerdata#id)|Representa a id exclusiva da segmentação de dados. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|Verdadeiro se todos os filtros atualmente aplicados na segmentação de dados estiverem desmarcados.|
||[left](/javascript/api/excel/excel.slicerdata#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicerdata#name)|Representa o nome da segmentação de dados.|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|Representa a coleção de SlicerItems que faz parte da segmentação de dados. Somente leitura.|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicerdata#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicerdata#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicerdata#width)|Representa a largura, em pontos, da segmentação de dados.|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|Representa a planilha que contém a segmentação de dados. Somente leitura.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True se o item da segmentação de dados estiver selecionado.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True se o item de segmentação de dados tiver dados.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Representa o título exibido na interface do usuário.|
||[Set (Propriedades: Excel. SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. SlicerItemUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Retorna o número de itens da segmentação de dados na segmentação de dados.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtém um objeto de item da segmentação de dados usando sua chave ou nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtém um item da segmentação de dados com base em sua posição na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtém um item da segmentação de dados usando sua chave ou nome. Se o item da segmentação de dados não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlicerItemCollectionData](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[SlicerItemCollectionLoadOptions](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|Para cada ITEM na coleção: true se o item da segmentação de dados tiver dados.|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|Para cada ITEM na coleção: true se o item da segmentação de itens for selecionado.|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|Para cada ITEM na coleção: representa o valor exclusivo que representa o item da segmentação de itens.|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|Para cada ITEM na coleção: representa o título exibido na interface do usuário.|
|[SlicerItemCollectionUpdateData](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[SlicerItemData](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|True se o item de segmentação de dados tiver dados.|
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|True se o item da segmentação de dados estiver selecionado.|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|Representa o título exibido na interface do usuário.|
|[SlicerItemLoadOptions](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|True se o item de segmentação de dados tiver dados.|
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|True se o item da segmentação de dados estiver selecionado.|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|Representa o valor exclusivo que representa o item da segmentação de dados.|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|Representa o título exibido na interface do usuário.|
|[SlicerItemUpdateData](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|True se o item da segmentação de dados estiver selecionado.|
|[SlicerLoadOptions](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|Representa a legenda da segmentação de dados.|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|Representa a altura, em pontos, da segmentação de dados.|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|Representa a id exclusiva da segmentação de dados. Somente leitura.|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|Verdadeiro se todos os filtros atualmente aplicados na segmentação de dados estiverem desmarcados.|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|Representa o nome da segmentação de dados.|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|Representa a largura, em pontos, da segmentação de dados.|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|Representa a planilha que contém a segmentação de dados.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Exclui o SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Cria uma duplicata deste SlicerStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtém o nome o SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Especifica se este objeto SlicerStyle é de somente leitura. Somente leitura.|
||[Set (Propriedades: Excel. SlicerStyle)](/javascript/api/excel/excel.slicerstyle#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. SlicerStyleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Cria um SlicerStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtém o número de segmentação de estilos na coleção.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtém o padrão SlicerStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtém uma SlicerStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtém uma SlicerStyle por nome. Se o SlicerStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Define o padrão SlicerStyle para uso no escopo do objeto pai.|
|[SlicerStyleCollectionData](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[SlicerStyleCollectionLoadOptions](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|Para cada ITEM na coleção: Obtém o nome do SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|Para cada ITEM na coleção: especifica se este objeto SlicerStyle é somente leitura. Somente leitura.|
|[SlicerStyleCollectionUpdateData](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[SlicerStyleData](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|Obtém o nome o SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|Especifica se este objeto SlicerStyle é de somente leitura. Somente leitura.|
|[SlicerStyleLoadOptions](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|Obtém o nome o SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|Especifica se este objeto SlicerStyle é de somente leitura. Somente leitura.|
|[SlicerStyleUpdateData](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|Obtém o nome o SlicerStyle.|
|[SlicerUpdateData](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|Representa a legenda da segmentação de dados.|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|Representa a altura, em pontos, da segmentação de dados.|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|Representa a distância, em pontos, da lateral esquerda da segmentação de dados à esquerda da planilha.|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|Representa o nome da segmentação de dados.|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|Representa a ordem de classificação dos itens na segmentação de dados. Valores possíveis são: DataSourceOrder, Ordem crescente, Ordem decrescente.|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|Valor da constante que representa o estilo da Segmentação de dados. Os valores possíveis são: "SlicerStyleLight1" por meio de "SlicerStyleLight6", "TableStyleOther1" até "TableStyleOther2", "SlicerStyleDark1" até "SlicerStyleDark6". Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|Representa a distância, em pontos, da borda superior da segmentação de dados na parte superior da planilha.|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|Representa a largura, em pontos, da segmentação de dados.|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|Representa a planilha que contém a segmentação de dados.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Representa a id da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Representa a id da planilha que contém a tabela.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Cria uma duplicata deste TableStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Especifica se este objeto TableStyle é de somente leitura. Somente leitura.|
||[Set (Propriedades: Excel. TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TableStyleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Cria um TableStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtém o número de estilos de tabelas na coleção.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtém o padrão TableStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Obtém um TableStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Obtém um TableStyle por nome. Se o TableStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Define a TableStyle padrão para uso no escopo do objeto pai..|
|[TableStyleCollectionData](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[TableStyleCollectionLoadOptions](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|Para cada ITEM na coleção: Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|Para cada ITEM na coleção: especifica se este objeto TableStyle é somente leitura. Somente leitura.|
|[TableStyleCollectionUpdateData](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[TableStyleData](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|Especifica se este objeto TableStyle é de somente leitura. Somente leitura.|
|[TableStyleLoadOptions](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|Obtém o nome do TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|Especifica se este objeto TableStyle é de somente leitura. Somente leitura.|
|[TableStyleUpdateData](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|Obtém o nome do TableStyle.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Exclui o TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Cria uma duplicata deste TimelineStyle com cópias de todos os elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtém o nome do TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Especifica se este objeto TimelineStyle é de somente leitura. Somente leitura.|
||[Set (Propriedades: Excel. timelinestyle)](/javascript/api/excel/excel.timelinestyle#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TimelineStyleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Cria um TimelineStyle em branco com o nome especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtém o número de estilos de linha do tempo na coleção.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtém o padrão TimelineStyle para o escopo do objeto pai.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtém uma TimelineStyle por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtém uma TimelineStyle por nome. Se o TimelineStyle não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Define o padrão TimelineStyle para uso no escopo do objeto pai.|
|[TimelineStyleCollectionData](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[TimelineStyleCollectionLoadOptions](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|Para cada ITEM na coleção: Obtém o nome do timelinestyle.|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|Para cada ITEM na coleção: especifica se esse objeto timelinestyle é somente leitura. Somente leitura.|
|[TimelineStyleCollectionUpdateData](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[TimelineStyleData](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|Obtém o nome do TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|Especifica se este objeto TimelineStyle é de somente leitura. Somente leitura.|
|[TimelineStyleLoadOptions](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|Obtém o nome do TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|Especifica se este objeto TimelineStyle é de somente leitura. Somente leitura.|
|[TimelineStyleUpdateData](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|Obtém o nome do TimelineStyle.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver um slicer ativo, uma `ItemNotFound` exceção será lançada.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtém a segmentação de dados ativa no momento na pasta de trabalho. Se não houver segmentação de dados ativa, um objeto nulo será retornado.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Representa uma coleção de comentários associados à pasta de trabalho. Somente leitura.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Representa uma coleção de SlicerStyles associados à pasta de trabalho. Somente leitura.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Representa uma coleção de segmentações de dados associados à pasta de trabalho. Somente leitura.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Representa uma coleção de TableStyles associadas à pasta de trabalho. Somente leitura.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Representa uma coleção de TimelineStyles associados à pasta de trabalho. Somente leitura.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|Representa uma coleção de comentários associados à pasta de trabalho. Somente leitura.|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|Representa uma coleção de SlicerStyles associados à pasta de trabalho. Somente leitura.|
||[slicers](/javascript/api/excel/excel.workbookdata#slicers)|Representa uma coleção de segmentações de dados associados à pasta de trabalho. Somente leitura.|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|Representa uma coleção de TableStyles associadas à pasta de trabalho. Somente leitura.|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|Representa uma coleção de TimelineStyles associados à pasta de trabalho. Somente leitura.|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Retorna um conjunto de todos os objetos Comments na planilha. Somente leitura.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Ocorre durante a classificação de colunas.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Ocorre quando o estado da linha oculto é alterado em uma planilha específica.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Ocorre durante a classificação de linhas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Ocorre quando a operação clicada/pressionada à esquerda ocorre na planilha.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Retorna uma coleção de segmentações de dados que fazem parte da planilha. Somente leitura.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Ocorre durante a classificação de colunas.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Ocorre quando qualquer planilha na pasta de trabalho tem o estado oculto de linha alterado.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Ocorre durante a classificação de linhas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Ocorre quando a operação com o botão esquerdo/tocado acontece na coleção de planilhas.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtém o id da planilha onde a classificação aconteceu.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|Retorna um conjunto de todos os objetos Comments na planilha. Somente leitura.|
||[slicers](/javascript/api/excel/excel.worksheetdata#slicers)|Retorna uma coleção de segmentações de dados que fazem parte da planilha. Somente leitura.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Representa a id da planilha na qual o filtro é aplicado.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtém o tipo de mudança que representa como o evento Changed é acionado. Consulte Excel. RowHiddenChangeType para obter detalhes.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtém o endereço do intervalo que representa as áreas classificadas de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtém o id da planilha onde a classificação aconteceu.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtém o endereço que representa a célula que foi clicada/tocada para uma planilha específica.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|A distância, em pontos, do ponto clicado com o botão esquerdo/tocado até a borda esquerda (direita para RTL) da linha de grade da célula clicada com o botão esquerdo/tocada.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|A distância, em pontos, desde o ponto clicado/tocado com o botão esquerdo até a borda da linha de grade superior da célula clicada/tocada com o botão esquerdo.|
||[tipo](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtém o id da planilha na qual a célula foi clicada com o botão esquerdo/tocada.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
