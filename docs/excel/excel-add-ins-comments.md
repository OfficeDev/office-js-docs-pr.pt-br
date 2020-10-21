---
title: Trabalhar com comentários usando a API JavaScript do Excel
description: Informações sobre como usar as APIs para adicionar, remover e editar comentários e encadeamentos de comentários.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 00f7dd22fb2148902152197521098482071e5284
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626418"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="84b29-103">Trabalhar com comentários usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="84b29-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="84b29-104">Este artigo descreve como adicionar, ler, modificar e remover comentários em uma pasta de trabalho com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="84b29-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="84b29-105">Você pode saber mais sobre o recurso comentário do artigo [inserir comentários e anotações no Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="84b29-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="84b29-106">Na API JavaScript do Excel, um comentário inclui o único comentário inicial e a discussão encadeada conectada.</span><span class="sxs-lookup"><span data-stu-id="84b29-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="84b29-107">Ele está vinculado a uma célula individual.</span><span class="sxs-lookup"><span data-stu-id="84b29-107">It is tied to an individual cell.</span></span> <span data-ttu-id="84b29-108">Qualquer pessoa que exiba a pasta de trabalho com permissões suficientes pode responder a um comentário.</span><span class="sxs-lookup"><span data-stu-id="84b29-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="84b29-109">Um objeto [comment](/javascript/api/excel/excel.comment) armazena as respostas como objetos [CommentReply](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="84b29-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="84b29-110">Você deve considerar um comentário para ser um thread e que um thread deve ter uma entrada especial como o ponto de partida.</span><span class="sxs-lookup"><span data-stu-id="84b29-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Um comentário do Excel, rotulado "comentário" com duas respostas, rotuladas "comentário. respostas [0]" e "comentário. respostas [1].](../images/excel-comments.png)

<span data-ttu-id="84b29-112">Os comentários em uma pasta de trabalho são rastreados pela `Workbook.comments` propriedade.</span><span class="sxs-lookup"><span data-stu-id="84b29-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="84b29-113">Isso inclui comentários criados por usuários e comentários criados por seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="84b29-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="84b29-114">A propriedade `Workbook.comments` é um objeto [CommentCollection](/javascript/api/excel/excel.commentcollection) que contém um conjunto de objetos [Comentário](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="84b29-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="84b29-115">Os comentários também podem ser acessados no nível da [planilha](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="84b29-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="84b29-116">Os exemplos neste artigo trabalham com comentários no nível da pasta de trabalho, mas eles podem ser facilmente modificados para usar a `Worksheet.comments` propriedade.</span><span class="sxs-lookup"><span data-stu-id="84b29-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="84b29-117">Adicionar comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-117">Add comments</span></span>

<span data-ttu-id="84b29-118">Use o `CommentCollection.add` método para adicionar comentários a uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="84b29-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="84b29-119">Este método utiliza até três parâmetros:</span><span class="sxs-lookup"><span data-stu-id="84b29-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="84b29-120">`cellAddress`: A célula onde o comentário é adicionado.</span><span class="sxs-lookup"><span data-stu-id="84b29-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="84b29-121">Pode ser um objeto String ou [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="84b29-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="84b29-122">O intervalo deve ser uma única célula.</span><span class="sxs-lookup"><span data-stu-id="84b29-122">The range must be a single cell.</span></span>
- <span data-ttu-id="84b29-123">`content`: O conteúdo do comentário.</span><span class="sxs-lookup"><span data-stu-id="84b29-123">`content`: The comment's content.</span></span> <span data-ttu-id="84b29-124">Use uma cadeia de caracteres para comentários de texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="84b29-124">Use a string for plain text comments.</span></span> <span data-ttu-id="84b29-125">Use um objeto [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) para comentários com [menção](#mentions).</span><span class="sxs-lookup"><span data-stu-id="84b29-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="84b29-126">`contentType`: Um enum [ContentType](/javascript/api/excel/excel.contenttype) especificando o tipo de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="84b29-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="84b29-127">O valor padrão é `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="84b29-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="84b29-128">O exemplo a seguir adiciona um comentário à célula **A2**.</span><span class="sxs-lookup"><span data-stu-id="84b29-128">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> <span data-ttu-id="84b29-129">Os comentários adicionados por um suplemento são atribuídos ao usuário atual desse suplemento.</span><span class="sxs-lookup"><span data-stu-id="84b29-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="84b29-130">Adicionar respostas de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-130">Add comment replies</span></span>

<span data-ttu-id="84b29-131">Um `Comment` objeto é um thread de comentário que contém zero ou mais respostas.</span><span class="sxs-lookup"><span data-stu-id="84b29-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="84b29-132">os objetos `Comment` têm uma propriedade `replies`, que é [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) que contém objetos [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="84b29-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="84b29-133">Para adicionar uma resposta a um comentário, use o método `CommentReplyCollection.add`, passando o texto da resposta.</span><span class="sxs-lookup"><span data-stu-id="84b29-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="84b29-134">As respostas são exibidas na ordem em que são adicionadas.</span><span class="sxs-lookup"><span data-stu-id="84b29-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="84b29-135">Eles também são atribuídos ao usuário atual do suplemento.</span><span class="sxs-lookup"><span data-stu-id="84b29-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="84b29-136">O exemplo a seguir adiciona uma resposta ao primeiro comentário da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="84b29-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="84b29-137">Editar comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-137">Edit comments</span></span>

<span data-ttu-id="84b29-138">Para editar um comentário ou uma resposta de comentário, defina uma propriedade`Comment.content` e uma propriedade`CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="84b29-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="84b29-139">Editar respostas de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-139">Edit comment replies</span></span>

<span data-ttu-id="84b29-140">Para editar uma resposta de comentário, defina sua `CommentReply.content` propriedade.</span><span class="sxs-lookup"><span data-stu-id="84b29-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="84b29-141">Excluir comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-141">Delete comments</span></span>

<span data-ttu-id="84b29-142">Para excluir um comentário, use o `Comment.delete` método.</span><span class="sxs-lookup"><span data-stu-id="84b29-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="84b29-143">A exclusão de um comentário também exclui as respostas associadas a esse comentário.</span><span class="sxs-lookup"><span data-stu-id="84b29-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="84b29-144">Excluir respostas de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-144">Delete comment replies</span></span>

<span data-ttu-id="84b29-145">Para excluir uma resposta de comentário, use o `CommentReply.delete` método.</span><span class="sxs-lookup"><span data-stu-id="84b29-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="84b29-146">Resolver threads de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-146">Resolve comment threads</span></span>

<span data-ttu-id="84b29-147">Um thread de comentário tem um valor booliano configurável, `resolved` para indicar se ele foi resolvido.</span><span class="sxs-lookup"><span data-stu-id="84b29-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="84b29-148">Um valor de `true` significa que o thread de comentários é resolvido.</span><span class="sxs-lookup"><span data-stu-id="84b29-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="84b29-149">Um valor de `false` significa que o thread de comentários é novo ou reaberto.</span><span class="sxs-lookup"><span data-stu-id="84b29-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="84b29-150">Respostas de comentário têm uma `resolved` propriedade ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="84b29-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="84b29-151">Seu valor é sempre igual ao do restante do thread.</span><span class="sxs-lookup"><span data-stu-id="84b29-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="84b29-152">Metadados de comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-152">Comment metadata</span></span>

<span data-ttu-id="84b29-153">Cada comentário contém metadados sobre a criação, como o autor e a data de criação.</span><span class="sxs-lookup"><span data-stu-id="84b29-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="84b29-154">Os comentários criados por seu suplemento são considerados criados pelo usuário atual.</span><span class="sxs-lookup"><span data-stu-id="84b29-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="84b29-155">O exemplo a seguir mostra como exibir o email do autor, o nome do autor e a data de criação de um comentário em **A2**.</span><span class="sxs-lookup"><span data-stu-id="84b29-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a><span data-ttu-id="84b29-156">Metadados de resposta de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-156">Comment reply metadata</span></span>

<span data-ttu-id="84b29-157">Respostas de comentários armazenam os mesmos tipos de metadados que o comentário inicial.</span><span class="sxs-lookup"><span data-stu-id="84b29-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="84b29-158">O exemplo a seguir mostra como exibir o email do autor, o nome do autor e a data de criação da resposta de comentário mais recente em **a2**.</span><span class="sxs-lookup"><span data-stu-id="84b29-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions"></a><span data-ttu-id="84b29-159">Menções</span><span class="sxs-lookup"><span data-stu-id="84b29-159">Mentions</span></span>

<span data-ttu-id="84b29-160">As [mencionas](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) são usadas para marcar colegas em um comentário.</span><span class="sxs-lookup"><span data-stu-id="84b29-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="84b29-161">Isso envia notificações com o conteúdo do comentário.</span><span class="sxs-lookup"><span data-stu-id="84b29-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="84b29-162">O suplemento pode criar essas menção em seu nome.</span><span class="sxs-lookup"><span data-stu-id="84b29-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="84b29-163">Comentários com menção precisam ser criados com objetos [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="84b29-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="84b29-164">Call `CommentCollection.add` com um `CommentRichContent` contendo um ou mais mencionas e especifique `ContentType.mention` como o `contentType` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="84b29-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="84b29-165">A `content` cadeia de caracteres também precisa ser formatada para inserir o menção no texto.</span><span class="sxs-lookup"><span data-stu-id="84b29-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="84b29-166">O formato de um menção é: `<at id="{replyIndex}">{mentionName}</at>` .</span><span class="sxs-lookup"><span data-stu-id="84b29-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="84b29-167">Atualmente, apenas o nome exato de menção pode ser usado como o texto do link de menção.</span><span class="sxs-lookup"><span data-stu-id="84b29-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="84b29-168">O suporte para versões reduzidas de um nome será adicionado posteriormente.</span><span class="sxs-lookup"><span data-stu-id="84b29-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="84b29-169">O exemplo a seguir mostra um comentário com uma única menção.</span><span class="sxs-lookup"><span data-stu-id="84b29-169">The following example shows a comment with a single mention.</span></span>

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="comment-events"></a><span data-ttu-id="84b29-170">Eventos de comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-170">Comment events</span></span>

<span data-ttu-id="84b29-171">O suplemento pode ouvir adições, alterações e exclusões de comentários.</span><span class="sxs-lookup"><span data-stu-id="84b29-171">Your add-in can listen for comment additions, changes, and deletions.</span></span> <span data-ttu-id="84b29-172">[Eventos de comentários](/javascript/api/excel/excel.commentcollection#event-details) ocorrem no `CommentCollection` objeto.</span><span class="sxs-lookup"><span data-stu-id="84b29-172">[Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.</span></span> <span data-ttu-id="84b29-173">Para ouvir eventos de comentários, registre o `onAdded` , `onChanged` ou o `onDeleted` manipulador de eventos comment.</span><span class="sxs-lookup"><span data-stu-id="84b29-173">To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler.</span></span> <span data-ttu-id="84b29-174">Quando um evento Comment é detectado, use este manipulador de eventos para recuperar dados sobre o Comentário adicionado, alterado ou excluído.</span><span class="sxs-lookup"><span data-stu-id="84b29-174">When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment.</span></span> <span data-ttu-id="84b29-175">O `onChanged` evento também trata de adições de comentários, alterações e exclusões.</span><span class="sxs-lookup"><span data-stu-id="84b29-175">The `onChanged` event also handles comment reply additions, changes, and deletions.</span></span> 

<span data-ttu-id="84b29-176">Cada evento de comentário é acionado apenas uma vez quando várias adições, alterações ou exclusões são realizadas ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="84b29-176">Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time.</span></span> <span data-ttu-id="84b29-177">Todos os objetos [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)e [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contêm matrizes de IDs de comentários para mapear as ações de evento de volta para as coleções de comentários.</span><span class="sxs-lookup"><span data-stu-id="84b29-177">All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.</span></span>

<span data-ttu-id="84b29-178">Confira o artigo [trabalhar com eventos usando o Excel JavaScript API](excel-add-ins-events.md) para obter mais informações sobre como registrar manipuladores de eventos, manipular eventos e remover manipuladores de eventos.</span><span class="sxs-lookup"><span data-stu-id="84b29-178">See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.</span></span> 

### <a name="comment-addition-events"></a><span data-ttu-id="84b29-179">Eventos de adição de comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-179">Comment addition events</span></span> 
<span data-ttu-id="84b29-180">O `onAdded` evento é disparado quando um ou mais comentários novos são adicionados à coleção comment.</span><span class="sxs-lookup"><span data-stu-id="84b29-180">The `onAdded` event is triggered when one or more new comments are added to the comment collection.</span></span> <span data-ttu-id="84b29-181">Esse evento *não* é disparado quando as respostas são adicionadas a um thread de comentários (consulte comentários sobre eventos de [alteração](#comment-change-events) para saber mais sobre eventos de resposta de comentários).</span><span class="sxs-lookup"><span data-stu-id="84b29-181">This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).</span></span>

<span data-ttu-id="84b29-182">O exemplo a seguir mostra como registrar o `onAdded` manipulador de eventos e, em seguida, usar o `CommentAddedEventArgs` objeto para recuperar a `commentDetails` matriz do Comentário adicionado.</span><span class="sxs-lookup"><span data-stu-id="84b29-182">The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.</span></span>

> [!NOTE]
> <span data-ttu-id="84b29-183">Este exemplo só funciona quando um único comentário é adicionado.</span><span class="sxs-lookup"><span data-stu-id="84b29-183">This sample only works when a single comment is added.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a><span data-ttu-id="84b29-184">Eventos de alteração de comentário</span><span class="sxs-lookup"><span data-stu-id="84b29-184">Comment change events</span></span> 
<span data-ttu-id="84b29-185">O `onChanged` evento Comment é disparado nos cenários a seguir.</span><span class="sxs-lookup"><span data-stu-id="84b29-185">The `onChanged` comment event is triggered in the following scenarios.</span></span>

- <span data-ttu-id="84b29-186">O conteúdo de um comentário é atualizado.</span><span class="sxs-lookup"><span data-stu-id="84b29-186">A comment's content is updated.</span></span>
- <span data-ttu-id="84b29-187">Um thread de comentários é resolvido.</span><span class="sxs-lookup"><span data-stu-id="84b29-187">A comment thread is resolved.</span></span>
- <span data-ttu-id="84b29-188">Um thread de comentários é reaberto.</span><span class="sxs-lookup"><span data-stu-id="84b29-188">A comment thread is reopened.</span></span>
- <span data-ttu-id="84b29-189">Uma resposta é adicionada a um thread de comentários.</span><span class="sxs-lookup"><span data-stu-id="84b29-189">A reply is added to a comment thread.</span></span>
- <span data-ttu-id="84b29-190">Uma resposta é atualizada em um thread de comentários.</span><span class="sxs-lookup"><span data-stu-id="84b29-190">A reply is updated in a comment thread.</span></span>
- <span data-ttu-id="84b29-191">Uma resposta é excluída em um thread de comentários.</span><span class="sxs-lookup"><span data-stu-id="84b29-191">A reply is deleted in a comment thread.</span></span>

<span data-ttu-id="84b29-192">O exemplo a seguir mostra como registrar o `onChanged` manipulador de eventos e, em seguida, usar o `CommentChangedEventArgs` objeto para recuperar a `commentDetails` matriz do comentário alterado.</span><span class="sxs-lookup"><span data-stu-id="84b29-192">The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.</span></span>

> [!NOTE]
> <span data-ttu-id="84b29-193">Este exemplo só funciona quando um único comentário é alterado.</span><span class="sxs-lookup"><span data-stu-id="84b29-193">This sample only works when a single comment is changed.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a><span data-ttu-id="84b29-194">Eventos de exclusão de comentários</span><span class="sxs-lookup"><span data-stu-id="84b29-194">Comment deletion events</span></span>
<span data-ttu-id="84b29-195">O `onDeleted` evento é disparado quando um comentário é excluído da coleção comment.</span><span class="sxs-lookup"><span data-stu-id="84b29-195">The `onDeleted` event is triggered when a comment is deleted from the comment collection.</span></span> <span data-ttu-id="84b29-196">Após a exclusão de um comentário, seus metadados não estão mais disponíveis.</span><span class="sxs-lookup"><span data-stu-id="84b29-196">Once a comment has been deleted, its metadata is no longer available.</span></span> <span data-ttu-id="84b29-197">O objeto [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fornece IDs de comentários, caso o suplemento esteja gerenciando Comentários individuais.</span><span class="sxs-lookup"><span data-stu-id="84b29-197">The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.</span></span>

<span data-ttu-id="84b29-198">O exemplo a seguir mostra como registrar o `onDeleted` manipulador de eventos e, em seguida, usar o `CommentDeletedEventArgs` objeto para recuperar a `commentDetails` matriz do comentário excluído.</span><span class="sxs-lookup"><span data-stu-id="84b29-198">The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.</span></span>

> [!NOTE]
> <span data-ttu-id="84b29-199">Este exemplo só funciona quando um único comentário é excluído.</span><span class="sxs-lookup"><span data-stu-id="84b29-199">This sample only works when a single comment is deleted.</span></span> 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a><span data-ttu-id="84b29-200">Confira também</span><span class="sxs-lookup"><span data-stu-id="84b29-200">See also</span></span>

- [<span data-ttu-id="84b29-201">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="84b29-201">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="84b29-202">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="84b29-202">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="84b29-203">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="84b29-203">Work with Events using the Excel JavaScript API</span></span>](excel-add-ins-events.md)
- [<span data-ttu-id="84b29-204">Inserir comentários e anotações no Excel</span><span class="sxs-lookup"><span data-stu-id="84b29-204">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
