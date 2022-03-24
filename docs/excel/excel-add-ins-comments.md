---
title: Trabalhar com comentários usando a EXCEL JavaScript
description: Informações sobre como usar as APIs para adicionar, remover e editar comentários e threads de comentários.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2b8994ed9941eec1321b4fbc5029388cd82a65fc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744814"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Trabalhar com comentários usando a EXCEL JavaScript

Este artigo descreve como adicionar, ler, modificar e remover comentários em uma manual de trabalho com a API JavaScript Excel JavaScript. Você pode saber mais sobre o recurso de comentário no [artigo Inserir comentários e anotações Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) artigo.

Na api Excel JavaScript, um comentário inclui o comentário inicial único e a discussão encadeada conectada. Ela está vinculada a uma célula individual. Qualquer pessoa que exibir a workbook com permissões suficientes pode responder a um comentário. Um [objeto Comment](/javascript/api/excel/excel.comment) armazena essas respostas como [objetos CommentReply](/javascript/api/excel/excel.commentreply) . Você deve considerar um comentário como um thread e que um thread deve ter uma entrada especial como ponto de partida.

![Um Excel comentário, rotulado como "Comentário" com duas respostas, rotulado "Comment.replies[0]" e "Comment.replies[1].](../images/excel-comments.png)

Os comentários em uma workbook são acompanhados pela `Workbook.comments` propriedade. Isso inclui comentários criados por usuários e comentários criados por seu suplemento. A propriedade `Workbook.comments` é um objeto [CommentCollection](/javascript/api/excel/excel.commentcollection) que contém um conjunto de objetos [Comentário](/javascript/api/excel/excel.comment). Os comentários também são acessíveis no [nível da Planilha](/javascript/api/excel/excel.worksheet) . Os exemplos neste artigo funcionam com comentários no nível da workbook, mas podem ser facilmente modificados para usar a `Worksheet.comments` propriedade.

## <a name="add-comments"></a>Adicionar comentários

Use o `CommentCollection.add` método para adicionar comentários a uma workbook. Este método tem até três parâmetros:

- `cellAddress`: a célula onde o comentário é adicionado. Pode ser uma cadeia de caracteres ou [um objeto Range](/javascript/api/excel/excel.range) . O intervalo deve ser uma única célula.
- `content`: O conteúdo do comentário. Use uma cadeia de caracteres para comentários de texto sem texto. Use um [objeto CommentRichContent](/javascript/api/excel/excel.commentrichcontent) para comentários com [menções](#mentions).
- `contentType`: Um [número de ContentType](/javascript/api/excel/excel.contenttype) que especifica o tipo de conteúdo. O valor padrão é `ContentType.plain`.

O exemplo a seguir adiciona um comentário à célula **A2**.

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> Comentários adicionados por um complemento são atribuídos ao usuário atual desse complemento.

### <a name="add-comment-replies"></a>Adicionar respostas de comentários

Um `Comment` objeto é um thread de comentário que contém zero ou mais respostas. os objetos `Comment` têm uma propriedade `replies`, que é [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) que contém objetos [CommentReply](/javascript/api/excel/excel.commentreply). Para adicionar uma resposta a um comentário, use o método `CommentReplyCollection.add`, passando o texto da resposta. As respostas são exibidas na ordem em que são adicionadas. Eles também são atribuídos ao usuário atual do complemento.

O exemplo a seguir adiciona uma resposta ao primeiro comentário da pasta de trabalho.

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## <a name="edit-comments"></a>Editar comentários

Para editar um comentário ou uma resposta de comentário, defina uma propriedade`Comment.content` e uma propriedade`CommentReply.content`.

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### <a name="edit-comment-replies"></a>Editar respostas de comentários

Para editar uma resposta de comentário, de definir sua `CommentReply.content` propriedade.

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## <a name="delete-comments"></a>Excluir comentários

Para excluir um comentário, use o `Comment.delete` método. Excluir um comentário também exclui as respostas associadas a esse comentário.

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### <a name="delete-comment-replies"></a>Excluir respostas de comentários

Para excluir uma resposta de comentário, use o `CommentReply.delete` método.

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="resolve-comment-threads"></a>Resolver threads de comentário

Um thread de comentário tem um valor boolano configurável, `resolved`, para indicar se ele foi resolvido. Um valor de significa `true` que o thread de comentário é resolvido. Um valor significa que `false` o thread de comentário é novo ou reaberto.

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

As respostas de comentários têm uma propriedade readonly `resolved` . Seu valor é sempre igual ao do restante do thread.

## <a name="comment-metadata"></a>Metadados de comentário

Cada comentário contém metadados sobre a criação, como o autor e a data de criação. Os comentários criados por seu suplemento são considerados criados pelo usuário atual.

O exemplo a seguir mostra como exibir o email do autor, o nome do autor e a data de criação de um comentário em **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### <a name="comment-reply-metadata"></a>Metadados de resposta de comentário

Respostas de comentário armazenam os mesmos tipos de metadados que o comentário inicial.

O exemplo a seguir mostra como exibir o email do autor, o nome do autor e a data de criação da resposta de comentário mais recente em **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## <a name="mentions"></a>Menções

[As menções](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) são usadas para marcar colegas em um comentário. Isso envia notificações com o conteúdo do seu comentário. Seu complemento pode criar essas menções em seu nome.

Comentários com menções precisam ser criados com [objetos CommentRichContent](/javascript/api/excel/excel.commentrichcontent) . Chame `CommentCollection.add` com uma `CommentRichContent` ou mais menções e especifique `ContentType.mention` como o `contentType` parâmetro. A `content` cadeia de caracteres também precisa ser formatada para inserir a menção no texto. O formato de uma menção é: `<at id="{replyIndex}">{mentionName}</at>`.

> [!NOTE]
> Atualmente, apenas o nome exato da menção pode ser usado como o texto do link de menção. O suporte para versões reduzidas de um nome será adicionado posteriormente.

O exemplo a seguir mostra um comentário com uma única menção.

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## <a name="comment-events"></a>Eventos de comentário

Seu complemento pode escutar adições, alterações e exclusões de comentários. [Eventos de comentário](/javascript/api/excel/excel.commentcollection#event-details) ocorrem no `CommentCollection` objeto. Para ouvir eventos de comentário, registre `onAdded`o manipulador de eventos , `onChanged`ou `onDeleted` comentário. Quando um evento de comentário é detectado, use esse manipulador de eventos para recuperar dados sobre o comentário adicionado, alterado ou excluído. O `onChanged` evento também lida com adições, alterações e exclusões de resposta de comentários.

Cada evento de comentário dispara apenas uma vez quando várias adições, alterações ou exclusões são executadas ao mesmo tempo. Todos os [objetos CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs) e [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contêm matrizes de IDs de comentário para mapear as ações de evento de volta para as coleções de comentários.

Consulte o [artigo Trabalhar com Eventos usando o Excel da API JavaScript](excel-add-ins-events.md) para obter informações adicionais sobre como registrar manipuladores de eventos, manipular eventos e remover manipuladores de eventos.

### <a name="comment-addition-events"></a>Eventos de adição de comentários

O `onAdded` evento é acionado quando um ou mais novos comentários são adicionados à coleção de comentários. Esse evento não *é* acionado quando as respostas são adicionadas a um thread de comentários (consulte [Eventos](#comment-change-events) de alteração de comentários para saber mais sobre eventos de resposta de comentários).

O exemplo a seguir mostra como registrar o manipulador `onAdded` de eventos e, em seguida, usar `CommentAddedEventArgs` o objeto para recuperar `commentDetails` a matriz do comentário adicionado.

> [!NOTE]
> Este exemplo só funciona quando um único comentário é adicionado.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-change-events"></a>Eventos de alteração de comentários

O `onChanged` evento comment é acionado nos seguintes cenários.

- O conteúdo de um comentário é atualizado.
- Um thread de comentário é resolvido.
- Um thread de comentário é reaberto.
- Uma resposta é adicionada a um thread de comentário.
- Uma resposta é atualizada em um thread de comentário.
- Uma resposta é excluída em um thread de comentário.

O exemplo a seguir mostra como registrar o manipulador `onChanged` de eventos e, em seguida, usar `CommentChangedEventArgs` o objeto para recuperar `commentDetails` a matriz do comentário alterado.

> [!NOTE]
> Este exemplo só funciona quando um único comentário é alterado.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-deletion-events"></a>Eventos de exclusão de comentários

O `onDeleted` evento é acionado quando um comentário é excluído da coleção de comentários. Depois que um comentário for excluído, seus metadados não estarão mais disponíveis. O [objeto CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fornece IDs de comentário, caso seu add-in gere comentários individuais.

O exemplo a seguir mostra como registrar o manipulador `onDeleted` de eventos e, em seguida, usar `CommentDeletedEventArgs` o objeto `commentDetails` para recuperar a matriz do comentário excluído.

> [!NOTE]
> Este exemplo só funciona quando um único comentário é excluído.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com pastas de trabalho usando a API JavaScript do Excel](excel-add-ins-workbooks.md)
- [Trabalhar com eventos usando a API JavaScript do Excel](excel-add-ins-events.md)
- [Inserir comentários e anotações no Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
