---
title: Obter e definir categorias
description: Como gerenciar categorias de caixa de correio e item
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d4589571de47218741308c01caec0166d72919d8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608975"
---
# <a name="get-and-set-categories"></a>Obter e definir categorias

No Outlook, um usuário pode aplicar categorias a mensagens e compromissos como uma maneira de organizar seus dados de caixa de correio. O usuário define a lista mestra de categorias codificadas por cores para a caixa de correio e pode aplicar uma ou mais dessas categorias a qualquer item de compromisso ou mensagem. Cada [categoria](/javascript/api/outlook/office.categorydetails) na lista mestra é representada pelo nome e pela [cor](/javascript/api/outlook/office.mailboxenums.categorycolor) que o usuário especifica. Você pode usar a API JavaScript do Office para gerenciar a lista mestra de categorias na caixa de correio e as categorias aplicadas a um item.

> [!NOTE]
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="manage-categories-in-the-master-list"></a>Gerenciar categorias na lista mestra

Somente as categorias na lista mestra da caixa de correio estão disponíveis para que você se aplique a uma mensagem ou um compromisso. Você pode usar a API para adicionar, obter e remover categorias mestras.

> [!IMPORTANT]
> Para o suplemento gerenciar a lista mestra de categorias, você deve definir o `Permissions` nó no manifesto `ReadWriteMailbox` .

### <a name="add-master-categories"></a>Adicionar categorias mestras

O exemplo a seguir mostra como adicionar uma categoria chamada "urgente!" para a lista mestra chamando [addasync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) em [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a>Obter categorias mestre

O exemplo a seguir mostra como obter a lista de categorias chamando [getasync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) no [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>Remover categorias mestre

O exemplo a seguir mostra como remover a categoria chamada "urgente!" na lista mestra chamando [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) em [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>Gerenciar categorias em uma mensagem ou compromisso

Você pode usar a API para adicionar, obter e remover categorias de um item de compromisso ou mensagem.

> [!IMPORTANT]
> Somente as categorias na lista mestra da caixa de correio estão disponíveis para que você se aplique a uma mensagem ou um compromisso. Consulte a seção anterior [gerenciar categorias na lista mestra](#manage-categories-in-the-master-list) para obter mais informações.
>
> No Outlook na Web, você não pode usar a API para gerenciar categorias em uma mensagem no modo de leitura.

### <a name="add-categories-to-an-item"></a>Adicionar categorias a um item

O exemplo a seguir mostra como aplicar a categoria chamada "urgente!" para o item atual chamando [addasync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories` .

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a>Obter categorias de um item

O exemplo a seguir mostra como obter as categorias aplicadas ao item atual chamando [getasync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories` .

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>Remover categorias de um item

O exemplo a seguir mostra como remover a categoria chamada "urgente!" do item atual chamando [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) `item.categories` .

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a>Confira também

- [Permissões do Outlook](understanding-outlook-add-in-permissions.md)
- [Elemento Permissions no manifesto](../reference/manifest/permissions.md)
