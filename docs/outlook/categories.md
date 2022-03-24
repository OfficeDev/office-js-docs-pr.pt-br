---
title: Obter e definir categorias
description: Como gerenciar categorias na caixa de correio e no item.
ms.date: 01/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: 93f9167fcc31110543d08019e5428952beab0ccc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746297"
---
# <a name="get-and-set-categories"></a>Obter e definir categorias

Em Outlook, um usuário pode aplicar categorias a mensagens e compromissos como um meio de organizar seus dados de caixa de correio. O usuário define a lista mestra de categorias codificadas por cores para sua caixa de correio e, em seguida, pode aplicar uma ou mais dessas categorias a qualquer mensagem ou item de compromisso. Cada [categoria](/javascript/api/outlook/office.categorydetails) na lista mestra é representada pelo nome e [cor](/javascript/api/outlook/office.mailboxenums.categorycolor) especificados pelo usuário. Você pode usar a api Office JavaScript para gerenciar a lista mestra categorias na caixa de correio e as categorias aplicadas a um item.

> [!NOTE]
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.8. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="manage-categories-in-the-master-list"></a>Gerenciar categorias na lista mestra

Somente categorias na lista mestra em sua caixa de correio estão disponíveis para você aplicar a uma mensagem ou compromisso. Você pode usar a API para adicionar, obter e remover categorias mestras.

> [!IMPORTANT]
> Para que o add-in gerencie a lista mestra de categorias, você deve definir `Permissions` o nó no manifesto como `ReadWriteMailbox`.

### <a name="add-master-categories"></a>Adicionar categorias mestras

O exemplo a seguir mostra como adicionar uma categoria chamada "Urgente!" para a lista mestra chamando [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) em [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

### <a name="get-master-categories"></a>Obter categorias mestras

O exemplo a seguir mostra como obter a lista de categorias chamando [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) em [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

### <a name="remove-master-categories"></a>Remover categorias mestras

O exemplo a seguir mostra como remover a categoria chamada "Urgente!" da lista mestra chamando [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) em [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

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

Você pode usar a API para adicionar, obter e remover categorias para um item de mensagem ou compromisso.

> [!IMPORTANT]
> Somente categorias na lista mestra em sua caixa de correio estão disponíveis para você aplicar a uma mensagem ou compromisso. Consulte a seção anterior [Gerenciar categorias na lista mestra](#manage-categories-in-the-master-list) para obter mais informações.
>
> No Outlook na Web, você não pode usar a API para gerenciar categorias em uma mensagem no modo De leitura.

### <a name="add-categories-to-an-item"></a>Adicionar categorias a um item

O exemplo a seguir mostra como aplicar a categoria chamada "Urgente!" para o item atual chamando [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) em `item.categories`.

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

O exemplo a seguir mostra como obter as categorias aplicadas ao item atual chamando [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) em `item.categories`.

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

O exemplo a seguir mostra como remover a categoria chamada "Urgente!" do item atual chamando [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) em `item.categories`.

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

- [Outlook permissões](understanding-outlook-add-in-permissions.md)
- [Elemento Permissions no manifesto](../reference/manifest/permissions.md)
