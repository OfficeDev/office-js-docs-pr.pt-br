---
title: Obter e definir categorias
description: Como gerenciar categorias de caixa de correio e item
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 50b98191661674b50c5636733075e4a882183d82
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165836"
---
# <a name="get-and-set-categories"></a><span data-ttu-id="6bee1-103">Obter e definir categorias</span><span class="sxs-lookup"><span data-stu-id="6bee1-103">Get and set categories</span></span>

<span data-ttu-id="6bee1-104">No Outlook, um usuário pode aplicar categorias a mensagens e compromissos como uma maneira de organizar seus dados de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="6bee1-104">In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data.</span></span> <span data-ttu-id="6bee1-105">O usuário define a lista mestra de categorias codificadas por cores para a caixa de correio e pode aplicar uma ou mais dessas categorias a qualquer item de compromisso ou mensagem.</span><span class="sxs-lookup"><span data-stu-id="6bee1-105">The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item.</span></span> <span data-ttu-id="6bee1-106">Cada [categoria](/javascript/api/outlook/office.categorydetails) na lista mestra é representada pelo nome e pela [cor](/javascript/api/outlook/office.mailboxenums.categorycolor) que o usuário especifica.</span><span class="sxs-lookup"><span data-stu-id="6bee1-106">Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies.</span></span> <span data-ttu-id="6bee1-107">Você pode usar a API JavaScript do Office para gerenciar a lista mestra de categorias na caixa de correio e as categorias aplicadas a um item.</span><span class="sxs-lookup"><span data-stu-id="6bee1-107">You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="6bee1-108">O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="6bee1-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="6bee1-109">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="6bee1-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="manage-categories-in-the-master-list"></a><span data-ttu-id="6bee1-110">Gerenciar categorias na lista mestra</span><span class="sxs-lookup"><span data-stu-id="6bee1-110">Manage categories in the master list</span></span>

<span data-ttu-id="6bee1-111">Somente as categorias na lista mestra da caixa de correio estão disponíveis para que você se aplique a uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="6bee1-111">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="6bee1-112">Você pode usar a API para adicionar, obter e remover categorias mestras.</span><span class="sxs-lookup"><span data-stu-id="6bee1-112">You can use the API to add, get, and remove master categories.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6bee1-113">Para o suplemento gerenciar a lista mestra de categorias, você deve definir o `Permissions` nó no manifesto. `ReadWriteMailbox`</span><span class="sxs-lookup"><span data-stu-id="6bee1-113">For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.</span></span>

### <a name="add-master-categories"></a><span data-ttu-id="6bee1-114">Adicionar categorias mestras</span><span class="sxs-lookup"><span data-stu-id="6bee1-114">Add master categories</span></span>

<span data-ttu-id="6bee1-115">O exemplo a seguir mostra como adicionar uma categoria chamada "urgente!"</span><span class="sxs-lookup"><span data-stu-id="6bee1-115">The following example shows how to add a category named "Urgent!"</span></span> <span data-ttu-id="6bee1-116">para a lista mestra chamando [addasync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) em [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="6bee1-116">to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

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

### <a name="get-master-categories"></a><span data-ttu-id="6bee1-117">Obter categorias mestre</span><span class="sxs-lookup"><span data-stu-id="6bee1-117">Get master categories</span></span>

<span data-ttu-id="6bee1-118">O exemplo a seguir mostra como obter a lista de categorias chamando [getasync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) no [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="6bee1-118">The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

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

### <a name="remove-master-categories"></a><span data-ttu-id="6bee1-119">Remover categorias mestre</span><span class="sxs-lookup"><span data-stu-id="6bee1-119">Remove master categories</span></span>

<span data-ttu-id="6bee1-120">O exemplo a seguir mostra como remover a categoria chamada "urgente!"</span><span class="sxs-lookup"><span data-stu-id="6bee1-120">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="6bee1-121">na lista mestra chamando [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) em [Mailbox. nova mastercategories](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="6bee1-121">from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

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

## <a name="manage-categories-on-a-message-or-appointment"></a><span data-ttu-id="6bee1-122">Gerenciar categorias em uma mensagem ou compromisso</span><span class="sxs-lookup"><span data-stu-id="6bee1-122">Manage categories on a message or appointment</span></span>

<span data-ttu-id="6bee1-123">Você pode usar a API para adicionar, obter e remover categorias de um item de compromisso ou mensagem.</span><span class="sxs-lookup"><span data-stu-id="6bee1-123">You can use the API to add, get, and remove categories for a message or appointment item.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6bee1-124">Somente as categorias na lista mestra da caixa de correio estão disponíveis para que você se aplique a uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="6bee1-124">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="6bee1-125">Consulte a seção anterior [gerenciar categorias na lista mestra](#manage-categories-in-the-master-list) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6bee1-125">See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.</span></span>
>
> <span data-ttu-id="6bee1-126">No Outlook na Web, você não pode usar a API para gerenciar categorias em uma mensagem no modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="6bee1-126">In Outlook on the web, you can't use the API to manage categories on a message in Read mode.</span></span>

### <a name="add-categories-to-an-item"></a><span data-ttu-id="6bee1-127">Adicionar categorias a um item</span><span class="sxs-lookup"><span data-stu-id="6bee1-127">Add categories to an item</span></span>

<span data-ttu-id="6bee1-128">O exemplo a seguir mostra como aplicar a categoria chamada "urgente!"</span><span class="sxs-lookup"><span data-stu-id="6bee1-128">The following example shows how to apply the category named "Urgent!"</span></span> <span data-ttu-id="6bee1-129">para o item atual chamando [addasync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span><span class="sxs-lookup"><span data-stu-id="6bee1-129">to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span></span>

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

### <a name="get-an-items-categories"></a><span data-ttu-id="6bee1-130">Obter categorias de um item</span><span class="sxs-lookup"><span data-stu-id="6bee1-130">Get an item's categories</span></span>

<span data-ttu-id="6bee1-131">O exemplo a seguir mostra como obter as categorias aplicadas ao item atual chamando [getasync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span><span class="sxs-lookup"><span data-stu-id="6bee1-131">The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span></span>

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

### <a name="remove-categories-from-an-item"></a><span data-ttu-id="6bee1-132">Remover categorias de um item</span><span class="sxs-lookup"><span data-stu-id="6bee1-132">Remove categories from an item</span></span>

<span data-ttu-id="6bee1-133">O exemplo a seguir mostra como remover a categoria chamada "urgente!"</span><span class="sxs-lookup"><span data-stu-id="6bee1-133">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="6bee1-134">do item atual chamando [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) `item.categories`.</span><span class="sxs-lookup"><span data-stu-id="6bee1-134">from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6bee1-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="6bee1-135">See also</span></span>

- [<span data-ttu-id="6bee1-136">Permissões do Outlook</span><span class="sxs-lookup"><span data-stu-id="6bee1-136">Outlook permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="6bee1-137">Elemento Permissions no manifesto</span><span class="sxs-lookup"><span data-stu-id="6bee1-137">Permissions element in the manifest</span></span>](../reference/manifest/permissions.md)
