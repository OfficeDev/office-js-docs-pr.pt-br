---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 1bf24eb39329be0139957cc6e0f8629fb9f3b166
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815015"
---
# <a name="userprofile"></a><span data-ttu-id="85e78-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="85e78-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="85e78-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="85e78-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="85e78-104">Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="85e78-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="85e78-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="85e78-105">Requirements</span></span>

|<span data-ttu-id="85e78-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="85e78-106">Requirement</span></span>| <span data-ttu-id="85e78-107">Valor</span><span class="sxs-lookup"><span data-stu-id="85e78-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="85e78-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="85e78-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="85e78-109">1.1</span><span class="sxs-lookup"><span data-stu-id="85e78-109">1.1</span></span>|
|[<span data-ttu-id="85e78-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="85e78-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85e78-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85e78-111">ReadItem</span></span>|
|[<span data-ttu-id="85e78-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="85e78-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="85e78-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="85e78-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="85e78-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="85e78-114">Properties</span></span>

| <span data-ttu-id="85e78-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="85e78-115">Property</span></span> | <span data-ttu-id="85e78-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="85e78-116">Minimum</span></span><br><span data-ttu-id="85e78-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="85e78-117">permission level</span></span> | <span data-ttu-id="85e78-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="85e78-118">Modes</span></span> | <span data-ttu-id="85e78-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="85e78-119">Return type</span></span> | <span data-ttu-id="85e78-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="85e78-120">Minimum</span></span><br><span data-ttu-id="85e78-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="85e78-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="85e78-122">displayName</span><span class="sxs-lookup"><span data-stu-id="85e78-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#displayname) | <span data-ttu-id="85e78-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85e78-123">ReadItem</span></span> | <span data-ttu-id="85e78-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="85e78-124">Compose</span></span><br><span data-ttu-id="85e78-125">Leitura</span><span class="sxs-lookup"><span data-stu-id="85e78-125">Read</span></span> | <span data-ttu-id="85e78-126">String</span><span class="sxs-lookup"><span data-stu-id="85e78-126">String</span></span> | [<span data-ttu-id="85e78-127">1.1</span><span class="sxs-lookup"><span data-stu-id="85e78-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="85e78-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="85e78-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#emailaddress) | <span data-ttu-id="85e78-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85e78-129">ReadItem</span></span> | <span data-ttu-id="85e78-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="85e78-130">Compose</span></span><br><span data-ttu-id="85e78-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="85e78-131">Read</span></span> | <span data-ttu-id="85e78-132">String</span><span class="sxs-lookup"><span data-stu-id="85e78-132">String</span></span> | [<span data-ttu-id="85e78-133">1.1</span><span class="sxs-lookup"><span data-stu-id="85e78-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="85e78-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="85e78-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1#timezone) | <span data-ttu-id="85e78-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85e78-135">ReadItem</span></span> | <span data-ttu-id="85e78-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="85e78-136">Compose</span></span><br><span data-ttu-id="85e78-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="85e78-137">Read</span></span> | <span data-ttu-id="85e78-138">String</span><span class="sxs-lookup"><span data-stu-id="85e78-138">String</span></span> | [<span data-ttu-id="85e78-139">1.1</span><span class="sxs-lookup"><span data-stu-id="85e78-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
