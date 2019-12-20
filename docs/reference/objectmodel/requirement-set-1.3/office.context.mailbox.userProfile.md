---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 4d63dfe1b32de2ac7fe55f324f938b85a865ec02
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814910"
---
# <a name="userprofile"></a><span data-ttu-id="cd1a6-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="cd1a6-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="cd1a6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="cd1a6-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="cd1a6-104">Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="cd1a6-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cd1a6-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd1a6-105">Requirements</span></span>

|<span data-ttu-id="cd1a6-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd1a6-106">Requirement</span></span>| <span data-ttu-id="cd1a6-107">Valor</span><span class="sxs-lookup"><span data-stu-id="cd1a6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd1a6-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd1a6-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd1a6-109">1.1</span><span class="sxs-lookup"><span data-stu-id="cd1a6-109">1.1</span></span>|
|[<span data-ttu-id="cd1a6-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="cd1a6-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cd1a6-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cd1a6-111">ReadItem</span></span>|
|[<span data-ttu-id="cd1a6-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd1a6-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cd1a6-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd1a6-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="cd1a6-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="cd1a6-114">Properties</span></span>

| <span data-ttu-id="cd1a6-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cd1a6-115">Property</span></span> | <span data-ttu-id="cd1a6-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="cd1a6-116">Minimum</span></span><br><span data-ttu-id="cd1a6-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="cd1a6-117">permission level</span></span> | <span data-ttu-id="cd1a6-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="cd1a6-118">Modes</span></span> | <span data-ttu-id="cd1a6-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="cd1a6-119">Return type</span></span> | <span data-ttu-id="cd1a6-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="cd1a6-120">Minimum</span></span><br><span data-ttu-id="cd1a6-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="cd1a6-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="cd1a6-122">displayName</span><span class="sxs-lookup"><span data-stu-id="cd1a6-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#displayname) | <span data-ttu-id="cd1a6-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cd1a6-123">ReadItem</span></span> | <span data-ttu-id="cd1a6-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd1a6-124">Compose</span></span><br><span data-ttu-id="cd1a6-125">Leitura</span><span class="sxs-lookup"><span data-stu-id="cd1a6-125">Read</span></span> | <span data-ttu-id="cd1a6-126">String</span><span class="sxs-lookup"><span data-stu-id="cd1a6-126">String</span></span> | [<span data-ttu-id="cd1a6-127">1.1</span><span class="sxs-lookup"><span data-stu-id="cd1a6-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd1a6-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="cd1a6-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#emailaddress) | <span data-ttu-id="cd1a6-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cd1a6-129">ReadItem</span></span> | <span data-ttu-id="cd1a6-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd1a6-130">Compose</span></span><br><span data-ttu-id="cd1a6-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="cd1a6-131">Read</span></span> | <span data-ttu-id="cd1a6-132">String</span><span class="sxs-lookup"><span data-stu-id="cd1a6-132">String</span></span> | [<span data-ttu-id="cd1a6-133">1.1</span><span class="sxs-lookup"><span data-stu-id="cd1a6-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd1a6-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="cd1a6-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#timezone) | <span data-ttu-id="cd1a6-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cd1a6-135">ReadItem</span></span> | <span data-ttu-id="cd1a6-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd1a6-136">Compose</span></span><br><span data-ttu-id="cd1a6-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="cd1a6-137">Read</span></span> | <span data-ttu-id="cd1a6-138">String</span><span class="sxs-lookup"><span data-stu-id="cd1a6-138">String</span></span> | [<span data-ttu-id="cd1a6-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cd1a6-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
