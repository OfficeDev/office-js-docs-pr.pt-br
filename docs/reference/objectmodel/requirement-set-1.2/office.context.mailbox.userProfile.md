---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7b7b9c7facd0542335094a42a3d1f53dab1f6aef
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814317"
---
# <a name="userprofile"></a><span data-ttu-id="fe33e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="fe33e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="fe33e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="fe33e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="fe33e-104">Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe33e-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe33e-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe33e-105">Requirements</span></span>

|<span data-ttu-id="fe33e-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe33e-106">Requirement</span></span>| <span data-ttu-id="fe33e-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fe33e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe33e-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe33e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fe33e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="fe33e-109">1.1</span></span>|
|[<span data-ttu-id="fe33e-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe33e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe33e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe33e-111">ReadItem</span></span>|
|[<span data-ttu-id="fe33e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe33e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fe33e-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fe33e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="fe33e-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="fe33e-114">Properties</span></span>

| <span data-ttu-id="fe33e-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="fe33e-115">Property</span></span> | <span data-ttu-id="fe33e-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="fe33e-116">Minimum</span></span><br><span data-ttu-id="fe33e-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="fe33e-117">permission level</span></span> | <span data-ttu-id="fe33e-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="fe33e-118">Modes</span></span> | <span data-ttu-id="fe33e-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="fe33e-119">Return type</span></span> | <span data-ttu-id="fe33e-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="fe33e-120">Minimum</span></span><br><span data-ttu-id="fe33e-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="fe33e-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="fe33e-122">displayName</span><span class="sxs-lookup"><span data-stu-id="fe33e-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#displayname) | <span data-ttu-id="fe33e-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe33e-123">ReadItem</span></span> | <span data-ttu-id="fe33e-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="fe33e-124">Compose</span></span><br><span data-ttu-id="fe33e-125">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe33e-125">Read</span></span> | <span data-ttu-id="fe33e-126">String</span><span class="sxs-lookup"><span data-stu-id="fe33e-126">String</span></span> | [<span data-ttu-id="fe33e-127">1.1</span><span class="sxs-lookup"><span data-stu-id="fe33e-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fe33e-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="fe33e-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#emailaddress) | <span data-ttu-id="fe33e-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe33e-129">ReadItem</span></span> | <span data-ttu-id="fe33e-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="fe33e-130">Compose</span></span><br><span data-ttu-id="fe33e-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe33e-131">Read</span></span> | <span data-ttu-id="fe33e-132">String</span><span class="sxs-lookup"><span data-stu-id="fe33e-132">String</span></span> | [<span data-ttu-id="fe33e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="fe33e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fe33e-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="fe33e-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#timezone) | <span data-ttu-id="fe33e-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe33e-135">ReadItem</span></span> | <span data-ttu-id="fe33e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="fe33e-136">Compose</span></span><br><span data-ttu-id="fe33e-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe33e-137">Read</span></span> | <span data-ttu-id="fe33e-138">String</span><span class="sxs-lookup"><span data-stu-id="fe33e-138">String</span></span> | [<span data-ttu-id="fe33e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="fe33e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
