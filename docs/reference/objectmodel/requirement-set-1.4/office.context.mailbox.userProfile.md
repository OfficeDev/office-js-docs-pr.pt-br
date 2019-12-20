---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 6b5229c1bc300d11714f3aa2cf8fa8ff2465667c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814261"
---
# <a name="userprofile"></a><span data-ttu-id="82ce9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="82ce9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="82ce9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="82ce9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="82ce9-104">Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="82ce9-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="82ce9-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82ce9-105">Requirements</span></span>

|<span data-ttu-id="82ce9-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="82ce9-106">Requirement</span></span>| <span data-ttu-id="82ce9-107">Valor</span><span class="sxs-lookup"><span data-stu-id="82ce9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="82ce9-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82ce9-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82ce9-109">1.1</span><span class="sxs-lookup"><span data-stu-id="82ce9-109">1.1</span></span>|
|[<span data-ttu-id="82ce9-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="82ce9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="82ce9-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="82ce9-111">ReadItem</span></span>|
|[<span data-ttu-id="82ce9-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82ce9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82ce9-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82ce9-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="82ce9-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="82ce9-114">Properties</span></span>

| <span data-ttu-id="82ce9-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="82ce9-115">Property</span></span> | <span data-ttu-id="82ce9-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="82ce9-116">Minimum</span></span><br><span data-ttu-id="82ce9-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="82ce9-117">permission level</span></span> | <span data-ttu-id="82ce9-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="82ce9-118">Modes</span></span> | <span data-ttu-id="82ce9-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="82ce9-119">Return type</span></span> | <span data-ttu-id="82ce9-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="82ce9-120">Minimum</span></span><br><span data-ttu-id="82ce9-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="82ce9-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="82ce9-122">displayName</span><span class="sxs-lookup"><span data-stu-id="82ce9-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="82ce9-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="82ce9-123">ReadItem</span></span> | <span data-ttu-id="82ce9-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="82ce9-124">Compose</span></span><br><span data-ttu-id="82ce9-125">Leitura</span><span class="sxs-lookup"><span data-stu-id="82ce9-125">Read</span></span> | <span data-ttu-id="82ce9-126">String</span><span class="sxs-lookup"><span data-stu-id="82ce9-126">String</span></span> | [<span data-ttu-id="82ce9-127">1.1</span><span class="sxs-lookup"><span data-stu-id="82ce9-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82ce9-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="82ce9-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="82ce9-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="82ce9-129">ReadItem</span></span> | <span data-ttu-id="82ce9-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="82ce9-130">Compose</span></span><br><span data-ttu-id="82ce9-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="82ce9-131">Read</span></span> | <span data-ttu-id="82ce9-132">String</span><span class="sxs-lookup"><span data-stu-id="82ce9-132">String</span></span> | [<span data-ttu-id="82ce9-133">1.1</span><span class="sxs-lookup"><span data-stu-id="82ce9-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82ce9-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="82ce9-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="82ce9-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="82ce9-135">ReadItem</span></span> | <span data-ttu-id="82ce9-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="82ce9-136">Compose</span></span><br><span data-ttu-id="82ce9-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="82ce9-137">Read</span></span> | <span data-ttu-id="82ce9-138">String</span><span class="sxs-lookup"><span data-stu-id="82ce9-138">String</span></span> | [<span data-ttu-id="82ce9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="82ce9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
