---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950989"
---
# <a name="userprofile"></a><span data-ttu-id="d58de-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d58de-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d58de-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d58de-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="d58de-104">Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d58de-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d58de-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d58de-105">Requirements</span></span>

|<span data-ttu-id="d58de-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="d58de-106">Requirement</span></span>| <span data-ttu-id="d58de-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d58de-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d58de-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d58de-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d58de-109">1.1</span><span class="sxs-lookup"><span data-stu-id="d58de-109">1.1</span></span>|
|[<span data-ttu-id="d58de-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d58de-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d58de-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d58de-111">ReadItem</span></span>|
|[<span data-ttu-id="d58de-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d58de-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d58de-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d58de-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="d58de-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="d58de-114">Properties</span></span>

| <span data-ttu-id="d58de-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="d58de-115">Property</span></span> | <span data-ttu-id="d58de-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="d58de-116">Minimum</span></span><br><span data-ttu-id="d58de-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="d58de-117">permission level</span></span> | <span data-ttu-id="d58de-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="d58de-118">Modes</span></span> | <span data-ttu-id="d58de-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="d58de-119">Return type</span></span> | <span data-ttu-id="d58de-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="d58de-120">Minimum</span></span><br><span data-ttu-id="d58de-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="d58de-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="d58de-122">displayName</span><span class="sxs-lookup"><span data-stu-id="d58de-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="d58de-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d58de-123">ReadItem</span></span> | <span data-ttu-id="d58de-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="d58de-124">Compose</span></span><br><span data-ttu-id="d58de-125">Ler</span><span class="sxs-lookup"><span data-stu-id="d58de-125">Read</span></span> | <span data-ttu-id="d58de-126">String</span><span class="sxs-lookup"><span data-stu-id="d58de-126">String</span></span> | [<span data-ttu-id="d58de-127">1.1</span><span class="sxs-lookup"><span data-stu-id="d58de-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d58de-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d58de-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="d58de-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d58de-129">ReadItem</span></span> | <span data-ttu-id="d58de-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="d58de-130">Compose</span></span><br><span data-ttu-id="d58de-131">Ler</span><span class="sxs-lookup"><span data-stu-id="d58de-131">Read</span></span> | <span data-ttu-id="d58de-132">String</span><span class="sxs-lookup"><span data-stu-id="d58de-132">String</span></span> | [<span data-ttu-id="d58de-133">1.1</span><span class="sxs-lookup"><span data-stu-id="d58de-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d58de-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="d58de-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="d58de-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d58de-135">ReadItem</span></span> | <span data-ttu-id="d58de-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="d58de-136">Compose</span></span><br><span data-ttu-id="d58de-137">Ler</span><span class="sxs-lookup"><span data-stu-id="d58de-137">Read</span></span> | <span data-ttu-id="d58de-138">String</span><span class="sxs-lookup"><span data-stu-id="d58de-138">String</span></span> | [<span data-ttu-id="d58de-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d58de-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
