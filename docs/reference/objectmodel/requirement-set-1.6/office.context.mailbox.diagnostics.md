---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ee10e511ed81a591e5e7b89c7650e16fca27da09
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814735"
---
# <a name="diagnostics"></a><span data-ttu-id="551c7-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="551c7-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="551c7-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="551c7-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="551c7-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="551c7-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="551c7-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="551c7-105">Requirements</span></span>

|<span data-ttu-id="551c7-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="551c7-106">Requirement</span></span>| <span data-ttu-id="551c7-107">Valor</span><span class="sxs-lookup"><span data-stu-id="551c7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="551c7-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="551c7-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="551c7-109">1.1</span><span class="sxs-lookup"><span data-stu-id="551c7-109">1.1</span></span>|
|[<span data-ttu-id="551c7-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="551c7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="551c7-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="551c7-111">ReadItem</span></span>|
|[<span data-ttu-id="551c7-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="551c7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="551c7-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="551c7-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="551c7-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="551c7-114">Properties</span></span>

| <span data-ttu-id="551c7-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="551c7-115">Property</span></span> | <span data-ttu-id="551c7-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="551c7-116">Minimum</span></span><br><span data-ttu-id="551c7-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="551c7-117">permission level</span></span> | <span data-ttu-id="551c7-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="551c7-118">Modes</span></span> | <span data-ttu-id="551c7-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="551c7-119">Return type</span></span> | <span data-ttu-id="551c7-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="551c7-120">Minimum</span></span><br><span data-ttu-id="551c7-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="551c7-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="551c7-122">hostName</span><span class="sxs-lookup"><span data-stu-id="551c7-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#hostname) | <span data-ttu-id="551c7-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="551c7-123">ReadItem</span></span> | <span data-ttu-id="551c7-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="551c7-124">Compose</span></span><br><span data-ttu-id="551c7-125">Leitura</span><span class="sxs-lookup"><span data-stu-id="551c7-125">Read</span></span> | <span data-ttu-id="551c7-126">String</span><span class="sxs-lookup"><span data-stu-id="551c7-126">String</span></span> | [<span data-ttu-id="551c7-127">1.1</span><span class="sxs-lookup"><span data-stu-id="551c7-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="551c7-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="551c7-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#hostversion) | <span data-ttu-id="551c7-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="551c7-129">ReadItem</span></span> | <span data-ttu-id="551c7-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="551c7-130">Compose</span></span><br><span data-ttu-id="551c7-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="551c7-131">Read</span></span> | <span data-ttu-id="551c7-132">String</span><span class="sxs-lookup"><span data-stu-id="551c7-132">String</span></span> | [<span data-ttu-id="551c7-133">1.1</span><span class="sxs-lookup"><span data-stu-id="551c7-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="551c7-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="551c7-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6#owaview) | <span data-ttu-id="551c7-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="551c7-135">ReadItem</span></span> | <span data-ttu-id="551c7-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="551c7-136">Compose</span></span><br><span data-ttu-id="551c7-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="551c7-137">Read</span></span> | <span data-ttu-id="551c7-138">String</span><span class="sxs-lookup"><span data-stu-id="551c7-138">String</span></span> | [<span data-ttu-id="551c7-139">1.1</span><span class="sxs-lookup"><span data-stu-id="551c7-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
