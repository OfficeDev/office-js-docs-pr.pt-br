---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 58325b88019988d55c565b5dc83efc25aca7266b
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127342"
---
# <a name="diagnostics"></a><span data-ttu-id="1bb0d-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="1bb0d-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="1bb0d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="1bb0d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="1bb0d-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1bb0d-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1bb0d-105">Requirements</span></span>

|<span data-ttu-id="1bb0d-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="1bb0d-106">Requirement</span></span>| <span data-ttu-id="1bb0d-107">Valor</span><span class="sxs-lookup"><span data-stu-id="1bb0d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1bb0d-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1bb0d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1bb0d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1bb0d-109">1.0</span></span>|
|[<span data-ttu-id="1bb0d-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1bb0d-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1bb0d-111">ReadItem</span></span>|
|[<span data-ttu-id="1bb0d-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1bb0d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1bb0d-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1bb0d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1bb0d-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="1bb0d-114">Members and methods</span></span>

| <span data-ttu-id="1bb0d-115">Membro</span><span class="sxs-lookup"><span data-stu-id="1bb0d-115">Member</span></span> | <span data-ttu-id="1bb0d-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1bb0d-117">hostName</span><span class="sxs-lookup"><span data-stu-id="1bb0d-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="1bb0d-118">Membro</span><span class="sxs-lookup"><span data-stu-id="1bb0d-118">Member</span></span> |
| [<span data-ttu-id="1bb0d-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="1bb0d-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="1bb0d-120">Membro</span><span class="sxs-lookup"><span data-stu-id="1bb0d-120">Member</span></span> |
| [<span data-ttu-id="1bb0d-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="1bb0d-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="1bb0d-122">Membro</span><span class="sxs-lookup"><span data-stu-id="1bb0d-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1bb0d-123">Membros</span><span class="sxs-lookup"><span data-stu-id="1bb0d-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="1bb0d-124">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1bb0d-124">hostName: String</span></span>

<span data-ttu-id="1bb0d-125">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="1bb0d-126">Uma cadeia de caracteres que pode ser um dos seguintes valores `Outlook`: `OutlookIOS`, ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="1bb0d-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-127">Type</span></span>

*   <span data-ttu-id="1bb0d-128">String</span><span class="sxs-lookup"><span data-stu-id="1bb0d-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1bb0d-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1bb0d-129">Requirements</span></span>

|<span data-ttu-id="1bb0d-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="1bb0d-130">Requirement</span></span>| <span data-ttu-id="1bb0d-131">Valor</span><span class="sxs-lookup"><span data-stu-id="1bb0d-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="1bb0d-132">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1bb0d-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1bb0d-133">1.0</span><span class="sxs-lookup"><span data-stu-id="1bb0d-133">1.0</span></span>|
|[<span data-ttu-id="1bb0d-134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1bb0d-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1bb0d-135">ReadItem</span></span>|
|[<span data-ttu-id="1bb0d-136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1bb0d-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1bb0d-137">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1bb0d-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="1bb0d-138">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1bb0d-138">hostVersion: String</span></span>

<span data-ttu-id="1bb0d-139">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="1bb0d-140">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou `hostVersion` Ios, a propriedade retornará a versão do aplicativo host, Outlook.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="1bb0d-141">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="1bb0d-142">Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-142">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="1bb0d-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-143">Type</span></span>

*   <span data-ttu-id="1bb0d-144">String</span><span class="sxs-lookup"><span data-stu-id="1bb0d-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1bb0d-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1bb0d-145">Requirements</span></span>

|<span data-ttu-id="1bb0d-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="1bb0d-146">Requirement</span></span>| <span data-ttu-id="1bb0d-147">Valor</span><span class="sxs-lookup"><span data-stu-id="1bb0d-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="1bb0d-148">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1bb0d-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1bb0d-149">1.0</span><span class="sxs-lookup"><span data-stu-id="1bb0d-149">1.0</span></span>|
|[<span data-ttu-id="1bb0d-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1bb0d-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1bb0d-151">ReadItem</span></span>|
|[<span data-ttu-id="1bb0d-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1bb0d-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1bb0d-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1bb0d-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="1bb0d-154">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1bb0d-154">OWAView: String</span></span>

<span data-ttu-id="1bb0d-155">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="1bb0d-156">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="1bb0d-157">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="1bb0d-158">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="1bb0d-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="1bb0d-159">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="1bb0d-160">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="1bb0d-161">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="1bb0d-162">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="1bb0d-163">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="1bb0d-164">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="1bb0d-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="1bb0d-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-165">Type</span></span>

*   <span data-ttu-id="1bb0d-166">String</span><span class="sxs-lookup"><span data-stu-id="1bb0d-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1bb0d-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1bb0d-167">Requirements</span></span>

|<span data-ttu-id="1bb0d-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="1bb0d-168">Requirement</span></span>| <span data-ttu-id="1bb0d-169">Valor</span><span class="sxs-lookup"><span data-stu-id="1bb0d-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="1bb0d-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1bb0d-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1bb0d-171">1.0</span><span class="sxs-lookup"><span data-stu-id="1bb0d-171">1.0</span></span>|
|[<span data-ttu-id="1bb0d-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1bb0d-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1bb0d-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1bb0d-173">ReadItem</span></span>|
|[<span data-ttu-id="1bb0d-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1bb0d-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1bb0d-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1bb0d-175">Compose or Read</span></span>|
