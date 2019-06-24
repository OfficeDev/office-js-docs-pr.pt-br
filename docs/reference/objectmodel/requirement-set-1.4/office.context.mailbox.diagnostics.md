---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: edbaa100ba82b0dd1077e518c1090c07890a3d41
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127405"
---
# <a name="diagnostics"></a><span data-ttu-id="3d8c9-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="3d8c9-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="3d8c9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="3d8c9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="3d8c9-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d8c9-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3d8c9-105">Requirements</span></span>

|<span data-ttu-id="3d8c9-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="3d8c9-106">Requirement</span></span>| <span data-ttu-id="3d8c9-107">Valor</span><span class="sxs-lookup"><span data-stu-id="3d8c9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d8c9-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3d8c9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d8c9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3d8c9-109">1.0</span></span>|
|[<span data-ttu-id="3d8c9-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d8c9-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d8c9-111">ReadItem</span></span>|
|[<span data-ttu-id="3d8c9-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3d8c9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d8c9-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3d8c9-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="3d8c9-114">Members</span><span class="sxs-lookup"><span data-stu-id="3d8c9-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="3d8c9-115">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3d8c9-115">hostName: String</span></span>

<span data-ttu-id="3d8c9-116">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="3d8c9-117">Uma cadeia de caracteres que pode ser um dos seguintes valores `Outlook`: `OutlookIOS`, ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="3d8c9-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-118">Type</span></span>

*   <span data-ttu-id="3d8c9-119">String</span><span class="sxs-lookup"><span data-stu-id="3d8c9-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d8c9-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3d8c9-120">Requirements</span></span>

|<span data-ttu-id="3d8c9-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="3d8c9-121">Requirement</span></span>| <span data-ttu-id="3d8c9-122">Valor</span><span class="sxs-lookup"><span data-stu-id="3d8c9-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d8c9-123">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3d8c9-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d8c9-124">1.0</span><span class="sxs-lookup"><span data-stu-id="3d8c9-124">1.0</span></span>|
|[<span data-ttu-id="3d8c9-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d8c9-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d8c9-126">ReadItem</span></span>|
|[<span data-ttu-id="3d8c9-127">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3d8c9-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d8c9-128">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3d8c9-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="3d8c9-129">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3d8c9-129">hostVersion: String</span></span>

<span data-ttu-id="3d8c9-130">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="3d8c9-131">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou `hostVersion` Ios, a propriedade retornará a versão do aplicativo host, Outlook.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="3d8c9-132">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="3d8c9-133">Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-133">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="3d8c9-134">Tipo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-134">Type</span></span>

*   <span data-ttu-id="3d8c9-135">String</span><span class="sxs-lookup"><span data-stu-id="3d8c9-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d8c9-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3d8c9-136">Requirements</span></span>

|<span data-ttu-id="3d8c9-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="3d8c9-137">Requirement</span></span>| <span data-ttu-id="3d8c9-138">Valor</span><span class="sxs-lookup"><span data-stu-id="3d8c9-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d8c9-139">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3d8c9-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d8c9-140">1.0</span><span class="sxs-lookup"><span data-stu-id="3d8c9-140">1.0</span></span>|
|[<span data-ttu-id="3d8c9-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d8c9-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d8c9-142">ReadItem</span></span>|
|[<span data-ttu-id="3d8c9-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3d8c9-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d8c9-144">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3d8c9-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="3d8c9-145">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3d8c9-145">OWAView: String</span></span>

<span data-ttu-id="3d8c9-146">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="3d8c9-147">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="3d8c9-148">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="3d8c9-149">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="3d8c9-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="3d8c9-150">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="3d8c9-151">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="3d8c9-152">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="3d8c9-153">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="3d8c9-154">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="3d8c9-155">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="3d8c9-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="3d8c9-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-156">Type</span></span>

*   <span data-ttu-id="3d8c9-157">String</span><span class="sxs-lookup"><span data-stu-id="3d8c9-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d8c9-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3d8c9-158">Requirements</span></span>

|<span data-ttu-id="3d8c9-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="3d8c9-159">Requirement</span></span>| <span data-ttu-id="3d8c9-160">Valor</span><span class="sxs-lookup"><span data-stu-id="3d8c9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d8c9-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3d8c9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d8c9-162">1.0</span><span class="sxs-lookup"><span data-stu-id="3d8c9-162">1.0</span></span>|
|[<span data-ttu-id="3d8c9-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3d8c9-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d8c9-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d8c9-164">ReadItem</span></span>|
|[<span data-ttu-id="3d8c9-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3d8c9-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d8c9-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3d8c9-166">Compose or Read</span></span>|
