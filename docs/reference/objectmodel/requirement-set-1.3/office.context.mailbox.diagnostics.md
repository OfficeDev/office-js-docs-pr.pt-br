---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 46f58490d3f16f2a1bcd4caf0a21734861477e9f
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231274"
---
# <a name="diagnostics"></a><span data-ttu-id="9032b-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9032b-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="9032b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="9032b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="9032b-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9032b-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9032b-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9032b-105">Requirements</span></span>

|<span data-ttu-id="9032b-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="9032b-106">Requirement</span></span>| <span data-ttu-id="9032b-107">Valor</span><span class="sxs-lookup"><span data-stu-id="9032b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9032b-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9032b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9032b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9032b-109">1.0</span></span>|
|[<span data-ttu-id="9032b-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9032b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9032b-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9032b-111">ReadItem</span></span>|
|[<span data-ttu-id="9032b-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9032b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9032b-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9032b-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="9032b-114">Members</span><span class="sxs-lookup"><span data-stu-id="9032b-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="9032b-115">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9032b-115">hostName: String</span></span>

<span data-ttu-id="9032b-116">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="9032b-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="9032b-117">Uma cadeia de caracteres que pode ser um dos seguintes valores `Outlook`: `OutlookIOS`, ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="9032b-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="9032b-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="9032b-118">Type</span></span>

*   <span data-ttu-id="9032b-119">String</span><span class="sxs-lookup"><span data-stu-id="9032b-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9032b-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9032b-120">Requirements</span></span>

|<span data-ttu-id="9032b-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="9032b-121">Requirement</span></span>| <span data-ttu-id="9032b-122">Valor</span><span class="sxs-lookup"><span data-stu-id="9032b-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="9032b-123">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9032b-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9032b-124">1.0</span><span class="sxs-lookup"><span data-stu-id="9032b-124">1.0</span></span>|
|[<span data-ttu-id="9032b-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9032b-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9032b-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9032b-126">ReadItem</span></span>|
|[<span data-ttu-id="9032b-127">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9032b-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9032b-128">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9032b-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="9032b-129">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9032b-129">hostVersion: String</span></span>

<span data-ttu-id="9032b-130">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="9032b-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="9032b-131">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou `hostVersion` Ios, a propriedade retornará a versão do aplicativo host, Outlook.</span><span class="sxs-lookup"><span data-stu-id="9032b-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="9032b-132">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="9032b-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="9032b-133">Um exemplo é a cadeia de caracteres "15.0.468.0".</span><span class="sxs-lookup"><span data-stu-id="9032b-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="9032b-134">Tipo</span><span class="sxs-lookup"><span data-stu-id="9032b-134">Type</span></span>

*   <span data-ttu-id="9032b-135">String</span><span class="sxs-lookup"><span data-stu-id="9032b-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9032b-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9032b-136">Requirements</span></span>

|<span data-ttu-id="9032b-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="9032b-137">Requirement</span></span>| <span data-ttu-id="9032b-138">Valor</span><span class="sxs-lookup"><span data-stu-id="9032b-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="9032b-139">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9032b-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9032b-140">1.0</span><span class="sxs-lookup"><span data-stu-id="9032b-140">1.0</span></span>|
|[<span data-ttu-id="9032b-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9032b-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9032b-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9032b-142">ReadItem</span></span>|
|[<span data-ttu-id="9032b-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9032b-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9032b-144">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9032b-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="9032b-145">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9032b-145">OWAView: String</span></span>

<span data-ttu-id="9032b-146">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="9032b-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="9032b-147">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="9032b-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="9032b-148">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="9032b-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="9032b-149">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="9032b-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="9032b-150">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="9032b-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="9032b-151">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="9032b-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="9032b-152">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="9032b-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="9032b-153">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="9032b-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="9032b-154">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="9032b-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="9032b-155">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="9032b-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="9032b-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="9032b-156">Type</span></span>

*   <span data-ttu-id="9032b-157">String</span><span class="sxs-lookup"><span data-stu-id="9032b-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9032b-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9032b-158">Requirements</span></span>

|<span data-ttu-id="9032b-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="9032b-159">Requirement</span></span>| <span data-ttu-id="9032b-160">Valor</span><span class="sxs-lookup"><span data-stu-id="9032b-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="9032b-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9032b-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9032b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="9032b-162">1.0</span></span>|
|[<span data-ttu-id="9032b-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9032b-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9032b-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9032b-164">ReadItem</span></span>|
|[<span data-ttu-id="9032b-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9032b-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9032b-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9032b-166">Compose or Read</span></span>|
