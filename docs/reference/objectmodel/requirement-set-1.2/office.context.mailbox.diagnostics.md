---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2b57a26233ce76ed35ce9428cc12a4ba93b5ace8
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871497"
---
# <a name="diagnostics"></a><span data-ttu-id="624bd-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="624bd-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="624bd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="624bd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="624bd-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="624bd-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="624bd-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="624bd-105">Requirements</span></span>

|<span data-ttu-id="624bd-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="624bd-106">Requirement</span></span>| <span data-ttu-id="624bd-107">Valor</span><span class="sxs-lookup"><span data-stu-id="624bd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="624bd-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="624bd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="624bd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="624bd-109">1.0</span></span>|
|[<span data-ttu-id="624bd-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="624bd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="624bd-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="624bd-111">ReadItem</span></span>|
|[<span data-ttu-id="624bd-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="624bd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="624bd-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="624bd-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="624bd-114">Membros</span><span class="sxs-lookup"><span data-stu-id="624bd-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="624bd-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="624bd-115">hostName :String</span></span>

<span data-ttu-id="624bd-116">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="624bd-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="624bd-117">Uma cadeia de caracteres que pode ser um dos seguintes valores `Outlook`: `OutlookIOS`, ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="624bd-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="624bd-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="624bd-118">Type</span></span>

*   <span data-ttu-id="624bd-119">String</span><span class="sxs-lookup"><span data-stu-id="624bd-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="624bd-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="624bd-120">Requirements</span></span>

|<span data-ttu-id="624bd-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="624bd-121">Requirement</span></span>| <span data-ttu-id="624bd-122">Valor</span><span class="sxs-lookup"><span data-stu-id="624bd-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="624bd-123">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="624bd-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="624bd-124">1.0</span><span class="sxs-lookup"><span data-stu-id="624bd-124">1.0</span></span>|
|[<span data-ttu-id="624bd-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="624bd-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="624bd-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="624bd-126">ReadItem</span></span>|
|[<span data-ttu-id="624bd-127">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="624bd-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="624bd-128">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="624bd-128">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="624bd-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="624bd-129">hostVersion :String</span></span>

<span data-ttu-id="624bd-130">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="624bd-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="624bd-p101">Se o suplemento de email estiver em execução no cliente do Outlook para área de trabalho ou Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="624bd-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="624bd-134">Tipo</span><span class="sxs-lookup"><span data-stu-id="624bd-134">Type</span></span>

*   <span data-ttu-id="624bd-135">String</span><span class="sxs-lookup"><span data-stu-id="624bd-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="624bd-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="624bd-136">Requirements</span></span>

|<span data-ttu-id="624bd-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="624bd-137">Requirement</span></span>| <span data-ttu-id="624bd-138">Valor</span><span class="sxs-lookup"><span data-stu-id="624bd-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="624bd-139">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="624bd-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="624bd-140">1.0</span><span class="sxs-lookup"><span data-stu-id="624bd-140">1.0</span></span>|
|[<span data-ttu-id="624bd-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="624bd-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="624bd-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="624bd-142">ReadItem</span></span>|
|[<span data-ttu-id="624bd-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="624bd-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="624bd-144">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="624bd-144">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="624bd-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="624bd-145">OWAView :String</span></span>

<span data-ttu-id="624bd-146">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="624bd-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="624bd-147">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="624bd-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="624bd-148">Se o aplicativo host não for Outlook Web App, acessar essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="624bd-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="624bd-149">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e à quantidade de colunas que pode ser exibida:</span><span class="sxs-lookup"><span data-stu-id="624bd-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="624bd-p102">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="624bd-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="624bd-p103">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="624bd-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="624bd-p104">`ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook Web App usa esse modo de exibição em um modo de tela cheia em um computador de mesa.</span><span class="sxs-lookup"><span data-stu-id="624bd-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="624bd-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="624bd-156">Type</span></span>

*   <span data-ttu-id="624bd-157">String</span><span class="sxs-lookup"><span data-stu-id="624bd-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="624bd-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="624bd-158">Requirements</span></span>

|<span data-ttu-id="624bd-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="624bd-159">Requirement</span></span>| <span data-ttu-id="624bd-160">Valor</span><span class="sxs-lookup"><span data-stu-id="624bd-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="624bd-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="624bd-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="624bd-162">1.0</span><span class="sxs-lookup"><span data-stu-id="624bd-162">1.0</span></span>|
|[<span data-ttu-id="624bd-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="624bd-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="624bd-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="624bd-164">ReadItem</span></span>|
|[<span data-ttu-id="624bd-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="624bd-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="624bd-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="624bd-166">Compose or Read</span></span>|
