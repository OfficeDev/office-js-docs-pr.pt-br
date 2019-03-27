---
title: 'Office.context.mailbox.diagnostics: conjunto de requisitos da visualização'
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3bcc1004ddcee638965c15c7f320e2320a6faf06
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871896"
---
# <a name="diagnostics"></a><span data-ttu-id="98575-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="98575-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="98575-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="98575-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="98575-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="98575-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="98575-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98575-105">Requirements</span></span>

|<span data-ttu-id="98575-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="98575-106">Requirement</span></span>| <span data-ttu-id="98575-107">Valor</span><span class="sxs-lookup"><span data-stu-id="98575-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="98575-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98575-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98575-109">1.0</span><span class="sxs-lookup"><span data-stu-id="98575-109">1.0</span></span>|
|[<span data-ttu-id="98575-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98575-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98575-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98575-111">ReadItem</span></span>|
|[<span data-ttu-id="98575-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98575-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98575-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98575-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="98575-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="98575-114">Members and methods</span></span>

| <span data-ttu-id="98575-115">Membro</span><span class="sxs-lookup"><span data-stu-id="98575-115">Member</span></span> | <span data-ttu-id="98575-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="98575-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="98575-117">hostName</span><span class="sxs-lookup"><span data-stu-id="98575-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="98575-118">Member</span><span class="sxs-lookup"><span data-stu-id="98575-118">Member</span></span> |
| [<span data-ttu-id="98575-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="98575-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="98575-120">Member</span><span class="sxs-lookup"><span data-stu-id="98575-120">Member</span></span> |
| [<span data-ttu-id="98575-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="98575-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="98575-122">Membro</span><span class="sxs-lookup"><span data-stu-id="98575-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="98575-123">Membros</span><span class="sxs-lookup"><span data-stu-id="98575-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="98575-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="98575-124">hostName :String</span></span>

<span data-ttu-id="98575-125">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="98575-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="98575-126">Uma cadeia de caracteres que pode ser um dos valores a seguir: `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="98575-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="98575-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="98575-127">Type</span></span>

*   <span data-ttu-id="98575-128">String</span><span class="sxs-lookup"><span data-stu-id="98575-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98575-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98575-129">Requirements</span></span>

|<span data-ttu-id="98575-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="98575-130">Requirement</span></span>| <span data-ttu-id="98575-131">Valor</span><span class="sxs-lookup"><span data-stu-id="98575-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="98575-132">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98575-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98575-133">1.0</span><span class="sxs-lookup"><span data-stu-id="98575-133">1.0</span></span>|
|[<span data-ttu-id="98575-134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98575-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98575-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98575-135">ReadItem</span></span>|
|[<span data-ttu-id="98575-136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98575-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98575-137">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98575-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="98575-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="98575-138">hostVersion :String</span></span>

<span data-ttu-id="98575-139">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="98575-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="98575-p101">Se o suplemento de email estiver em execução no cliente do Outlook para área de trabalho ou Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="98575-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="98575-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="98575-143">Type</span></span>

*   <span data-ttu-id="98575-144">String</span><span class="sxs-lookup"><span data-stu-id="98575-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98575-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98575-145">Requirements</span></span>

|<span data-ttu-id="98575-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="98575-146">Requirement</span></span>| <span data-ttu-id="98575-147">Valor</span><span class="sxs-lookup"><span data-stu-id="98575-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="98575-148">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98575-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98575-149">1.0</span><span class="sxs-lookup"><span data-stu-id="98575-149">1.0</span></span>|
|[<span data-ttu-id="98575-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98575-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98575-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98575-151">ReadItem</span></span>|
|[<span data-ttu-id="98575-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98575-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98575-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98575-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="98575-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="98575-154">OWAView :String</span></span>

<span data-ttu-id="98575-155">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="98575-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="98575-156">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="98575-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="98575-157">Se o aplicativo host não for Outlook Web App, acessar essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="98575-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="98575-158">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e à quantidade de colunas que pode ser exibida:</span><span class="sxs-lookup"><span data-stu-id="98575-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="98575-p102">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="98575-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="98575-p103">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="98575-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="98575-p104">`ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook Web App usa esse modo de exibição em um modo de tela cheia em um computador de mesa.</span><span class="sxs-lookup"><span data-stu-id="98575-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="98575-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="98575-165">Type</span></span>

*   <span data-ttu-id="98575-166">String</span><span class="sxs-lookup"><span data-stu-id="98575-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98575-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98575-167">Requirements</span></span>

|<span data-ttu-id="98575-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="98575-168">Requirement</span></span>| <span data-ttu-id="98575-169">Valor</span><span class="sxs-lookup"><span data-stu-id="98575-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="98575-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98575-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98575-171">1.0</span><span class="sxs-lookup"><span data-stu-id="98575-171">1.0</span></span>|
|[<span data-ttu-id="98575-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98575-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98575-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98575-173">ReadItem</span></span>|
|[<span data-ttu-id="98575-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98575-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98575-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98575-175">Compose or Read</span></span>|
