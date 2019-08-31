---
title: 'Office.context.mailbox.diagnostics: conjunto de requisitos da visualização'
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 47949fb26629b6619514bbd0d7f760cfa31d5839
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696488"
---
# <a name="diagnostics"></a><span data-ttu-id="eade7-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="eade7-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="eade7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="eade7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="eade7-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="eade7-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="eade7-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eade7-105">Requirements</span></span>

|<span data-ttu-id="eade7-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="eade7-106">Requirement</span></span>| <span data-ttu-id="eade7-107">Valor</span><span class="sxs-lookup"><span data-stu-id="eade7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="eade7-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eade7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eade7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="eade7-109">1.0</span></span>|
|[<span data-ttu-id="eade7-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="eade7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eade7-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eade7-111">ReadItem</span></span>|
|[<span data-ttu-id="eade7-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eade7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eade7-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eade7-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="eade7-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="eade7-114">Members and methods</span></span>

| <span data-ttu-id="eade7-115">Membro</span><span class="sxs-lookup"><span data-stu-id="eade7-115">Member</span></span> | <span data-ttu-id="eade7-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="eade7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="eade7-117">hostName</span><span class="sxs-lookup"><span data-stu-id="eade7-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="eade7-118">Membro</span><span class="sxs-lookup"><span data-stu-id="eade7-118">Member</span></span> |
| [<span data-ttu-id="eade7-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="eade7-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="eade7-120">Membro</span><span class="sxs-lookup"><span data-stu-id="eade7-120">Member</span></span> |
| [<span data-ttu-id="eade7-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="eade7-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="eade7-122">Membro</span><span class="sxs-lookup"><span data-stu-id="eade7-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="eade7-123">Membros</span><span class="sxs-lookup"><span data-stu-id="eade7-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="eade7-124">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eade7-124">hostName: String</span></span>

<span data-ttu-id="eade7-125">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="eade7-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="eade7-126">Uma cadeia de caracteres que pode ser um dos valores a seguir: `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="eade7-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="eade7-127">O `Outlook` valor é retornado para o Outlook em clientes de área de trabalho (ou seja, Windows e Mac).</span><span class="sxs-lookup"><span data-stu-id="eade7-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="eade7-128">Tipo</span><span class="sxs-lookup"><span data-stu-id="eade7-128">Type</span></span>

*   <span data-ttu-id="eade7-129">String</span><span class="sxs-lookup"><span data-stu-id="eade7-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eade7-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eade7-130">Requirements</span></span>

|<span data-ttu-id="eade7-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="eade7-131">Requirement</span></span>| <span data-ttu-id="eade7-132">Valor</span><span class="sxs-lookup"><span data-stu-id="eade7-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="eade7-133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eade7-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eade7-134">1.0</span><span class="sxs-lookup"><span data-stu-id="eade7-134">1.0</span></span>|
|[<span data-ttu-id="eade7-135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="eade7-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eade7-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eade7-136">ReadItem</span></span>|
|[<span data-ttu-id="eade7-137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eade7-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eade7-138">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eade7-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="eade7-139">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eade7-139">hostVersion: String</span></span>

<span data-ttu-id="eade7-140">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do servidor Exchange (por exemplo, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="eade7-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="eade7-141">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou no `hostVersion` Ios, a propriedade retornará a versão do aplicativo host do Outlook.</span><span class="sxs-lookup"><span data-stu-id="eade7-141">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="eade7-142">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="eade7-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="eade7-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="eade7-143">Type</span></span>

*   <span data-ttu-id="eade7-144">String</span><span class="sxs-lookup"><span data-stu-id="eade7-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eade7-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eade7-145">Requirements</span></span>

|<span data-ttu-id="eade7-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="eade7-146">Requirement</span></span>| <span data-ttu-id="eade7-147">Valor</span><span class="sxs-lookup"><span data-stu-id="eade7-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="eade7-148">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eade7-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eade7-149">1.0</span><span class="sxs-lookup"><span data-stu-id="eade7-149">1.0</span></span>|
|[<span data-ttu-id="eade7-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="eade7-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eade7-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eade7-151">ReadItem</span></span>|
|[<span data-ttu-id="eade7-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eade7-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eade7-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eade7-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="eade7-154">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eade7-154">OWAView: String</span></span>

<span data-ttu-id="eade7-155">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="eade7-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="eade7-156">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="eade7-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="eade7-157">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="eade7-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="eade7-158">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="eade7-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="eade7-159">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="eade7-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="eade7-160">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="eade7-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="eade7-161">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="eade7-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="eade7-162">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="eade7-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="eade7-163">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="eade7-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="eade7-164">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="eade7-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="eade7-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="eade7-165">Type</span></span>

*   <span data-ttu-id="eade7-166">String</span><span class="sxs-lookup"><span data-stu-id="eade7-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eade7-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eade7-167">Requirements</span></span>

|<span data-ttu-id="eade7-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="eade7-168">Requirement</span></span>| <span data-ttu-id="eade7-169">Valor</span><span class="sxs-lookup"><span data-stu-id="eade7-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="eade7-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eade7-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eade7-171">1.0</span><span class="sxs-lookup"><span data-stu-id="eade7-171">1.0</span></span>|
|[<span data-ttu-id="eade7-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="eade7-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eade7-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eade7-173">ReadItem</span></span>|
|[<span data-ttu-id="eade7-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eade7-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eade7-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eade7-175">Compose or Read</span></span>|
