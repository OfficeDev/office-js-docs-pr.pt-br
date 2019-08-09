---
title: 'Office.context.mailbox.diagnostics: conjunto de requisitos da visualização'
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0d0283edc0b9c331afb15a9b485c8f1447d5e8e5
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268527"
---
# <a name="diagnostics"></a><span data-ttu-id="084dd-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="084dd-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="084dd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="084dd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="084dd-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="084dd-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="084dd-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="084dd-105">Requirements</span></span>

|<span data-ttu-id="084dd-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="084dd-106">Requirement</span></span>| <span data-ttu-id="084dd-107">Valor</span><span class="sxs-lookup"><span data-stu-id="084dd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="084dd-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="084dd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="084dd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="084dd-109">1.0</span></span>|
|[<span data-ttu-id="084dd-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="084dd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="084dd-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="084dd-111">ReadItem</span></span>|
|[<span data-ttu-id="084dd-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="084dd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="084dd-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="084dd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="084dd-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="084dd-114">Members and methods</span></span>

| <span data-ttu-id="084dd-115">Membro</span><span class="sxs-lookup"><span data-stu-id="084dd-115">Member</span></span> | <span data-ttu-id="084dd-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="084dd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="084dd-117">hostName</span><span class="sxs-lookup"><span data-stu-id="084dd-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="084dd-118">Membro</span><span class="sxs-lookup"><span data-stu-id="084dd-118">Member</span></span> |
| [<span data-ttu-id="084dd-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="084dd-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="084dd-120">Membro</span><span class="sxs-lookup"><span data-stu-id="084dd-120">Member</span></span> |
| [<span data-ttu-id="084dd-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="084dd-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="084dd-122">Membro</span><span class="sxs-lookup"><span data-stu-id="084dd-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="084dd-123">Membros</span><span class="sxs-lookup"><span data-stu-id="084dd-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="084dd-124">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="084dd-124">hostName: String</span></span>

<span data-ttu-id="084dd-125">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="084dd-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="084dd-126">Uma cadeia de caracteres que pode ser um dos valores a seguir: `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="084dd-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="084dd-127">O `Outlook` valor é retornado para o Outlook em clientes de área de trabalho (ou seja, Windows e Mac).</span><span class="sxs-lookup"><span data-stu-id="084dd-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="084dd-128">Tipo</span><span class="sxs-lookup"><span data-stu-id="084dd-128">Type</span></span>

*   <span data-ttu-id="084dd-129">String</span><span class="sxs-lookup"><span data-stu-id="084dd-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="084dd-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="084dd-130">Requirements</span></span>

|<span data-ttu-id="084dd-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="084dd-131">Requirement</span></span>| <span data-ttu-id="084dd-132">Valor</span><span class="sxs-lookup"><span data-stu-id="084dd-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="084dd-133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="084dd-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="084dd-134">1.0</span><span class="sxs-lookup"><span data-stu-id="084dd-134">1.0</span></span>|
|[<span data-ttu-id="084dd-135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="084dd-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="084dd-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="084dd-136">ReadItem</span></span>|
|[<span data-ttu-id="084dd-137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="084dd-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="084dd-138">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="084dd-138">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="084dd-139">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="084dd-139">hostVersion: String</span></span>

<span data-ttu-id="084dd-140">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do servidor Exchange (por exemplo, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="084dd-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="084dd-141">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou no `hostVersion` Ios, a propriedade retornará a versão do aplicativo host do Outlook.</span><span class="sxs-lookup"><span data-stu-id="084dd-141">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="084dd-142">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="084dd-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="084dd-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="084dd-143">Type</span></span>

*   <span data-ttu-id="084dd-144">String</span><span class="sxs-lookup"><span data-stu-id="084dd-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="084dd-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="084dd-145">Requirements</span></span>

|<span data-ttu-id="084dd-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="084dd-146">Requirement</span></span>| <span data-ttu-id="084dd-147">Valor</span><span class="sxs-lookup"><span data-stu-id="084dd-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="084dd-148">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="084dd-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="084dd-149">1.0</span><span class="sxs-lookup"><span data-stu-id="084dd-149">1.0</span></span>|
|[<span data-ttu-id="084dd-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="084dd-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="084dd-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="084dd-151">ReadItem</span></span>|
|[<span data-ttu-id="084dd-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="084dd-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="084dd-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="084dd-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="084dd-154">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="084dd-154">OWAView: String</span></span>

<span data-ttu-id="084dd-155">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="084dd-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="084dd-156">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="084dd-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="084dd-157">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="084dd-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="084dd-158">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="084dd-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="084dd-159">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="084dd-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="084dd-160">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="084dd-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="084dd-161">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="084dd-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="084dd-162">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="084dd-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="084dd-163">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="084dd-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="084dd-164">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="084dd-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="084dd-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="084dd-165">Type</span></span>

*   <span data-ttu-id="084dd-166">String</span><span class="sxs-lookup"><span data-stu-id="084dd-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="084dd-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="084dd-167">Requirements</span></span>

|<span data-ttu-id="084dd-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="084dd-168">Requirement</span></span>| <span data-ttu-id="084dd-169">Valor</span><span class="sxs-lookup"><span data-stu-id="084dd-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="084dd-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="084dd-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="084dd-171">1.0</span><span class="sxs-lookup"><span data-stu-id="084dd-171">1.0</span></span>|
|[<span data-ttu-id="084dd-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="084dd-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="084dd-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="084dd-173">ReadItem</span></span>|
|[<span data-ttu-id="084dd-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="084dd-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="084dd-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="084dd-175">Compose or Read</span></span>|
