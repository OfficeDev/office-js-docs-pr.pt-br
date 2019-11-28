---
title: 'Office.context.mailbox.diagnostics: conjunto de requisitos da visualização'
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629199"
---
# <a name="diagnostics"></a><span data-ttu-id="e2d0e-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="e2d0e-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="e2d0e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="e2d0e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="e2d0e-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d0e-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-105">Requirements</span></span>

|<span data-ttu-id="e2d0e-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="e2d0e-106">Requirement</span></span>| <span data-ttu-id="e2d0e-107">Valor</span><span class="sxs-lookup"><span data-stu-id="e2d0e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d0e-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e2d0e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2d0e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-109">1.0</span></span>|
|[<span data-ttu-id="e2d0e-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2d0e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-111">ReadItem</span></span>|
|[<span data-ttu-id="e2d0e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e2d0e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2d0e-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e2d0e-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="e2d0e-114">Properties</span></span>

| <span data-ttu-id="e2d0e-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="e2d0e-115">Property</span></span> | <span data-ttu-id="e2d0e-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-116">Minimum</span></span><br><span data-ttu-id="e2d0e-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="e2d0e-117">permission level</span></span> | <span data-ttu-id="e2d0e-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-118">Modes</span></span> | <span data-ttu-id="e2d0e-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="e2d0e-119">Return type</span></span> | <span data-ttu-id="e2d0e-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-120">Minimum</span></span><br><span data-ttu-id="e2d0e-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="e2d0e-122">hostName</span><span class="sxs-lookup"><span data-stu-id="e2d0e-122">hostName</span></span>](#hostname-string) | <span data-ttu-id="e2d0e-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-123">ReadItem</span></span> | <span data-ttu-id="e2d0e-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="e2d0e-124">Compose</span></span><br><span data-ttu-id="e2d0e-125">Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-125">Read</span></span> | <span data-ttu-id="e2d0e-126">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-126">String</span></span> | <span data-ttu-id="e2d0e-127">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-127">1.0</span></span> |
| [<span data-ttu-id="e2d0e-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="e2d0e-128">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="e2d0e-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-129">ReadItem</span></span> | <span data-ttu-id="e2d0e-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="e2d0e-130">Compose</span></span><br><span data-ttu-id="e2d0e-131">Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-131">Read</span></span> | <span data-ttu-id="e2d0e-132">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-132">String</span></span> | <span data-ttu-id="e2d0e-133">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-133">1.0</span></span> |
| [<span data-ttu-id="e2d0e-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="e2d0e-134">OWAView</span></span>](#owaview-string) | <span data-ttu-id="e2d0e-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-135">ReadItem</span></span> | <span data-ttu-id="e2d0e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="e2d0e-136">Compose</span></span><br><span data-ttu-id="e2d0e-137">Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-137">Read</span></span> | <span data-ttu-id="e2d0e-138">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-138">String</span></span> | <span data-ttu-id="e2d0e-139">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-139">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="e2d0e-140">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="e2d0e-140">Property details</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="e2d0e-141">Nome do host: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e2d0e-141">hostName: String</span></span>

<span data-ttu-id="e2d0e-142">Obtém uma cadeia de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-142">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="e2d0e-143">Uma cadeia de caracteres que pode ser um dos valores a seguir: `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-143">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d0e-144">O `Outlook` valor é retornado para o Outlook em clientes de área de trabalho (ou seja, Windows e Mac).</span><span class="sxs-lookup"><span data-stu-id="e2d0e-144">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="e2d0e-145">Tipo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-145">Type</span></span>

*   <span data-ttu-id="e2d0e-146">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-146">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d0e-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-147">Requirements</span></span>

|<span data-ttu-id="e2d0e-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="e2d0e-148">Requirement</span></span>| <span data-ttu-id="e2d0e-149">Valor</span><span class="sxs-lookup"><span data-stu-id="e2d0e-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d0e-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e2d0e-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2d0e-151">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-151">1.0</span></span>|
|[<span data-ttu-id="e2d0e-152">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2d0e-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-153">ReadItem</span></span>|
|[<span data-ttu-id="e2d0e-154">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e2d0e-154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2d0e-155">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-155">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="e2d0e-156">hostVersion: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e2d0e-156">hostVersion: String</span></span>

<span data-ttu-id="e2d0e-157">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do servidor Exchange (por exemplo, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="e2d0e-157">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="e2d0e-158">Se o suplemento de email estiver em execução em uma área de trabalho do Outlook ou cliente `hostVersion` móvel, a propriedade retornará a versão do aplicativo host, Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-158">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="e2d0e-159">No Outlook na Web, a propriedade retorna a versão do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-159">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d0e-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-160">Type</span></span>

*   <span data-ttu-id="e2d0e-161">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-161">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d0e-162">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-162">Requirements</span></span>

|<span data-ttu-id="e2d0e-163">Requisito</span><span class="sxs-lookup"><span data-stu-id="e2d0e-163">Requirement</span></span>| <span data-ttu-id="e2d0e-164">Valor</span><span class="sxs-lookup"><span data-stu-id="e2d0e-164">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d0e-165">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e2d0e-165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2d0e-166">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-166">1.0</span></span>|
|[<span data-ttu-id="e2d0e-167">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-167">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2d0e-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-168">ReadItem</span></span>|
|[<span data-ttu-id="e2d0e-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e2d0e-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2d0e-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="e2d0e-171">OWAView: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e2d0e-171">OWAView: String</span></span>

<span data-ttu-id="e2d0e-172">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-172">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="e2d0e-173">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-173">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="e2d0e-174">Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-174">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="e2d0e-175">O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:</span><span class="sxs-lookup"><span data-stu-id="e2d0e-175">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="e2d0e-176">`OneColumn`, que é exibido quando a tela é estreita.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-176">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="e2d0e-177">O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-177">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="e2d0e-178">`TwoColumns`, que é exibido quando a tela é mais larga.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-178">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="e2d0e-179">O Outlook na Web usa esse modo de exibição na maioria dos Tablets.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-179">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="e2d0e-180">`ThreeColumns`, que é exibido quando a tela é ainda mais larga.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-180">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="e2d0e-181">Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.</span><span class="sxs-lookup"><span data-stu-id="e2d0e-181">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d0e-182">Tipo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-182">Type</span></span>

*   <span data-ttu-id="e2d0e-183">String</span><span class="sxs-lookup"><span data-stu-id="e2d0e-183">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d0e-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2d0e-184">Requirements</span></span>

|<span data-ttu-id="e2d0e-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="e2d0e-185">Requirement</span></span>| <span data-ttu-id="e2d0e-186">Valor</span><span class="sxs-lookup"><span data-stu-id="e2d0e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d0e-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e2d0e-187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2d0e-188">1.0</span><span class="sxs-lookup"><span data-stu-id="e2d0e-188">1.0</span></span>|
|[<span data-ttu-id="e2d0e-189">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e2d0e-189">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2d0e-190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2d0e-190">ReadItem</span></span>|
|[<span data-ttu-id="e2d0e-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e2d0e-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2d0e-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e2d0e-192">Compose or Read</span></span>|
