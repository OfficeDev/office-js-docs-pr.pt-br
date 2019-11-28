---
title: Office. Context. Mailbox-visualização do conjunto de requisitos
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 8c67f7cf9231dd1c0db0d9a8d4ae9fb48e458435
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629192"
---
# <a name="mailbox"></a><span data-ttu-id="03f8a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="03f8a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="03f8a-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="03f8a-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03f8a-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-105">Requirements</span></span>

|<span data-ttu-id="03f8a-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-106">Requirement</span></span>| <span data-ttu-id="03f8a-107">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-109">1.0</span></span>|
|[<span data-ttu-id="03f8a-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="03f8a-111">Restricted</span></span>|
|[<span data-ttu-id="03f8a-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="03f8a-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="03f8a-114">Properties</span></span>

| <span data-ttu-id="03f8a-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="03f8a-115">Property</span></span> | <span data-ttu-id="03f8a-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-116">Minimum</span></span><br><span data-ttu-id="03f8a-117">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="03f8a-117">permission level</span></span> | <span data-ttu-id="03f8a-118">Modelos</span><span class="sxs-lookup"><span data-stu-id="03f8a-118">Modes</span></span> | <span data-ttu-id="03f8a-119">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="03f8a-119">Return type</span></span> | <span data-ttu-id="03f8a-120">Mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-120">Minimum</span></span><br><span data-ttu-id="03f8a-121">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="03f8a-122">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="03f8a-122">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="03f8a-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-123">ReadItem</span></span> | <span data-ttu-id="03f8a-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-124">Compose</span></span><br><span data-ttu-id="03f8a-125">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-125">Read</span></span> | <span data-ttu-id="03f8a-126">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-126">String</span></span> | <span data-ttu-id="03f8a-127">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-127">1.0</span></span> |
| [<span data-ttu-id="03f8a-128">Nova mastercategories</span><span class="sxs-lookup"><span data-stu-id="03f8a-128">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="03f8a-129">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-129">ReadWriteMailbox</span></span> | <span data-ttu-id="03f8a-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-130">Compose</span></span><br><span data-ttu-id="03f8a-131">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-131">Read</span></span> | [<span data-ttu-id="03f8a-132">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="03f8a-132">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories) | <span data-ttu-id="03f8a-133">Visualização</span><span class="sxs-lookup"><span data-stu-id="03f8a-133">Preview</span></span> |
| [<span data-ttu-id="03f8a-134">restUrl</span><span class="sxs-lookup"><span data-stu-id="03f8a-134">restUrl</span></span>](#resturl-string) | <span data-ttu-id="03f8a-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-135">ReadItem</span></span> | <span data-ttu-id="03f8a-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-136">Compose</span></span><br><span data-ttu-id="03f8a-137">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-137">Read</span></span> | <span data-ttu-id="03f8a-138">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-138">String</span></span> | <span data-ttu-id="03f8a-139">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-139">1.5</span></span> |

##### <a name="methods"></a><span data-ttu-id="03f8a-140">Métodos</span><span class="sxs-lookup"><span data-stu-id="03f8a-140">Methods</span></span>

| <span data-ttu-id="03f8a-141">Método</span><span class="sxs-lookup"><span data-stu-id="03f8a-141">Method</span></span> | <span data-ttu-id="03f8a-142">Mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-142">Minimum</span></span><br><span data-ttu-id="03f8a-143">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="03f8a-143">permission level</span></span> | <span data-ttu-id="03f8a-144">Modelos</span><span class="sxs-lookup"><span data-stu-id="03f8a-144">Modes</span></span> | <span data-ttu-id="03f8a-145">Mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-145">Minimum</span></span><br><span data-ttu-id="03f8a-146">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-146">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="03f8a-147">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-147">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="03f8a-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-148">ReadItem</span></span> | <span data-ttu-id="03f8a-149">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-149">Compose</span></span><br><span data-ttu-id="03f8a-150">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-150">Read</span></span> | <span data-ttu-id="03f8a-151">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-151">1.5</span></span> |
| [<span data-ttu-id="03f8a-152">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="03f8a-152">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="03f8a-153">Restrito</span><span class="sxs-lookup"><span data-stu-id="03f8a-153">Restricted</span></span> | <span data-ttu-id="03f8a-154">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-154">Compose</span></span><br><span data-ttu-id="03f8a-155">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-155">Read</span></span> | <span data-ttu-id="03f8a-156">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-156">1.3</span></span> |
| [<span data-ttu-id="03f8a-157">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="03f8a-157">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="03f8a-158">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-158">ReadItem</span></span> | <span data-ttu-id="03f8a-159">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-159">Compose</span></span><br><span data-ttu-id="03f8a-160">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-160">Read</span></span> | <span data-ttu-id="03f8a-161">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-161">1.0</span></span> |
| [<span data-ttu-id="03f8a-162">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="03f8a-162">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="03f8a-163">Restrito</span><span class="sxs-lookup"><span data-stu-id="03f8a-163">Restricted</span></span> | <span data-ttu-id="03f8a-164">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-164">Compose</span></span><br><span data-ttu-id="03f8a-165">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-165">Read</span></span> | <span data-ttu-id="03f8a-166">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-166">1.3</span></span> |
| [<span data-ttu-id="03f8a-167">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="03f8a-167">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="03f8a-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-168">ReadItem</span></span> | <span data-ttu-id="03f8a-169">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-169">Compose</span></span><br><span data-ttu-id="03f8a-170">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-170">Read</span></span> | <span data-ttu-id="03f8a-171">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-171">1.0</span></span> |
| [<span data-ttu-id="03f8a-172">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="03f8a-172">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="03f8a-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-173">ReadItem</span></span> | <span data-ttu-id="03f8a-174">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-174">Compose</span></span><br><span data-ttu-id="03f8a-175">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-175">Read</span></span> | <span data-ttu-id="03f8a-176">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-176">1.0</span></span> |
| [<span data-ttu-id="03f8a-177">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="03f8a-177">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="03f8a-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-178">ReadItem</span></span> | <span data-ttu-id="03f8a-179">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-179">Compose</span></span><br><span data-ttu-id="03f8a-180">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-180">Read</span></span> | <span data-ttu-id="03f8a-181">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-181">1.0</span></span> |
| [<span data-ttu-id="03f8a-182">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="03f8a-182">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="03f8a-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-183">ReadItem</span></span> | <span data-ttu-id="03f8a-184">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-184">Read</span></span> | <span data-ttu-id="03f8a-185">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-185">1.0</span></span> |
| [<span data-ttu-id="03f8a-186">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="03f8a-186">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="03f8a-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-187">ReadItem</span></span> | <span data-ttu-id="03f8a-188">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-188">Compose</span></span><br><span data-ttu-id="03f8a-189">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-189">Read</span></span> | <span data-ttu-id="03f8a-190">1.6</span><span class="sxs-lookup"><span data-stu-id="03f8a-190">1.6</span></span> |
| [<span data-ttu-id="03f8a-191">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-191">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="03f8a-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-192">ReadItem</span></span> | <span data-ttu-id="03f8a-193">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-193">Compose</span></span><br><span data-ttu-id="03f8a-194">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-194">Read</span></span> | <span data-ttu-id="03f8a-195">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-195">1.5</span></span> |
| [<span data-ttu-id="03f8a-196">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-196">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="03f8a-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-197">ReadItem</span></span> | <span data-ttu-id="03f8a-198">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-198">Compose</span></span><br><span data-ttu-id="03f8a-199">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-199">Read</span></span> | <span data-ttu-id="03f8a-200">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-200">1.3</span></span><br><span data-ttu-id="03f8a-201">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-201">1.0</span></span> |
| [<span data-ttu-id="03f8a-202">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-202">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="03f8a-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-203">ReadItem</span></span> | <span data-ttu-id="03f8a-204">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-204">Compose</span></span><br><span data-ttu-id="03f8a-205">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-205">Read</span></span> | <span data-ttu-id="03f8a-206">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-206">1.0</span></span> |
| [<span data-ttu-id="03f8a-207">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-207">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="03f8a-208">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-208">ReadWriteMailbox</span></span> | <span data-ttu-id="03f8a-209">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-209">Compose</span></span><br><span data-ttu-id="03f8a-210">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-210">Read</span></span> | <span data-ttu-id="03f8a-211">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-211">1.0</span></span> |
| [<span data-ttu-id="03f8a-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="03f8a-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="03f8a-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-213">ReadItem</span></span> | <span data-ttu-id="03f8a-214">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-214">Compose</span></span><br><span data-ttu-id="03f8a-215">Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-215">Read</span></span> | <span data-ttu-id="03f8a-216">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-216">1.5</span></span> |

##### <a name="events"></a><span data-ttu-id="03f8a-217">Eventos</span><span class="sxs-lookup"><span data-stu-id="03f8a-217">Events</span></span>

<span data-ttu-id="03f8a-218">Você pode assinar e cancelar a assinatura dos eventos a seguir usando o [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) e o [removeHandlerAsync](#removehandlerasynceventtype-options-callback) , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-218">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="03f8a-219">Evento</span><span class="sxs-lookup"><span data-stu-id="03f8a-219">Event</span></span> | <span data-ttu-id="03f8a-220">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-220">Description</span></span> | <span data-ttu-id="03f8a-221">Mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-221">Minimum</span></span><br><span data-ttu-id="03f8a-222">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-222">requirement set</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="03f8a-223">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="03f8a-223">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="03f8a-224">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-224">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="03f8a-225">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="03f8a-225">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="03f8a-226">Visualização</span><span class="sxs-lookup"><span data-stu-id="03f8a-226">Preview</span></span> |

### <a name="namespaces"></a><span data-ttu-id="03f8a-227">Namespaces</span><span class="sxs-lookup"><span data-stu-id="03f8a-227">Namespaces</span></span>

<span data-ttu-id="03f8a-228">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03f8a-228">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="03f8a-229">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03f8a-229">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="03f8a-230">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03f8a-230">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

## <a name="property-details"></a><span data-ttu-id="03f8a-231">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="03f8a-231">Property details</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="03f8a-232">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="03f8a-232">ewsUrl: String</span></span>

<span data-ttu-id="03f8a-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-235">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-235">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03f8a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="03f8a-238">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="03f8a-238">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="03f8a-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="03f8a-241">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-241">Type</span></span>

*   <span data-ttu-id="03f8a-242">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-242">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03f8a-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-243">Requirements</span></span>

|<span data-ttu-id="03f8a-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-244">Requirement</span></span>| <span data-ttu-id="03f8a-245">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-246">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-247">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-247">1.0</span></span>|
|[<span data-ttu-id="03f8a-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-248">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-249">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-250">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-250">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-251">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-251">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="03f8a-252">Nova mastercategories: [nova mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="03f8a-252">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="03f8a-253">Obtém um objeto que fornece métodos para gerenciar a lista mestra de categorias nesta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="03f8a-253">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-254">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-254">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="03f8a-255">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-255">Type</span></span>

*   [<span data-ttu-id="03f8a-256">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="03f8a-256">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="03f8a-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-257">Requirements</span></span>

|<span data-ttu-id="03f8a-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-258">Requirement</span></span>| <span data-ttu-id="03f8a-259">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-261">1,8</span><span class="sxs-lookup"><span data-stu-id="03f8a-261">1.8</span></span> |
|[<span data-ttu-id="03f8a-262">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-262">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-263">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-263">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="03f8a-264">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-264">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-265">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-265">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="03f8a-266">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-266">Example</span></span>

<span data-ttu-id="03f8a-267">Este exemplo obtém a lista mestra de categorias para esta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="03f8a-267">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="03f8a-268">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="03f8a-268">restUrl: String</span></span>

<span data-ttu-id="03f8a-269">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="03f8a-269">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="03f8a-270">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03f8a-270">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="03f8a-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-271">Type</span></span>

*   <span data-ttu-id="03f8a-272">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03f8a-273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-273">Requirements</span></span>

|<span data-ttu-id="03f8a-274">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-274">Requirement</span></span>| <span data-ttu-id="03f8a-275">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-277">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-277">1.5</span></span> |
|[<span data-ttu-id="03f8a-278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-279">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-281">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-281">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="03f8a-282">Detalhes do método</span><span class="sxs-lookup"><span data-stu-id="03f8a-282">Method details</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="03f8a-283">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03f8a-283">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="03f8a-284">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="03f8a-284">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="03f8a-285">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="03f8a-285">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-286">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-286">Parameters</span></span>

| <span data-ttu-id="03f8a-287">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-287">Name</span></span> | <span data-ttu-id="03f8a-288">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-288">Type</span></span> | <span data-ttu-id="03f8a-289">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-289">Attributes</span></span> | <span data-ttu-id="03f8a-290">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-290">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="03f8a-291">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="03f8a-291">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="03f8a-292">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="03f8a-292">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="03f8a-293">Função</span><span class="sxs-lookup"><span data-stu-id="03f8a-293">Function</span></span> || <span data-ttu-id="03f8a-p104">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="03f8a-297">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-297">Object</span></span> | <span data-ttu-id="03f8a-298">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-298">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-299">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03f8a-299">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03f8a-300">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-300">Object</span></span> | <span data-ttu-id="03f8a-301">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-301">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-302">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="03f8a-302">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="03f8a-303">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-303">function</span></span>| <span data-ttu-id="03f8a-304">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-304">&lt;optional&gt;</span></span>|<span data-ttu-id="03f8a-305">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-305">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-306">Requirements</span></span>

|<span data-ttu-id="03f8a-307">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-307">Requirement</span></span>| <span data-ttu-id="03f8a-308">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-309">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-310">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-310">1.5</span></span> |
|[<span data-ttu-id="03f8a-311">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-312">ReadItem</span></span> |
|[<span data-ttu-id="03f8a-313">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-314">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-315">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-315">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="03f8a-316">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="03f8a-316">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="03f8a-317">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="03f8a-317">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-318">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-318">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-p105">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-321">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-321">Parameters</span></span>

|<span data-ttu-id="03f8a-322">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-322">Name</span></span>| <span data-ttu-id="03f8a-323">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-323">Type</span></span>| <span data-ttu-id="03f8a-324">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-324">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03f8a-325">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-325">String</span></span>|<span data-ttu-id="03f8a-326">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-326">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="03f8a-327">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="03f8a-327">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="03f8a-328">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-328">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-329">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-329">Requirements</span></span>

|<span data-ttu-id="03f8a-330">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-330">Requirement</span></span>| <span data-ttu-id="03f8a-331">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-331">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-332">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-332">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-333">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-333">1.3</span></span>|
|[<span data-ttu-id="03f8a-334">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-334">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-335">Restrito</span><span class="sxs-lookup"><span data-stu-id="03f8a-335">Restricted</span></span>|
|[<span data-ttu-id="03f8a-336">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-336">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-337">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-337">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03f8a-338">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03f8a-338">Returns:</span></span>

<span data-ttu-id="03f8a-339">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="03f8a-339">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="03f8a-340">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-340">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="03f8a-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="03f8a-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="03f8a-342">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-342">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="03f8a-p106">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="03f8a-p107">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-348">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-348">Parameters</span></span>

|<span data-ttu-id="03f8a-349">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-349">Name</span></span>| <span data-ttu-id="03f8a-350">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-350">Type</span></span>| <span data-ttu-id="03f8a-351">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-351">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="03f8a-352">Date</span><span class="sxs-lookup"><span data-stu-id="03f8a-352">Date</span></span>|<span data-ttu-id="03f8a-353">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="03f8a-353">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-354">Requirements</span></span>

|<span data-ttu-id="03f8a-355">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-355">Requirement</span></span>| <span data-ttu-id="03f8a-356">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-357">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-358">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-358">1.0</span></span>|
|[<span data-ttu-id="03f8a-359">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-360">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-361">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-362">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-362">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03f8a-363">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03f8a-363">Returns:</span></span>

<span data-ttu-id="03f8a-364">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="03f8a-364">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="03f8a-365">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="03f8a-365">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="03f8a-366">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="03f8a-366">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-367">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-p108">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-370">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-370">Parameters</span></span>

|<span data-ttu-id="03f8a-371">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-371">Name</span></span>| <span data-ttu-id="03f8a-372">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-372">Type</span></span>| <span data-ttu-id="03f8a-373">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-373">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03f8a-374">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-374">String</span></span>|<span data-ttu-id="03f8a-375">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="03f8a-375">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="03f8a-376">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="03f8a-376">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="03f8a-377">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="03f8a-377">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-378">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-378">Requirements</span></span>

|<span data-ttu-id="03f8a-379">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-379">Requirement</span></span>| <span data-ttu-id="03f8a-380">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-381">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-382">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-382">1.3</span></span>|
|[<span data-ttu-id="03f8a-383">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-384">Restrito</span><span class="sxs-lookup"><span data-stu-id="03f8a-384">Restricted</span></span>|
|[<span data-ttu-id="03f8a-385">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-386">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-386">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03f8a-387">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03f8a-387">Returns:</span></span>

<span data-ttu-id="03f8a-388">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="03f8a-388">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="03f8a-389">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-389">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="03f8a-390">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="03f8a-390">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="03f8a-391">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="03f8a-391">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="03f8a-392">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-392">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-393">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-393">Parameters</span></span>

|<span data-ttu-id="03f8a-394">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-394">Name</span></span>| <span data-ttu-id="03f8a-395">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-395">Type</span></span>| <span data-ttu-id="03f8a-396">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-396">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="03f8a-397">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="03f8a-397">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="03f8a-398">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="03f8a-398">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-399">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-399">Requirements</span></span>

|<span data-ttu-id="03f8a-400">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-400">Requirement</span></span>| <span data-ttu-id="03f8a-401">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-402">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-403">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-403">1.0</span></span>|
|[<span data-ttu-id="03f8a-404">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-405">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-406">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-407">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-407">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03f8a-408">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03f8a-408">Returns:</span></span>

<span data-ttu-id="03f8a-409">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="03f8a-409">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="03f8a-410">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="03f8a-410">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="03f8a-411">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-411">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="03f8a-412">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="03f8a-412">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="03f8a-413">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-413">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-414">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-414">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-415">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="03f8a-415">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="03f8a-p109">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="03f8a-418">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="03f8a-418">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="03f8a-419">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="03f8a-419">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-420">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-420">Parameters</span></span>

|<span data-ttu-id="03f8a-421">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-421">Name</span></span>| <span data-ttu-id="03f8a-422">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-422">Type</span></span>| <span data-ttu-id="03f8a-423">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-423">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03f8a-424">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-424">String</span></span>|<span data-ttu-id="03f8a-425">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-425">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-426">Requirements</span></span>

|<span data-ttu-id="03f8a-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-427">Requirement</span></span>| <span data-ttu-id="03f8a-428">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-430">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-430">1.0</span></span>|
|[<span data-ttu-id="03f8a-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-432">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-433">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-434">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-435">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="03f8a-436">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="03f8a-436">displayMessageForm(itemId)</span></span>

<span data-ttu-id="03f8a-437">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-437">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-438">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-438">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-439">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="03f8a-439">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="03f8a-440">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="03f8a-440">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="03f8a-441">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="03f8a-441">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="03f8a-p110">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-444">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-444">Parameters</span></span>

|<span data-ttu-id="03f8a-445">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-445">Name</span></span>| <span data-ttu-id="03f8a-446">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-446">Type</span></span>| <span data-ttu-id="03f8a-447">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-447">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03f8a-448">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-448">String</span></span>|<span data-ttu-id="03f8a-449">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-449">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-450">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-450">Requirements</span></span>

|<span data-ttu-id="03f8a-451">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-451">Requirement</span></span>| <span data-ttu-id="03f8a-452">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-453">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-454">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-454">1.0</span></span>|
|[<span data-ttu-id="03f8a-455">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-456">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-457">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-458">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-458">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-459">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-459">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="03f8a-460">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="03f8a-460">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="03f8a-461">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="03f8a-461">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-462">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03f8a-462">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03f8a-p111">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="03f8a-p112">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="03f8a-p113">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="03f8a-470">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="03f8a-470">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-471">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-471">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-472">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-472">All parameters are optional.</span></span>

|<span data-ttu-id="03f8a-473">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-473">Name</span></span>| <span data-ttu-id="03f8a-474">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-474">Type</span></span>| <span data-ttu-id="03f8a-475">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-475">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="03f8a-476">Object</span><span class="sxs-lookup"><span data-stu-id="03f8a-476">Object</span></span> | <span data-ttu-id="03f8a-477">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="03f8a-477">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="03f8a-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="03f8a-p114">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="03f8a-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="03f8a-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="03f8a-484">Data</span><span class="sxs-lookup"><span data-stu-id="03f8a-484">Date</span></span> | <span data-ttu-id="03f8a-485">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="03f8a-485">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="03f8a-486">Data</span><span class="sxs-lookup"><span data-stu-id="03f8a-486">Date</span></span> | <span data-ttu-id="03f8a-487">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="03f8a-487">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="03f8a-488">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-488">String</span></span> | <span data-ttu-id="03f8a-p116">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="03f8a-491">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-491">Array.&lt;String&gt;</span></span> | <span data-ttu-id="03f8a-p117">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="03f8a-494">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-494">String</span></span> | <span data-ttu-id="03f8a-p118">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="03f8a-497">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-497">String</span></span> | <span data-ttu-id="03f8a-p119">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03f8a-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-500">Requirements</span></span>

|<span data-ttu-id="03f8a-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-501">Requirement</span></span>| <span data-ttu-id="03f8a-502">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-504">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-504">1.0</span></span>|
|[<span data-ttu-id="03f8a-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-506">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-507">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-508">Read</span><span class="sxs-lookup"><span data-stu-id="03f8a-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-509">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-509">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="03f8a-510">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="03f8a-510">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="03f8a-511">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-511">Displays a form for creating a new message.</span></span>

<span data-ttu-id="03f8a-512">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-512">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="03f8a-513">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="03f8a-513">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="03f8a-514">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="03f8a-514">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-515">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-515">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-516">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-516">All parameters are optional.</span></span>

|<span data-ttu-id="03f8a-517">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-517">Name</span></span>| <span data-ttu-id="03f8a-518">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-518">Type</span></span>| <span data-ttu-id="03f8a-519">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-519">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="03f8a-520">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-520">Object</span></span> | <span data-ttu-id="03f8a-521">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-521">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="03f8a-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="03f8a-523">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="03f8a-523">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="03f8a-524">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-524">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="03f8a-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="03f8a-526">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="03f8a-526">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="03f8a-527">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-527">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="03f8a-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="03f8a-529">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="03f8a-529">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="03f8a-530">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03f8a-530">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="03f8a-531">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-531">String</span></span> | <span data-ttu-id="03f8a-532">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-532">A string containing the subject of the message.</span></span> <span data-ttu-id="03f8a-533">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03f8a-533">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="03f8a-534">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-534">String</span></span> | <span data-ttu-id="03f8a-535">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-535">The HTML body of the message.</span></span> <span data-ttu-id="03f8a-536">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="03f8a-536">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="03f8a-537">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-537">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="03f8a-538">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-538">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="03f8a-539">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-539">String</span></span> | <span data-ttu-id="03f8a-p126">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="03f8a-542">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-542">String</span></span> | <span data-ttu-id="03f8a-543">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="03f8a-543">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="03f8a-544">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-544">String</span></span> | <span data-ttu-id="03f8a-p127">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="03f8a-547">Booliano</span><span class="sxs-lookup"><span data-stu-id="03f8a-547">Boolean</span></span> | <span data-ttu-id="03f8a-p128">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="03f8a-550">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="03f8a-550">String</span></span> | <span data-ttu-id="03f8a-551">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-551">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="03f8a-552">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="03f8a-552">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="03f8a-553">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03f8a-553">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="03f8a-554">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-554">Requirements</span></span>

|<span data-ttu-id="03f8a-555">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-555">Requirement</span></span>| <span data-ttu-id="03f8a-556">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-557">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-558">1.6</span><span class="sxs-lookup"><span data-stu-id="03f8a-558">1.6</span></span> |
|[<span data-ttu-id="03f8a-559">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-560">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-561">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-562">Read</span><span class="sxs-lookup"><span data-stu-id="03f8a-562">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-563">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-563">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="03f8a-564">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="03f8a-564">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="03f8a-565">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="03f8a-565">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="03f8a-p130">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-568">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="03f8a-568">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="03f8a-569">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-569">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="03f8a-570">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-570">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="03f8a-571">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-571">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="03f8a-572">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="03f8a-572">**REST Tokens**</span></span>

<span data-ttu-id="03f8a-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="03f8a-576">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="03f8a-576">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="03f8a-577">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="03f8a-577">**EWS Tokens**</span></span>

<span data-ttu-id="03f8a-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="03f8a-580">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="03f8a-580">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="03f8a-581">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="03f8a-581">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="03f8a-582">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para recuperar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-582">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="03f8a-583">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03f8a-583">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-584">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-584">Parameters</span></span>

|<span data-ttu-id="03f8a-585">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-585">Name</span></span>| <span data-ttu-id="03f8a-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-586">Type</span></span>| <span data-ttu-id="03f8a-587">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-587">Attributes</span></span>| <span data-ttu-id="03f8a-588">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-588">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="03f8a-589">Object</span><span class="sxs-lookup"><span data-stu-id="03f8a-589">Object</span></span> | <span data-ttu-id="03f8a-590">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-590">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-591">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03f8a-591">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="03f8a-592">Booliano</span><span class="sxs-lookup"><span data-stu-id="03f8a-592">Boolean</span></span> |  <span data-ttu-id="03f8a-593">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-593">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03f8a-596">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-596">Object</span></span> |  <span data-ttu-id="03f8a-597">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-597">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-598">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03f8a-598">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="03f8a-599">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-599">function</span></span>||<span data-ttu-id="03f8a-600">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-600">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03f8a-601">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-601">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03f8a-602">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-602">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03f8a-603">Erros</span><span class="sxs-lookup"><span data-stu-id="03f8a-603">Errors</span></span>

|<span data-ttu-id="03f8a-604">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03f8a-604">Error code</span></span>|<span data-ttu-id="03f8a-605">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-605">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03f8a-606">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03f8a-606">The request has failed.</span></span> <span data-ttu-id="03f8a-607">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03f8a-607">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03f8a-608">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03f8a-608">The Exchange server returned an error.</span></span> <span data-ttu-id="03f8a-609">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03f8a-609">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03f8a-610">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03f8a-610">The user is no longer connected to the network.</span></span> <span data-ttu-id="03f8a-611">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-611">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-612">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-612">Requirements</span></span>

|<span data-ttu-id="03f8a-613">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-613">Requirement</span></span>| <span data-ttu-id="03f8a-614">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-615">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-616">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-616">1.5</span></span> |
|[<span data-ttu-id="03f8a-617">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-618">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-619">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-620">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="03f8a-620">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-621">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-621">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="03f8a-622">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03f8a-622">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="03f8a-623">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="03f8a-623">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="03f8a-p139">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="03f8a-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="03f8a-626">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="03f8a-626">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="03f8a-627">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-627">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="03f8a-628">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03f8a-628">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="03f8a-629">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-629">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="03f8a-630">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-630">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="03f8a-631">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="03f8a-631">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-632">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-632">Parameters</span></span>

|<span data-ttu-id="03f8a-633">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-633">Name</span></span>| <span data-ttu-id="03f8a-634">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-634">Type</span></span>| <span data-ttu-id="03f8a-635">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-635">Attributes</span></span>| <span data-ttu-id="03f8a-636">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-636">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="03f8a-637">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-637">function</span></span>||<span data-ttu-id="03f8a-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03f8a-639">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-639">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03f8a-640">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-640">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="03f8a-641">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-641">Object</span></span>| <span data-ttu-id="03f8a-642">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-642">&lt;optional&gt;</span></span>|<span data-ttu-id="03f8a-643">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03f8a-643">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03f8a-644">Erros</span><span class="sxs-lookup"><span data-stu-id="03f8a-644">Errors</span></span>

|<span data-ttu-id="03f8a-645">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03f8a-645">Error code</span></span>|<span data-ttu-id="03f8a-646">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-646">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03f8a-647">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03f8a-647">The request has failed.</span></span> <span data-ttu-id="03f8a-648">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03f8a-648">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03f8a-649">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03f8a-649">The Exchange server returned an error.</span></span> <span data-ttu-id="03f8a-650">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03f8a-650">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03f8a-651">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03f8a-651">The user is no longer connected to the network.</span></span> <span data-ttu-id="03f8a-652">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-652">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-653">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-653">Requirements</span></span>

|<span data-ttu-id="03f8a-654">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-654">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="03f8a-655">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-656">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-656">1.0</span></span> | <span data-ttu-id="03f8a-657">1.3</span><span class="sxs-lookup"><span data-stu-id="03f8a-657">1.3</span></span> |
|[<span data-ttu-id="03f8a-658">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-658">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-659">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-659">ReadItem</span></span> | <span data-ttu-id="03f8a-660">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-660">ReadItem</span></span> |
|[<span data-ttu-id="03f8a-661">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-661">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-662">Read</span><span class="sxs-lookup"><span data-stu-id="03f8a-662">Read</span></span> | <span data-ttu-id="03f8a-663">Escrever</span><span class="sxs-lookup"><span data-stu-id="03f8a-663">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="03f8a-664">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-664">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="03f8a-665">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03f8a-665">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="03f8a-666">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="03f8a-666">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="03f8a-667">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="03f8a-667">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-668">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-668">Parameters</span></span>

|<span data-ttu-id="03f8a-669">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-669">Name</span></span>| <span data-ttu-id="03f8a-670">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-670">Type</span></span>| <span data-ttu-id="03f8a-671">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-671">Attributes</span></span>| <span data-ttu-id="03f8a-672">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-672">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="03f8a-673">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-673">function</span></span>||<span data-ttu-id="03f8a-674">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-674">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03f8a-675">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-675">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03f8a-676">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03f8a-676">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="03f8a-677">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-677">Object</span></span>| <span data-ttu-id="03f8a-678">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-678">&lt;optional&gt;</span></span>|<span data-ttu-id="03f8a-679">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03f8a-679">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03f8a-680">Erros</span><span class="sxs-lookup"><span data-stu-id="03f8a-680">Errors</span></span>

|<span data-ttu-id="03f8a-681">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03f8a-681">Error code</span></span>|<span data-ttu-id="03f8a-682">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-682">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03f8a-683">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03f8a-683">The request has failed.</span></span> <span data-ttu-id="03f8a-684">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03f8a-684">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03f8a-685">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03f8a-685">The Exchange server returned an error.</span></span> <span data-ttu-id="03f8a-686">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03f8a-686">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03f8a-687">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03f8a-687">The user is no longer connected to the network.</span></span> <span data-ttu-id="03f8a-688">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03f8a-688">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-689">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-689">Requirements</span></span>

|<span data-ttu-id="03f8a-690">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-690">Requirement</span></span>| <span data-ttu-id="03f8a-691">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-691">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-692">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-692">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-693">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-693">1.0</span></span>|
|[<span data-ttu-id="03f8a-694">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-694">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-695">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-695">ReadItem</span></span>|
|[<span data-ttu-id="03f8a-696">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-696">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-697">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-697">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-698">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-698">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="03f8a-699">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03f8a-699">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="03f8a-700">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03f8a-700">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-701">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="03f8a-701">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="03f8a-702">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="03f8a-702">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="03f8a-703">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="03f8a-703">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="03f8a-704">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03f8a-704">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="03f8a-705">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="03f8a-705">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="03f8a-706">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="03f8a-706">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="03f8a-707">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-707">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="03f8a-708">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="03f8a-708">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="03f8a-p149">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="03f8a-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="03f8a-711">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="03f8a-711">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="03f8a-712">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="03f8a-712">Version differences</span></span>

<span data-ttu-id="03f8a-713">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-713">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="03f8a-714">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="03f8a-714">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="03f8a-715">Você pode determinar se o seu aplicativo de email está em execução no Outlook na Web ou em um cliente de desktop usando a propriedade Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="03f8a-715">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="03f8a-716">Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="03f8a-716">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-717">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-717">Parameters</span></span>

|<span data-ttu-id="03f8a-718">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-718">Name</span></span>| <span data-ttu-id="03f8a-719">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-719">Type</span></span>| <span data-ttu-id="03f8a-720">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-720">Attributes</span></span>| <span data-ttu-id="03f8a-721">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-721">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="03f8a-722">String</span><span class="sxs-lookup"><span data-stu-id="03f8a-722">String</span></span>||<span data-ttu-id="03f8a-723">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="03f8a-723">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="03f8a-724">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-724">function</span></span>||<span data-ttu-id="03f8a-725">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03f8a-726">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03f8a-726">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="03f8a-727">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="03f8a-727">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="03f8a-728">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-728">Object</span></span>| <span data-ttu-id="03f8a-729">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-729">&lt;optional&gt;</span></span>|<span data-ttu-id="03f8a-730">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03f8a-730">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-731">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-731">Requirements</span></span>

|<span data-ttu-id="03f8a-732">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-732">Requirement</span></span>| <span data-ttu-id="03f8a-733">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-734">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-735">1.0</span><span class="sxs-lookup"><span data-stu-id="03f8a-735">1.0</span></span>|
|[<span data-ttu-id="03f8a-736">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-737">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="03f8a-737">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="03f8a-738">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03f8a-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-739">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-739">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03f8a-740">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03f8a-740">Example</span></span>

<span data-ttu-id="03f8a-741">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="03f8a-741">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="03f8a-742">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03f8a-742">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="03f8a-743">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="03f8a-743">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="03f8a-744">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="03f8a-744">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03f8a-745">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03f8a-745">Parameters</span></span>

| <span data-ttu-id="03f8a-746">Nome</span><span class="sxs-lookup"><span data-stu-id="03f8a-746">Name</span></span> | <span data-ttu-id="03f8a-747">Tipo</span><span class="sxs-lookup"><span data-stu-id="03f8a-747">Type</span></span> | <span data-ttu-id="03f8a-748">Atributos</span><span class="sxs-lookup"><span data-stu-id="03f8a-748">Attributes</span></span> | <span data-ttu-id="03f8a-749">Descrição</span><span class="sxs-lookup"><span data-stu-id="03f8a-749">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="03f8a-750">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="03f8a-750">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="03f8a-751">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="03f8a-751">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="03f8a-752">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-752">Object</span></span> | <span data-ttu-id="03f8a-753">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-753">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-754">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03f8a-754">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03f8a-755">Objeto</span><span class="sxs-lookup"><span data-stu-id="03f8a-755">Object</span></span> | <span data-ttu-id="03f8a-756">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-756">&lt;optional&gt;</span></span> | <span data-ttu-id="03f8a-757">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="03f8a-757">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="03f8a-758">function</span><span class="sxs-lookup"><span data-stu-id="03f8a-758">function</span></span>| <span data-ttu-id="03f8a-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03f8a-759">&lt;optional&gt;</span></span>|<span data-ttu-id="03f8a-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03f8a-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03f8a-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03f8a-761">Requirements</span></span>

|<span data-ttu-id="03f8a-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="03f8a-762">Requirement</span></span>| <span data-ttu-id="03f8a-763">Valor</span><span class="sxs-lookup"><span data-stu-id="03f8a-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="03f8a-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03f8a-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03f8a-765">1,5</span><span class="sxs-lookup"><span data-stu-id="03f8a-765">1.5</span></span> |
|[<span data-ttu-id="03f8a-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03f8a-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03f8a-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03f8a-767">ReadItem</span></span> |
|[<span data-ttu-id="03f8a-768">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03f8a-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03f8a-769">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03f8a-769">Compose or Read</span></span>|
