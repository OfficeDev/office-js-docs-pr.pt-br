---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: a71d3869d5dbf91db7823118a8d0409699e17cd5
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629220"
---
# <a name="item"></a><span data-ttu-id="48574-102">item</span><span class="sxs-lookup"><span data-stu-id="48574-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="48574-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="48574-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="48574-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-mailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="48574-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-mailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-106">Requirements</span></span>

|<span data-ttu-id="48574-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-107">Requirement</span></span>|<span data-ttu-id="48574-108">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-110">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-110">1.0</span></span>|
|[<span data-ttu-id="48574-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="48574-112">Restricted</span></span>|
|[<span data-ttu-id="48574-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-114">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="48574-115">Propriedades</span><span class="sxs-lookup"><span data-stu-id="48574-115">Properties</span></span>

| <span data-ttu-id="48574-116">Propriedade</span><span class="sxs-lookup"><span data-stu-id="48574-116">Property</span></span> | <span data-ttu-id="48574-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-117">Minimum</span></span><br><span data-ttu-id="48574-118">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="48574-118">permission level</span></span> | <span data-ttu-id="48574-119">Modelos</span><span class="sxs-lookup"><span data-stu-id="48574-119">Modes</span></span> | <span data-ttu-id="48574-120">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="48574-120">Return type</span></span> | <span data-ttu-id="48574-121">Mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-121">Minimum</span></span><br><span data-ttu-id="48574-122">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-122">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="48574-123">attachments</span><span class="sxs-lookup"><span data-stu-id="48574-123">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="48574-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-124">ReadItem</span></span> | <span data-ttu-id="48574-125">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-125">Read</span></span> | <span data-ttu-id="48574-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48574-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span> | <span data-ttu-id="48574-127">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-127">1.0</span></span> |
| [<span data-ttu-id="48574-128">bcc</span><span class="sxs-lookup"><span data-stu-id="48574-128">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="48574-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-129">ReadItem</span></span> | <span data-ttu-id="48574-130">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-130">Message Compose</span></span> | [<span data-ttu-id="48574-131">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-131">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="48574-132">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-132">1.1</span></span> |
| [<span data-ttu-id="48574-133">body</span><span class="sxs-lookup"><span data-stu-id="48574-133">body</span></span>](#body-body) | <span data-ttu-id="48574-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-134">ReadItem</span></span> | <span data-ttu-id="48574-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-135">Compose</span></span> | [<span data-ttu-id="48574-136">Body</span><span class="sxs-lookup"><span data-stu-id="48574-136">Body</span></span>](/javascript/api/outlook/office.body) | <span data-ttu-id="48574-137">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-137">1.1</span></span> |
| | | <span data-ttu-id="48574-138">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-138">Read</span></span> | | |
| [<span data-ttu-id="48574-139">categories</span><span class="sxs-lookup"><span data-stu-id="48574-139">categories</span></span>](#categories-categories) | <span data-ttu-id="48574-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-140">ReadItem</span></span> | <span data-ttu-id="48574-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-141">Compose</span></span> | [<span data-ttu-id="48574-142">Categories</span><span class="sxs-lookup"><span data-stu-id="48574-142">Categories</span></span>](/javascript/api/outlook/office.categories) | <span data-ttu-id="48574-143">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-143">Preview</span></span> |
| | | <span data-ttu-id="48574-144">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-144">Read</span></span> | | |
| [<span data-ttu-id="48574-145">cc</span><span class="sxs-lookup"><span data-stu-id="48574-145">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48574-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-146">ReadItem</span></span> | <span data-ttu-id="48574-147">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-147">Message Compose</span></span> | [<span data-ttu-id="48574-148">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-148">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="48574-149">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-149">1.0</span></span> |
| | | <span data-ttu-id="48574-150">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-150">Message Read</span></span> | <span data-ttu-id="48574-151">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <></span><span class="sxs-lookup"><span data-stu-id="48574-151">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="48574-152">conversationId</span><span class="sxs-lookup"><span data-stu-id="48574-152">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="48574-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-153">ReadItem</span></span> | <span data-ttu-id="48574-154">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-154">Message Compose</span></span> | <span data-ttu-id="48574-155">String</span><span class="sxs-lookup"><span data-stu-id="48574-155">String</span></span> | <span data-ttu-id="48574-156">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-156">1.0</span></span> |
| | | <span data-ttu-id="48574-157">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-157">Message Read</span></span> | | |
| [<span data-ttu-id="48574-158">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="48574-158">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="48574-159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-159">ReadItem</span></span> | <span data-ttu-id="48574-160">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-160">Read</span></span> | <span data-ttu-id="48574-161">Data</span><span class="sxs-lookup"><span data-stu-id="48574-161">Date</span></span> | <span data-ttu-id="48574-162">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-162">1.0</span></span> |
| [<span data-ttu-id="48574-163">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="48574-163">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="48574-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-164">ReadItem</span></span> | <span data-ttu-id="48574-165">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-165">Read</span></span> | <span data-ttu-id="48574-166">Data</span><span class="sxs-lookup"><span data-stu-id="48574-166">Date</span></span> | <span data-ttu-id="48574-167">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-167">1.0</span></span> |
| [<span data-ttu-id="48574-168">end</span><span class="sxs-lookup"><span data-stu-id="48574-168">end</span></span>](#end-datetime) | <span data-ttu-id="48574-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-169">ReadItem</span></span> | <span data-ttu-id="48574-170">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-170">Appointment Organizer</span></span> | [<span data-ttu-id="48574-171">Time</span><span class="sxs-lookup"><span data-stu-id="48574-171">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="48574-172">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-172">1.0</span></span> |
| | | <span data-ttu-id="48574-173">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-173">Appointment Attendee</span></span> | <span data-ttu-id="48574-174">Data</span><span class="sxs-lookup"><span data-stu-id="48574-174">Date</span></span> | |
| | | <span data-ttu-id="48574-175">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-175">Message Read</span></span><br><span data-ttu-id="48574-176">(Solicitação de reunião)</span><span class="sxs-lookup"><span data-stu-id="48574-176">(Meeting Request)</span></span> | <span data-ttu-id="48574-177">Data</span><span class="sxs-lookup"><span data-stu-id="48574-177">Date</span></span> | |
| [<span data-ttu-id="48574-178">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="48574-178">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="48574-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-179">ReadItem</span></span> | <span data-ttu-id="48574-180">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-180">Appointment Organizer</span></span> | [<span data-ttu-id="48574-181">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="48574-181">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation) | <span data-ttu-id="48574-182">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-182">Preview</span></span> |
| | | <span data-ttu-id="48574-183">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-183">Appointment Attendee</span></span> | | |
| [<span data-ttu-id="48574-184">from</span><span class="sxs-lookup"><span data-stu-id="48574-184">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="48574-185">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-185">ReadWriteItem</span></span> | <span data-ttu-id="48574-186">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-186">Message Compose</span></span> | [<span data-ttu-id="48574-187">De</span><span class="sxs-lookup"><span data-stu-id="48574-187">From</span></span>](/javascript/api/outlook/office.from) | <span data-ttu-id="48574-188">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-188">1.7</span></span> |
| | <span data-ttu-id="48574-189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-189">ReadItem</span></span> | <span data-ttu-id="48574-190">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-190">Message Read</span></span> | [<span data-ttu-id="48574-191">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-191">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="48574-192">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-192">1.0</span></span> |
| [<span data-ttu-id="48574-193">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="48574-193">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="48574-194">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-194">ReadItem</span></span> | <span data-ttu-id="48574-195">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-195">Message Compose</span></span> | [<span data-ttu-id="48574-196">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="48574-196">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders) | <span data-ttu-id="48574-197">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-197">Preview</span></span> |
| [<span data-ttu-id="48574-198">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="48574-198">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="48574-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-199">ReadItem</span></span> | <span data-ttu-id="48574-200">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-200">Message Read</span></span> | <span data-ttu-id="48574-201">String</span><span class="sxs-lookup"><span data-stu-id="48574-201">String</span></span> | <span data-ttu-id="48574-202">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-202">1.0</span></span> |
| [<span data-ttu-id="48574-203">itemClass</span><span class="sxs-lookup"><span data-stu-id="48574-203">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="48574-204">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-204">ReadItem</span></span> | <span data-ttu-id="48574-205">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-205">Read</span></span> | <span data-ttu-id="48574-206">String</span><span class="sxs-lookup"><span data-stu-id="48574-206">String</span></span> | <span data-ttu-id="48574-207">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-207">1.0</span></span> |
| [<span data-ttu-id="48574-208">itemId</span><span class="sxs-lookup"><span data-stu-id="48574-208">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="48574-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-209">ReadItem</span></span> | <span data-ttu-id="48574-210">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-210">Read</span></span> | <span data-ttu-id="48574-211">String</span><span class="sxs-lookup"><span data-stu-id="48574-211">String</span></span> | <span data-ttu-id="48574-212">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-212">1.0</span></span> |
| [<span data-ttu-id="48574-213">itemType</span><span class="sxs-lookup"><span data-stu-id="48574-213">itemType</span></span>](#itemtype-mailboxenumsitemtype) | <span data-ttu-id="48574-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-214">ReadItem</span></span> | <span data-ttu-id="48574-215">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-215">Compose</span></span> | [<span data-ttu-id="48574-216">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="48574-216">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype) | <span data-ttu-id="48574-217">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-217">1.0</span></span> |
| | | <span data-ttu-id="48574-218">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-218">Read</span></span> | | |
| [<span data-ttu-id="48574-219">location</span><span class="sxs-lookup"><span data-stu-id="48574-219">location</span></span>](#location-stringlocation) | <span data-ttu-id="48574-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-220">ReadItem</span></span> | <span data-ttu-id="48574-221">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-221">Appointment Organizer</span></span> | [<span data-ttu-id="48574-222">Location</span><span class="sxs-lookup"><span data-stu-id="48574-222">Location</span></span>](/javascript/api/outlook/office.location) | <span data-ttu-id="48574-223">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-223">1.0</span></span> |
| | | <span data-ttu-id="48574-224">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-224">Appointment Attendee</span></span> | <span data-ttu-id="48574-225">String</span><span class="sxs-lookup"><span data-stu-id="48574-225">String</span></span> | |
| | | <span data-ttu-id="48574-226">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-226">Message Read</span></span><br><span data-ttu-id="48574-227">(Solicitação de reunião)</span><span class="sxs-lookup"><span data-stu-id="48574-227">(Meeting Request)</span></span> | <span data-ttu-id="48574-228">String</span><span class="sxs-lookup"><span data-stu-id="48574-228">String</span></span> | |
| [<span data-ttu-id="48574-229">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="48574-229">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="48574-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-230">ReadItem</span></span> | <span data-ttu-id="48574-231">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-231">Read</span></span> | <span data-ttu-id="48574-232">String</span><span class="sxs-lookup"><span data-stu-id="48574-232">String</span></span> | <span data-ttu-id="48574-233">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-233">1.0</span></span> |
| [<span data-ttu-id="48574-234">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="48574-234">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="48574-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-235">ReadItem</span></span> | <span data-ttu-id="48574-236">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-236">Message Compose</span></span> | [<span data-ttu-id="48574-237">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="48574-237">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages) | <span data-ttu-id="48574-238">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-238">1.3</span></span> |
| | <span data-ttu-id="48574-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-239">ReadItem</span></span> | <span data-ttu-id="48574-240">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-240">Message Read</span></span> | | |
| [<span data-ttu-id="48574-241">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="48574-241">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48574-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-242">ReadItem</span></span> | <span data-ttu-id="48574-243">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-243">Appointment Organizer</span></span> | [<span data-ttu-id="48574-244">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-244">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="48574-245">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-245">1.0</span></span> |
| | | <span data-ttu-id="48574-246">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-246">Appointment Attendee</span></span> | <span data-ttu-id="48574-247">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <></span><span class="sxs-lookup"><span data-stu-id="48574-247">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="48574-248">organizer</span><span class="sxs-lookup"><span data-stu-id="48574-248">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="48574-249">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-249">ReadWriteItem</span></span> | <span data-ttu-id="48574-250">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-250">Appointment Organizer</span></span> | [<span data-ttu-id="48574-251">Organizador</span><span class="sxs-lookup"><span data-stu-id="48574-251">Organizer</span></span>](/javascript/api/outlook/office.organizer) | <span data-ttu-id="48574-252">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-252">1.7</span></span> |
| | <span data-ttu-id="48574-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-253">ReadItem</span></span> | <span data-ttu-id="48574-254">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-254">Appointment Attendee</span></span> | [<span data-ttu-id="48574-255">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-255">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="48574-256">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-256">1.0</span></span> |
| [<span data-ttu-id="48574-257">recurrence</span><span class="sxs-lookup"><span data-stu-id="48574-257">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="48574-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-258">ReadItem</span></span> | <span data-ttu-id="48574-259">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-259">Appointment Organizer</span></span> | [<span data-ttu-id="48574-260">Recorrência</span><span class="sxs-lookup"><span data-stu-id="48574-260">Recurrence</span></span>](/javascript/api/outlook/office.recurrence) | <span data-ttu-id="48574-261">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-261">1.7</span></span> |
| | | <span data-ttu-id="48574-262">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-262">Appointment Attendee</span></span> | | |
| | | <span data-ttu-id="48574-263">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-263">Message Read</span></span><br><span data-ttu-id="48574-264">(Solicitação de reunião)</span><span class="sxs-lookup"><span data-stu-id="48574-264">(Meeting Request)</span></span> | | |
| [<span data-ttu-id="48574-265">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="48574-265">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48574-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-266">ReadItem</span></span> | <span data-ttu-id="48574-267">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-267">Appointment Organizer</span></span> | [<span data-ttu-id="48574-268">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-268">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="48574-269">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-269">1.0</span></span> |
| | | <span data-ttu-id="48574-270">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-270">Appointment Attendee</span></span> | <span data-ttu-id="48574-271">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <></span><span class="sxs-lookup"><span data-stu-id="48574-271">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="48574-272">sender</span><span class="sxs-lookup"><span data-stu-id="48574-272">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="48574-273">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-273">ReadItem</span></span> | <span data-ttu-id="48574-274">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-274">Message Read</span></span> | [<span data-ttu-id="48574-275">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-275">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="48574-276">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-276">1.0</span></span> |
| [<span data-ttu-id="48574-277">seriesid</span><span class="sxs-lookup"><span data-stu-id="48574-277">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="48574-278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-278">ReadItem</span></span> | <span data-ttu-id="48574-279">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-279">Compose</span></span> | <span data-ttu-id="48574-280">String</span><span class="sxs-lookup"><span data-stu-id="48574-280">String</span></span> | <span data-ttu-id="48574-281">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-281">1.7</span></span> |
| | | <span data-ttu-id="48574-282">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-282">Read</span></span> | | |
| [<span data-ttu-id="48574-283">start</span><span class="sxs-lookup"><span data-stu-id="48574-283">start</span></span>](#start-datetime) | <span data-ttu-id="48574-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-284">ReadItem</span></span> | <span data-ttu-id="48574-285">Organizador de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-285">Appointment Organizer</span></span> | [<span data-ttu-id="48574-286">Time</span><span class="sxs-lookup"><span data-stu-id="48574-286">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="48574-287">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-287">1.0</span></span> |
| | | <span data-ttu-id="48574-288">Participante do compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-288">Appointment Attendee</span></span> | <span data-ttu-id="48574-289">Data</span><span class="sxs-lookup"><span data-stu-id="48574-289">Date</span></span> | |
| | | <span data-ttu-id="48574-290">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-290">Message Read</span></span><br><span data-ttu-id="48574-291">(Solicitação de reunião)</span><span class="sxs-lookup"><span data-stu-id="48574-291">(Meeting Request)</span></span> | <span data-ttu-id="48574-292">Data</span><span class="sxs-lookup"><span data-stu-id="48574-292">Date</span></span> | |
| [<span data-ttu-id="48574-293">subject</span><span class="sxs-lookup"><span data-stu-id="48574-293">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="48574-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-294">ReadItem</span></span> | <span data-ttu-id="48574-295">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-295">Compose</span></span> | [<span data-ttu-id="48574-296">Subject</span><span class="sxs-lookup"><span data-stu-id="48574-296">Subject</span></span>](/javascript/api/outlook/office.subject) | <span data-ttu-id="48574-297">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-297">1.0</span></span> |
| | | <span data-ttu-id="48574-298">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-298">Read</span></span> | <span data-ttu-id="48574-299">String</span><span class="sxs-lookup"><span data-stu-id="48574-299">String</span></span> | |
| [<span data-ttu-id="48574-300">to</span><span class="sxs-lookup"><span data-stu-id="48574-300">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48574-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-301">ReadItem</span></span> | <span data-ttu-id="48574-302">Composição de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-302">Message Compose</span></span> | [<span data-ttu-id="48574-303">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-303">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="48574-304">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-304">1.0</span></span> |
| | | <span data-ttu-id="48574-305">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-305">Message Read</span></span> | <span data-ttu-id="48574-306">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) de matriz. <></span><span class="sxs-lookup"><span data-stu-id="48574-306">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |

##### <a name="methods"></a><span data-ttu-id="48574-307">Métodos</span><span class="sxs-lookup"><span data-stu-id="48574-307">Methods</span></span>

| <span data-ttu-id="48574-308">Método</span><span class="sxs-lookup"><span data-stu-id="48574-308">Method</span></span> | <span data-ttu-id="48574-309">Mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-309">Minimum</span></span><br><span data-ttu-id="48574-310">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="48574-310">permission level</span></span> | <span data-ttu-id="48574-311">Modelos</span><span class="sxs-lookup"><span data-stu-id="48574-311">Modes</span></span> | <span data-ttu-id="48574-312">Mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-312">Minimum</span></span><br><span data-ttu-id="48574-313">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-313">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="48574-314">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48574-314">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="48574-315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-315">ReadWriteItem</span></span> | <span data-ttu-id="48574-316">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-316">Compose</span></span> | <span data-ttu-id="48574-317">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-317">1.1</span></span> |
| [<span data-ttu-id="48574-318">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="48574-318">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="48574-319">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-319">ReadWriteItem</span></span> | <span data-ttu-id="48574-320">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-320">Compose</span></span> | <span data-ttu-id="48574-321">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-321">Preview</span></span> |
| [<span data-ttu-id="48574-322">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48574-322">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="48574-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-323">ReadItem</span></span> | <span data-ttu-id="48574-324">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-324">Compose</span></span><br><span data-ttu-id="48574-325">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-325">Read</span></span> | <span data-ttu-id="48574-326">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-326">1.7</span></span> |
| [<span data-ttu-id="48574-327">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48574-327">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="48574-328">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-328">ReadWriteItem</span></span> | <span data-ttu-id="48574-329">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-329">Compose</span></span> | <span data-ttu-id="48574-330">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-330">1.1</span></span> |
| [<span data-ttu-id="48574-331">close</span><span class="sxs-lookup"><span data-stu-id="48574-331">close</span></span>](#close) | <span data-ttu-id="48574-332">Restrito</span><span class="sxs-lookup"><span data-stu-id="48574-332">Restricted</span></span> | <span data-ttu-id="48574-333">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-333">Compose</span></span> | <span data-ttu-id="48574-334">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-334">1.3</span></span> |
| [<span data-ttu-id="48574-335">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="48574-335">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="48574-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-336">ReadItem</span></span> | <span data-ttu-id="48574-337">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-337">Read</span></span> | <span data-ttu-id="48574-338">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-338">1.0</span></span> |
| [<span data-ttu-id="48574-339">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="48574-339">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="48574-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-340">ReadItem</span></span> | <span data-ttu-id="48574-341">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-341">Read</span></span> | <span data-ttu-id="48574-342">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-342">1.0</span></span> |
| [<span data-ttu-id="48574-343">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="48574-343">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="48574-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-344">ReadItem</span></span> | <span data-ttu-id="48574-345">Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="48574-345">Message Read</span></span> | <span data-ttu-id="48574-346">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-346">1.8</span></span> |
| [<span data-ttu-id="48574-347">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="48574-347">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="48574-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-348">ReadItem</span></span> | <span data-ttu-id="48574-349">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-349">Compose</span></span><br><span data-ttu-id="48574-350">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-350">Read</span></span> | <span data-ttu-id="48574-351">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-351">Preview</span></span> |
| [<span data-ttu-id="48574-352">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="48574-352">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="48574-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-353">ReadItem</span></span> | <span data-ttu-id="48574-354">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-354">Compose</span></span> | <span data-ttu-id="48574-355">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-355">Preview</span></span> |
| [<span data-ttu-id="48574-356">getEntities</span><span class="sxs-lookup"><span data-stu-id="48574-356">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="48574-357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-357">ReadItem</span></span> | <span data-ttu-id="48574-358">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-358">Read</span></span> | <span data-ttu-id="48574-359">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-359">1.0</span></span> |
| [<span data-ttu-id="48574-360">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="48574-360">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="48574-361">Restrito</span><span class="sxs-lookup"><span data-stu-id="48574-361">Restricted</span></span> | <span data-ttu-id="48574-362">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-362">Read</span></span> | <span data-ttu-id="48574-363">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-363">1.0</span></span> |
| [<span data-ttu-id="48574-364">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="48574-364">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="48574-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-365">ReadItem</span></span> | <span data-ttu-id="48574-366">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-366">Read</span></span> | <span data-ttu-id="48574-367">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-367">1.0</span></span> |
| [<span data-ttu-id="48574-368">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="48574-368">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="48574-369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-369">ReadItem</span></span> | <span data-ttu-id="48574-370">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-370">Read</span></span> | <span data-ttu-id="48574-371">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-371">Preview</span></span> |
| [<span data-ttu-id="48574-372">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="48574-372">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="48574-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-373">ReadItem</span></span> | <span data-ttu-id="48574-374">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-374">Compose</span></span> | <span data-ttu-id="48574-375">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-375">Preview</span></span> |
| [<span data-ttu-id="48574-376">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="48574-376">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="48574-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-377">ReadItem</span></span> | <span data-ttu-id="48574-378">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-378">Read</span></span> | <span data-ttu-id="48574-379">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-379">1.0</span></span> |
| [<span data-ttu-id="48574-380">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="48574-380">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="48574-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-381">ReadItem</span></span> | <span data-ttu-id="48574-382">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-382">Read</span></span> | <span data-ttu-id="48574-383">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-383">1.0</span></span> |
| [<span data-ttu-id="48574-384">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="48574-384">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="48574-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-385">ReadItem</span></span> | <span data-ttu-id="48574-386">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-386">Compose</span></span> | <span data-ttu-id="48574-387">1.2</span><span class="sxs-lookup"><span data-stu-id="48574-387">1.2</span></span> |
| [<span data-ttu-id="48574-388">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="48574-388">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="48574-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-389">ReadItem</span></span> | <span data-ttu-id="48574-390">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-390">Read</span></span> | <span data-ttu-id="48574-391">1.6</span><span class="sxs-lookup"><span data-stu-id="48574-391">1.6</span></span> |
| [<span data-ttu-id="48574-392">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="48574-392">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="48574-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-393">ReadItem</span></span> | <span data-ttu-id="48574-394">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-394">Read</span></span> | <span data-ttu-id="48574-395">1.6</span><span class="sxs-lookup"><span data-stu-id="48574-395">1.6</span></span> |
| [<span data-ttu-id="48574-396">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="48574-396">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="48574-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-397">ReadItem</span></span> | <span data-ttu-id="48574-398">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-398">Compose</span></span><br><span data-ttu-id="48574-399">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-399">Read</span></span> | <span data-ttu-id="48574-400">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-400">Preview</span></span> |
| [<span data-ttu-id="48574-401">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="48574-401">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="48574-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-402">ReadItem</span></span> | <span data-ttu-id="48574-403">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-403">Compose</span></span><br><span data-ttu-id="48574-404">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-404">Read</span></span> | <span data-ttu-id="48574-405">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-405">1.0</span></span> |
| [<span data-ttu-id="48574-406">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48574-406">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="48574-407">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-407">ReadWriteItem</span></span> | <span data-ttu-id="48574-408">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-408">Compose</span></span> | <span data-ttu-id="48574-409">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-409">1.1</span></span> |
| [<span data-ttu-id="48574-410">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48574-410">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="48574-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-411">ReadItem</span></span> | <span data-ttu-id="48574-412">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-412">Compose</span></span><br><span data-ttu-id="48574-413">Ler</span><span class="sxs-lookup"><span data-stu-id="48574-413">Read</span></span> | <span data-ttu-id="48574-414">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-414">1.7</span></span> |
| [<span data-ttu-id="48574-415">saveAsync</span><span class="sxs-lookup"><span data-stu-id="48574-415">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="48574-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-416">ReadWriteItem</span></span> | <span data-ttu-id="48574-417">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-417">Compose</span></span> | <span data-ttu-id="48574-418">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-418">1.3</span></span> |
| [<span data-ttu-id="48574-419">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="48574-419">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="48574-420">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-420">ReadWriteItem</span></span> | <span data-ttu-id="48574-421">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-421">Compose</span></span> | <span data-ttu-id="48574-422">1.2</span><span class="sxs-lookup"><span data-stu-id="48574-422">1.2</span></span> |

##### <a name="events"></a><span data-ttu-id="48574-423">Eventos</span><span class="sxs-lookup"><span data-stu-id="48574-423">Events</span></span>

<span data-ttu-id="48574-424">Você pode assinar e cancelar a assinatura dos eventos a seguir usando o [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) e o [removeHandlerAsync](#removehandlerasynceventtype-options-callback) , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="48574-424">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="48574-425">Evento</span><span class="sxs-lookup"><span data-stu-id="48574-425">Event</span></span> | <span data-ttu-id="48574-426">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-426">Description</span></span> | <span data-ttu-id="48574-427">Mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-427">Minimum</span></span><br><span data-ttu-id="48574-428">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-428">requirement set</span></span> |
|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="48574-429">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="48574-429">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="48574-430">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-430">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="48574-431">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="48574-431">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="48574-432">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-432">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="48574-433">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="48574-433">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="48574-434">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-434">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="48574-435">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="48574-435">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="48574-436">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-436">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="48574-437">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="48574-437">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="48574-438">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-438">1.7</span></span> |

### <a name="example"></a><span data-ttu-id="48574-439">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-439">Example</span></span>

<span data-ttu-id="48574-440">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="48574-440">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

## <a name="property-details"></a><span data-ttu-id="48574-441">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="48574-441">Property details</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="48574-442">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48574-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="48574-443">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="48574-443">Gets the item's attachments as an array.</span></span> <span data-ttu-id="48574-444">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-444">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-445">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="48574-445">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="48574-446">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="48574-446">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="48574-447">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-447">Type</span></span>

*   <span data-ttu-id="48574-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48574-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-449">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-449">Requirements</span></span>

|<span data-ttu-id="48574-450">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-450">Requirement</span></span>|<span data-ttu-id="48574-451">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-451">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-452">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-453">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-453">1.0</span></span>|
|[<span data-ttu-id="48574-454">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-455">ReadItem</span></span>|
|[<span data-ttu-id="48574-456">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-457">Read</span><span class="sxs-lookup"><span data-stu-id="48574-457">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-458">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-458">Example</span></span>

<span data-ttu-id="48574-459">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-459">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48574-460">cco :[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48574-461">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-461">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="48574-462">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="48574-462">Compose mode only.</span></span>

<span data-ttu-id="48574-463">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-463">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-464">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="48574-464">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48574-465">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-465">Get 500 members maximum.</span></span>
- <span data-ttu-id="48574-466">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="48574-466">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-467">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-467">Type</span></span>

*   [<span data-ttu-id="48574-468">Destinatários</span><span class="sxs-lookup"><span data-stu-id="48574-468">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="48574-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-469">Requirements</span></span>

|<span data-ttu-id="48574-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-470">Requirement</span></span>|<span data-ttu-id="48574-471">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-473">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-473">1.1</span></span>|
|[<span data-ttu-id="48574-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-475">ReadItem</span></span>|
|[<span data-ttu-id="48574-476">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-477">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-477">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-478">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="48574-479">corpo: [Corpo](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="48574-479">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="48574-480">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="48574-480">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-481">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-481">Type</span></span>

*   [<span data-ttu-id="48574-482">Body</span><span class="sxs-lookup"><span data-stu-id="48574-482">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="48574-483">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-483">Requirements</span></span>

|<span data-ttu-id="48574-484">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-484">Requirement</span></span>|<span data-ttu-id="48574-485">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-486">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-487">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-487">1.1</span></span>|
|[<span data-ttu-id="48574-488">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-488">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-489">ReadItem</span></span>|
|[<span data-ttu-id="48574-490">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-490">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-491">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-491">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-492">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-492">Example</span></span>

<span data-ttu-id="48574-493">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="48574-493">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="48574-494">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-494">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="48574-495">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="48574-495">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="48574-496">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="48574-496">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-497">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-497">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-498">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-498">Type</span></span>

*   [<span data-ttu-id="48574-499">Categories</span><span class="sxs-lookup"><span data-stu-id="48574-499">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="48574-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-500">Requirements</span></span>

|<span data-ttu-id="48574-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-501">Requirement</span></span>|<span data-ttu-id="48574-502">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-504">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-504">1.8</span></span>|
|[<span data-ttu-id="48574-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-506">ReadItem</span></span>|
|[<span data-ttu-id="48574-507">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-508">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-508">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-509">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-509">Example</span></span>

<span data-ttu-id="48574-510">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="48574-510">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48574-511">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48574-512">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-512">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="48574-513">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-513">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-514">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-514">Read mode</span></span>

<span data-ttu-id="48574-515">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-515">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="48574-516">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-516">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-517">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-517">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-518">Compose mode</span></span>

<span data-ttu-id="48574-519">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-519">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="48574-520">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-520">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-521">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="48574-521">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48574-522">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-522">Get 500 members maximum.</span></span>
- <span data-ttu-id="48574-523">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="48574-523">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48574-524">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-524">Type</span></span>

*   <span data-ttu-id="48574-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-526">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-526">Requirements</span></span>

|<span data-ttu-id="48574-527">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-527">Requirement</span></span>|<span data-ttu-id="48574-528">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-529">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-530">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-530">1.0</span></span>|
|[<span data-ttu-id="48574-531">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-531">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-532">ReadItem</span></span>|
|[<span data-ttu-id="48574-533">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-533">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-534">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-534">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="48574-535">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-535">(nullable) conversationId: String</span></span>

<span data-ttu-id="48574-536">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="48574-536">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="48574-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="48574-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="48574-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="48574-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-541">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-541">Type</span></span>

*   <span data-ttu-id="48574-542">String</span><span class="sxs-lookup"><span data-stu-id="48574-542">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-543">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-543">Requirements</span></span>

|<span data-ttu-id="48574-544">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-544">Requirement</span></span>|<span data-ttu-id="48574-545">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-547">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-547">1.0</span></span>|
|[<span data-ttu-id="48574-548">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-549">ReadItem</span></span>|
|[<span data-ttu-id="48574-550">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-551">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-552">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-552">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="48574-553">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="48574-553">dateTimeCreated: Date</span></span>

<span data-ttu-id="48574-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-556">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-556">Type</span></span>

*   <span data-ttu-id="48574-557">Data</span><span class="sxs-lookup"><span data-stu-id="48574-557">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-558">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-558">Requirements</span></span>

|<span data-ttu-id="48574-559">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-559">Requirement</span></span>|<span data-ttu-id="48574-560">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-561">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-562">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-562">1.0</span></span>|
|[<span data-ttu-id="48574-563">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-564">ReadItem</span></span>|
|[<span data-ttu-id="48574-565">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-566">Read</span><span class="sxs-lookup"><span data-stu-id="48574-566">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-567">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-567">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="48574-568">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="48574-568">dateTimeModified: Date</span></span>

<span data-ttu-id="48574-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-571">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-571">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-572">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-572">Type</span></span>

*   <span data-ttu-id="48574-573">Data</span><span class="sxs-lookup"><span data-stu-id="48574-573">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-574">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-574">Requirements</span></span>

|<span data-ttu-id="48574-575">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-575">Requirement</span></span>|<span data-ttu-id="48574-576">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-577">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-578">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-578">1.0</span></span>|
|[<span data-ttu-id="48574-579">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-580">ReadItem</span></span>|
|[<span data-ttu-id="48574-581">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-582">Read</span><span class="sxs-lookup"><span data-stu-id="48574-582">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-583">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-583">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="48574-584">fim: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48574-584">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="48574-585">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="48574-585">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="48574-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="48574-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-588">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-588">Read mode</span></span>

<span data-ttu-id="48574-589">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="48574-589">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-590">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-590">Compose mode</span></span>

<span data-ttu-id="48574-591">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="48574-591">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="48574-592">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-592">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="48574-593">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="48574-593">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="48574-594">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-594">Type</span></span>

*   <span data-ttu-id="48574-595">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48574-595">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-596">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-596">Requirements</span></span>

|<span data-ttu-id="48574-597">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-597">Requirement</span></span>|<span data-ttu-id="48574-598">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-599">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-600">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-600">1.0</span></span>|
|[<span data-ttu-id="48574-601">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-602">ReadItem</span></span>|
|[<span data-ttu-id="48574-603">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-604">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-604">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="48574-605">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="48574-605">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="48574-606">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-606">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-607">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-607">Read mode</span></span>

<span data-ttu-id="48574-608">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-608">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="48574-609">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-609">Compose mode</span></span>

<span data-ttu-id="48574-610">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-610">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-611">Type</span></span>

*   [<span data-ttu-id="48574-612">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="48574-612">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="48574-613">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-613">Requirements</span></span>

|<span data-ttu-id="48574-614">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-614">Requirement</span></span>|<span data-ttu-id="48574-615">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-616">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-617">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-617">1.8</span></span>|
|[<span data-ttu-id="48574-618">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-619">ReadItem</span></span>|
|[<span data-ttu-id="48574-620">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-621">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-622">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-622">Example</span></span>

<span data-ttu-id="48574-623">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-623">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="48574-624">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="48574-624">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="48574-625">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-625">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="48574-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="48574-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-628">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="48574-628">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-629">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-629">Read mode</span></span>

<span data-ttu-id="48574-630">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="48574-630">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-631">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-631">Compose mode</span></span>

<span data-ttu-id="48574-632">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="48574-632">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48574-633">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-633">Type</span></span>

*   <span data-ttu-id="48574-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="48574-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-635">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-635">Requirements</span></span>

|<span data-ttu-id="48574-636">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-636">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="48574-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-638">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-638">1.0</span></span>|<span data-ttu-id="48574-639">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-639">1.7</span></span>|
|[<span data-ttu-id="48574-640">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-640">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-641">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-641">ReadItem</span></span>|<span data-ttu-id="48574-642">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-642">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-643">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-643">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-644">Read</span><span class="sxs-lookup"><span data-stu-id="48574-644">Read</span></span>|<span data-ttu-id="48574-645">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-645">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="48574-646">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="48574-646">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="48574-647">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-647">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="48574-648">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="48574-648">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-649">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-649">Type</span></span>

*   [<span data-ttu-id="48574-650">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="48574-650">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="48574-651">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-651">Requirements</span></span>

|<span data-ttu-id="48574-652">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-652">Requirement</span></span>|<span data-ttu-id="48574-653">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-654">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-655">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-655">1.8</span></span>|
|[<span data-ttu-id="48574-656">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-657">ReadItem</span></span>|
|[<span data-ttu-id="48574-658">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-659">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-659">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-660">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-660">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="48574-661">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-661">internetMessageId: String</span></span>

<span data-ttu-id="48574-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-664">Type</span></span>

*   <span data-ttu-id="48574-665">String</span><span class="sxs-lookup"><span data-stu-id="48574-665">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-666">Requirements</span></span>

|<span data-ttu-id="48574-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-667">Requirement</span></span>|<span data-ttu-id="48574-668">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-670">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-670">1.0</span></span>|
|[<span data-ttu-id="48574-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-672">ReadItem</span></span>|
|[<span data-ttu-id="48574-673">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-674">Read</span><span class="sxs-lookup"><span data-stu-id="48574-674">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-675">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-675">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="48574-676">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="48574-676">itemClass: String</span></span>

<span data-ttu-id="48574-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="48574-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="48574-681">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-681">Type</span></span>|<span data-ttu-id="48574-682">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-682">Description</span></span>|<span data-ttu-id="48574-683">classe de item</span><span class="sxs-lookup"><span data-stu-id="48574-683">item class</span></span>|
|---|---|---|
|<span data-ttu-id="48574-684">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="48574-684">Appointment items</span></span>|<span data-ttu-id="48574-685">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="48574-685">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="48574-686">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="48574-686">Message items</span></span>|<span data-ttu-id="48574-687">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="48574-687">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="48574-688">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="48574-688">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-689">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-689">Type</span></span>

*   <span data-ttu-id="48574-690">String</span><span class="sxs-lookup"><span data-stu-id="48574-690">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-691">Requirements</span></span>

|<span data-ttu-id="48574-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-692">Requirement</span></span>|<span data-ttu-id="48574-693">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-695">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-695">1.0</span></span>|
|[<span data-ttu-id="48574-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-697">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-697">ReadItem</span></span>|
|[<span data-ttu-id="48574-698">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-699">Read</span><span class="sxs-lookup"><span data-stu-id="48574-699">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-700">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="48574-701">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-701">(nullable) itemId: String</span></span>

<span data-ttu-id="48574-p119">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-704">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="48574-704">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="48574-705">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="48574-705">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="48574-706">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="48574-706">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="48574-707">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="48574-707">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="48574-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-710">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-710">Type</span></span>

*   <span data-ttu-id="48574-711">String</span><span class="sxs-lookup"><span data-stu-id="48574-711">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-712">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-712">Requirements</span></span>

|<span data-ttu-id="48574-713">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-713">Requirement</span></span>|<span data-ttu-id="48574-714">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-715">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-716">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-716">1.0</span></span>|
|[<span data-ttu-id="48574-717">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-718">ReadItem</span></span>|
|[<span data-ttu-id="48574-719">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-720">Read</span><span class="sxs-lookup"><span data-stu-id="48574-720">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-721">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-721">Example</span></span>

<span data-ttu-id="48574-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="48574-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-mailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="48574-724">itemType: [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="48574-724">itemType: [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="48574-725">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="48574-725">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="48574-726">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-726">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-727">Type</span></span>

*   [<span data-ttu-id="48574-728">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="48574-728">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="48574-729">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-729">Requirements</span></span>

|<span data-ttu-id="48574-730">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-730">Requirement</span></span>|<span data-ttu-id="48574-731">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-731">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-732">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-732">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-733">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-733">1.0</span></span>|
|[<span data-ttu-id="48574-734">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-734">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-735">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-735">ReadItem</span></span>|
|[<span data-ttu-id="48574-736">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-736">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-737">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-737">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-738">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-738">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="48574-739">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="48574-739">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="48574-740">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-740">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-741">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-741">Read mode</span></span>

<span data-ttu-id="48574-742">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-742">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-743">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-743">Compose mode</span></span>

<span data-ttu-id="48574-744">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-744">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48574-745">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-745">Type</span></span>

*   <span data-ttu-id="48574-746">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="48574-746">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-747">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-747">Requirements</span></span>

|<span data-ttu-id="48574-748">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-748">Requirement</span></span>|<span data-ttu-id="48574-749">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-750">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-751">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-751">1.0</span></span>|
|[<span data-ttu-id="48574-752">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-752">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-753">ReadItem</span></span>|
|[<span data-ttu-id="48574-754">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-754">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-755">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-755">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="48574-756">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-756">normalizedSubject: String</span></span>

<span data-ttu-id="48574-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="48574-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="48574-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-761">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-761">Type</span></span>

*   <span data-ttu-id="48574-762">String</span><span class="sxs-lookup"><span data-stu-id="48574-762">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-763">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-763">Requirements</span></span>

|<span data-ttu-id="48574-764">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-764">Requirement</span></span>|<span data-ttu-id="48574-765">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-765">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-766">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-766">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-767">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-767">1.0</span></span>|
|[<span data-ttu-id="48574-768">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-768">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-769">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-769">ReadItem</span></span>|
|[<span data-ttu-id="48574-770">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-770">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-771">Read</span><span class="sxs-lookup"><span data-stu-id="48574-771">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-772">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-772">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="48574-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="48574-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="48574-774">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="48574-774">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-775">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-775">Type</span></span>

*   [<span data-ttu-id="48574-776">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="48574-776">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="48574-777">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-777">Requirements</span></span>

|<span data-ttu-id="48574-778">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-778">Requirement</span></span>|<span data-ttu-id="48574-779">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-780">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-781">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-781">1.3</span></span>|
|[<span data-ttu-id="48574-782">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-782">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-783">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-783">ReadItem</span></span>|
|[<span data-ttu-id="48574-784">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-784">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-785">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-785">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-786">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-786">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48574-787">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48574-788">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="48574-788">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="48574-789">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-789">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-790">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-790">Read mode</span></span>

<span data-ttu-id="48574-791">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-791">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="48574-792">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-792">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-793">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-793">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-794">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-794">Compose mode</span></span>

<span data-ttu-id="48574-795">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-795">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="48574-796">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-796">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-797">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="48574-797">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48574-798">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-798">Get 500 members maximum.</span></span>
- <span data-ttu-id="48574-799">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="48574-799">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48574-800">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-800">Type</span></span>

*   <span data-ttu-id="48574-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-802">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-802">Requirements</span></span>

|<span data-ttu-id="48574-803">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-803">Requirement</span></span>|<span data-ttu-id="48574-804">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-804">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-805">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-805">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-806">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-806">1.0</span></span>|
|[<span data-ttu-id="48574-807">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-807">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-808">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-808">ReadItem</span></span>|
|[<span data-ttu-id="48574-809">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-809">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-810">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-810">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="48574-811">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-811">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="48574-812">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="48574-812">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-813">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-813">Read mode</span></span>

<span data-ttu-id="48574-814">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-814">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-815">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-815">Compose mode</span></span>

<span data-ttu-id="48574-816">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="48574-816">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="48574-817">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-817">Type</span></span>

*   <span data-ttu-id="48574-818">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-819">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-819">Requirements</span></span>

|<span data-ttu-id="48574-820">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-820">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="48574-821">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-822">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-822">1.0</span></span>|<span data-ttu-id="48574-823">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-823">1.7</span></span>|
|[<span data-ttu-id="48574-824">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-824">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-825">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-825">ReadItem</span></span>|<span data-ttu-id="48574-826">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-826">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-827">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-827">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-828">Read</span><span class="sxs-lookup"><span data-stu-id="48574-828">Read</span></span>|<span data-ttu-id="48574-829">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-829">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="48574-830">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="48574-830">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="48574-831">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-831">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="48574-832">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-832">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="48574-833">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-833">Read and compose modes for appointment items.</span></span> <span data-ttu-id="48574-834">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-834">Read mode for meeting request items.</span></span>

<span data-ttu-id="48574-835">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="48574-835">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="48574-836">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="48574-836">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="48574-837">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-837">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="48574-838">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="48574-838">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="48574-839">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="48574-839">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-840">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-840">Read mode</span></span>

<span data-ttu-id="48574-841">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-841">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="48574-842">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-842">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-843">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-843">Compose mode</span></span>

<span data-ttu-id="48574-844">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-844">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="48574-845">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="48574-845">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="48574-846">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-846">Type</span></span>

* [<span data-ttu-id="48574-847">Recorrência</span><span class="sxs-lookup"><span data-stu-id="48574-847">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="48574-848">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-848">Requirement</span></span>|<span data-ttu-id="48574-849">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-850">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-851">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-851">1.7</span></span>|
|[<span data-ttu-id="48574-852">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-853">ReadItem</span></span>|
|[<span data-ttu-id="48574-854">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-855">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-855">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48574-856">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48574-857">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="48574-857">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="48574-858">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-858">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-859">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-859">Read mode</span></span>

<span data-ttu-id="48574-860">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-860">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="48574-861">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-861">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-862">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-862">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-863">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-863">Compose mode</span></span>

<span data-ttu-id="48574-864">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-864">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="48574-865">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-865">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-866">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="48574-866">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48574-867">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-867">Get 500 members maximum.</span></span>
- <span data-ttu-id="48574-868">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="48574-868">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="48574-869">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-869">Type</span></span>

*   <span data-ttu-id="48574-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-871">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-871">Requirements</span></span>

|<span data-ttu-id="48574-872">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-872">Requirement</span></span>|<span data-ttu-id="48574-873">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-874">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-875">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-875">1.0</span></span>|
|[<span data-ttu-id="48574-876">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-876">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-877">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-877">ReadItem</span></span>|
|[<span data-ttu-id="48574-878">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-878">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-879">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-879">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="48574-880">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="48574-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="48574-p135">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="48574-p136">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="48574-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-885">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="48574-885">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-886">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-886">Type</span></span>

*   [<span data-ttu-id="48574-887">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48574-887">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="48574-888">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-888">Requirements</span></span>

|<span data-ttu-id="48574-889">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-889">Requirement</span></span>|<span data-ttu-id="48574-890">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-891">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-892">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-892">1.0</span></span>|
|[<span data-ttu-id="48574-893">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-894">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-894">ReadItem</span></span>|
|[<span data-ttu-id="48574-895">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-896">Read</span><span class="sxs-lookup"><span data-stu-id="48574-896">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-897">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-897">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="48574-898">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="48574-898">(nullable) seriesId: String</span></span>

<span data-ttu-id="48574-899">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="48574-899">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="48574-900">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="48574-900">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="48574-901">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="48574-901">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-902">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="48574-902">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="48574-903">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="48574-903">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="48574-904">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="48574-904">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="48574-905">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="48574-905">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="48574-906">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="48574-906">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="48574-907">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-907">Type</span></span>

* <span data-ttu-id="48574-908">String</span><span class="sxs-lookup"><span data-stu-id="48574-908">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-909">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-909">Requirements</span></span>

|<span data-ttu-id="48574-910">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-910">Requirement</span></span>|<span data-ttu-id="48574-911">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-911">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-912">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-912">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-913">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-913">1.7</span></span>|
|[<span data-ttu-id="48574-914">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-914">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-915">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-915">ReadItem</span></span>|
|[<span data-ttu-id="48574-916">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-916">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-917">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-917">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-918">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-918">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="48574-919">início: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48574-919">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="48574-920">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="48574-920">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="48574-p139">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="48574-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-923">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-923">Read mode</span></span>

<span data-ttu-id="48574-924">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="48574-924">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-925">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-925">Compose mode</span></span>

<span data-ttu-id="48574-926">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="48574-926">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="48574-927">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-927">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="48574-928">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="48574-928">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="48574-929">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-929">Type</span></span>

*   <span data-ttu-id="48574-930">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48574-930">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-931">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-931">Requirements</span></span>

|<span data-ttu-id="48574-932">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-932">Requirement</span></span>|<span data-ttu-id="48574-933">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-934">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-935">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-935">1.0</span></span>|
|[<span data-ttu-id="48574-936">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-936">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-937">ReadItem</span></span>|
|[<span data-ttu-id="48574-938">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-938">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-939">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-939">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="48574-940">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="48574-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="48574-941">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="48574-941">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="48574-942">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="48574-942">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-943">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-943">Read mode</span></span>

<span data-ttu-id="48574-p140">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="48574-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="48574-946">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="48574-946">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-947">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-947">Compose mode</span></span>
<span data-ttu-id="48574-948">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="48574-948">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="48574-949">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-949">Type</span></span>

*   <span data-ttu-id="48574-950">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="48574-950">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-951">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-951">Requirements</span></span>

|<span data-ttu-id="48574-952">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-952">Requirement</span></span>|<span data-ttu-id="48574-953">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-953">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-954">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-954">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-955">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-955">1.0</span></span>|
|[<span data-ttu-id="48574-956">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-956">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-957">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-957">ReadItem</span></span>|
|[<span data-ttu-id="48574-958">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-958">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-959">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-959">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48574-960">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48574-961">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-961">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="48574-962">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-962">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48574-963">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="48574-963">Read mode</span></span>

<span data-ttu-id="48574-964">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-964">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="48574-965">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-965">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-966">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-966">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="48574-967">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="48574-967">Compose mode</span></span>

<span data-ttu-id="48574-968">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-968">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="48574-969">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-969">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48574-970">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="48574-970">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48574-971">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="48574-971">Get 500 members maximum.</span></span>
- <span data-ttu-id="48574-972">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="48574-972">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48574-973">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-973">Type</span></span>

*   <span data-ttu-id="48574-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48574-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-975">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-975">Requirements</span></span>

|<span data-ttu-id="48574-976">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-976">Requirement</span></span>|<span data-ttu-id="48574-977">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-977">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-978">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-978">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-979">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-979">1.0</span></span>|
|[<span data-ttu-id="48574-980">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-980">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-981">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-981">ReadItem</span></span>|
|[<span data-ttu-id="48574-982">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-982">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-983">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-983">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="48574-984">Detalhes do método</span><span class="sxs-lookup"><span data-stu-id="48574-984">Method details</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="48574-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48574-986">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="48574-986">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="48574-987">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="48574-987">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="48574-988">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-988">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-989">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-989">Parameters</span></span>
|<span data-ttu-id="48574-990">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-990">Name</span></span>|<span data-ttu-id="48574-991">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-991">Type</span></span>|<span data-ttu-id="48574-992">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-992">Attributes</span></span>|<span data-ttu-id="48574-993">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-993">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="48574-994">String</span><span class="sxs-lookup"><span data-stu-id="48574-994">String</span></span>||<span data-ttu-id="48574-p144">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="48574-997">String</span><span class="sxs-lookup"><span data-stu-id="48574-997">String</span></span>||<span data-ttu-id="48574-p145">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48574-1000">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1000">Object</span></span>|<span data-ttu-id="48574-1001">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1002">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1003">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1003">Object</span></span>|<span data-ttu-id="48574-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1005">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="48574-1006">Booliano</span><span class="sxs-lookup"><span data-stu-id="48574-1006">Boolean</span></span>|<span data-ttu-id="48574-1007">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1008">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-1008">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="48574-1009">function</span><span class="sxs-lookup"><span data-stu-id="48574-1009">function</span></span>|<span data-ttu-id="48574-1010">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1011">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1011">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48574-1012">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1012">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48574-1013">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="48574-1013">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48574-1014">Erros</span><span class="sxs-lookup"><span data-stu-id="48574-1014">Errors</span></span>

|<span data-ttu-id="48574-1015">Código de erro</span><span class="sxs-lookup"><span data-stu-id="48574-1015">Error code</span></span>|<span data-ttu-id="48574-1016">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1016">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="48574-1017">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="48574-1017">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="48574-1018">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="48574-1018">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48574-1019">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-1019">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1020">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1020">Requirements</span></span>

|<span data-ttu-id="48574-1021">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1021">Requirement</span></span>|<span data-ttu-id="48574-1022">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1023">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1024">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-1024">1.1</span></span>|
|[<span data-ttu-id="48574-1025">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1026">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1026">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1027">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1028">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1028">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1029">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1029">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="48574-1030">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-1030">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="48574-1031">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1031">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48574-1032">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="48574-1032">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="48574-1033">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="48574-1033">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="48574-1034">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="48574-1034">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="48574-1035">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-1035">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1036">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1036">Parameters</span></span>

|<span data-ttu-id="48574-1037">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1037">Name</span></span>|<span data-ttu-id="48574-1038">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1038">Type</span></span>|<span data-ttu-id="48574-1039">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1039">Attributes</span></span>|<span data-ttu-id="48574-1040">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1040">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="48574-1041">String</span><span class="sxs-lookup"><span data-stu-id="48574-1041">String</span></span>||<span data-ttu-id="48574-1042">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="48574-1042">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="48574-1043">String</span><span class="sxs-lookup"><span data-stu-id="48574-1043">String</span></span>||<span data-ttu-id="48574-p147">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48574-1046">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1046">Object</span></span>|<span data-ttu-id="48574-1047">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1048">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1049">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1049">Object</span></span>|<span data-ttu-id="48574-1050">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1051">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="48574-1052">Booliano</span><span class="sxs-lookup"><span data-stu-id="48574-1052">Boolean</span></span>|<span data-ttu-id="48574-1053">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1054">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-1054">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="48574-1055">function</span><span class="sxs-lookup"><span data-stu-id="48574-1055">function</span></span>|<span data-ttu-id="48574-1056">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1057">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48574-1058">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1058">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48574-1059">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="48574-1059">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48574-1060">Erros</span><span class="sxs-lookup"><span data-stu-id="48574-1060">Errors</span></span>

|<span data-ttu-id="48574-1061">Código de erro</span><span class="sxs-lookup"><span data-stu-id="48574-1061">Error code</span></span>|<span data-ttu-id="48574-1062">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1062">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="48574-1063">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="48574-1063">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="48574-1064">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="48574-1064">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48574-1065">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-1065">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1066">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1066">Requirements</span></span>

|<span data-ttu-id="48574-1067">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1067">Requirement</span></span>|<span data-ttu-id="48574-1068">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1069">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1070">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1070">1.8</span></span>|
|[<span data-ttu-id="48574-1071">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1072">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1072">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1073">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1074">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1074">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1075">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1075">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="48574-1076">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1076">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="48574-1077">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="48574-1077">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="48574-1078">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="48574-1078">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1079">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1079">Parameters</span></span>

| <span data-ttu-id="48574-1080">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1080">Name</span></span> | <span data-ttu-id="48574-1081">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1081">Type</span></span> | <span data-ttu-id="48574-1082">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1082">Attributes</span></span> | <span data-ttu-id="48574-1083">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1083">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48574-1084">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48574-1084">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48574-1085">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="48574-1085">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="48574-1086">Função</span><span class="sxs-lookup"><span data-stu-id="48574-1086">Function</span></span> || <span data-ttu-id="48574-p148">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="48574-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="48574-1090">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1090">Object</span></span> | <span data-ttu-id="48574-1091">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1091">&lt;optional&gt;</span></span> | <span data-ttu-id="48574-1092">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1092">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48574-1093">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1093">Object</span></span> | <span data-ttu-id="48574-1094">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1094">&lt;optional&gt;</span></span> | <span data-ttu-id="48574-1095">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1095">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48574-1096">function</span><span class="sxs-lookup"><span data-stu-id="48574-1096">function</span></span>| <span data-ttu-id="48574-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1098">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1099">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1099">Requirements</span></span>

|<span data-ttu-id="48574-1100">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1100">Requirement</span></span>| <span data-ttu-id="48574-1101">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1101">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1102">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1102">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48574-1103">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-1103">1.7</span></span> |
|[<span data-ttu-id="48574-1104">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1104">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48574-1105">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1105">ReadItem</span></span> |
|[<span data-ttu-id="48574-1106">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-1106">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48574-1107">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-1107">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="48574-1108">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1108">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="48574-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48574-1110">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-1110">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="48574-p149">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="48574-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="48574-1114">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-1114">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="48574-1115">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="48574-1115">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1116">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1116">Parameters</span></span>

|<span data-ttu-id="48574-1117">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1117">Name</span></span>|<span data-ttu-id="48574-1118">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1118">Type</span></span>|<span data-ttu-id="48574-1119">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1119">Attributes</span></span>|<span data-ttu-id="48574-1120">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1120">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="48574-1121">String</span><span class="sxs-lookup"><span data-stu-id="48574-1121">String</span></span>||<span data-ttu-id="48574-p150">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="48574-1124">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-1124">String</span></span>||<span data-ttu-id="48574-1125">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="48574-1125">The subject of the item to be attached.</span></span> <span data-ttu-id="48574-1126">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-1126">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48574-1127">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1127">Object</span></span>|<span data-ttu-id="48574-1128">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1129">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1129">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1130">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1130">Object</span></span>|<span data-ttu-id="48574-1131">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1131">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1132">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1132">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1133">function</span><span class="sxs-lookup"><span data-stu-id="48574-1133">function</span></span>|<span data-ttu-id="48574-1134">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1134">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1135">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48574-1136">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1136">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48574-1137">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="48574-1137">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48574-1138">Erros</span><span class="sxs-lookup"><span data-stu-id="48574-1138">Errors</span></span>

|<span data-ttu-id="48574-1139">Código de erro</span><span class="sxs-lookup"><span data-stu-id="48574-1139">Error code</span></span>|<span data-ttu-id="48574-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1140">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48574-1141">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-1141">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1142">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1142">Requirements</span></span>

|<span data-ttu-id="48574-1143">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1143">Requirement</span></span>|<span data-ttu-id="48574-1144">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1144">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1145">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1146">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-1146">1.1</span></span>|
|[<span data-ttu-id="48574-1147">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1148">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1148">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1149">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1150">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1151">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1151">Example</span></span>

<span data-ttu-id="48574-1152">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="48574-1152">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a><span data-ttu-id="48574-1153">close()</span><span class="sxs-lookup"><span data-stu-id="48574-1153">close()</span></span>

<span data-ttu-id="48574-1154">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="48574-1154">Closes the current item that is being composed.</span></span>

<span data-ttu-id="48574-p152">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="48574-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1157">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="48574-1157">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="48574-1158">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="48574-1158">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-1159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1159">Requirements</span></span>

|<span data-ttu-id="48574-1160">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1160">Requirement</span></span>|<span data-ttu-id="48574-1161">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1163">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-1163">1.3</span></span>|
|[<span data-ttu-id="48574-1164">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1165">Restrito</span><span class="sxs-lookup"><span data-stu-id="48574-1165">Restricted</span></span>|
|[<span data-ttu-id="48574-1166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1167">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1167">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="48574-1168">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1168">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="48574-1169">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="48574-1169">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1170">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1170">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-1171">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="48574-1171">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="48574-1172">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="48574-1172">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="48574-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="48574-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1176">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1176">Parameters</span></span>

|<span data-ttu-id="48574-1177">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1177">Name</span></span>|<span data-ttu-id="48574-1178">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1178">Type</span></span>|<span data-ttu-id="48574-1179">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1179">Attributes</span></span>|<span data-ttu-id="48574-1180">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1180">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="48574-1181">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="48574-1181">String &#124; Object</span></span>||<span data-ttu-id="48574-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="48574-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="48574-1184">**OU**</span><span class="sxs-lookup"><span data-stu-id="48574-1184">**OR**</span></span><br/><span data-ttu-id="48574-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="48574-1187">String</span><span class="sxs-lookup"><span data-stu-id="48574-1187">String</span></span>|<span data-ttu-id="48574-1188">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1188">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="48574-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="48574-1191">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1191">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="48574-1192">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1192">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1193">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="48574-1193">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="48574-1194">String</span><span class="sxs-lookup"><span data-stu-id="48574-1194">String</span></span>||<span data-ttu-id="48574-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="48574-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="48574-1197">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-1197">String</span></span>||<span data-ttu-id="48574-1198">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="48574-1198">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="48574-1199">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-1199">String</span></span>||<span data-ttu-id="48574-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="48574-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="48574-1202">Booliano</span><span class="sxs-lookup"><span data-stu-id="48574-1202">Boolean</span></span>||<span data-ttu-id="48574-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="48574-1205">String</span><span class="sxs-lookup"><span data-stu-id="48574-1205">String</span></span>||<span data-ttu-id="48574-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="48574-1209">function</span><span class="sxs-lookup"><span data-stu-id="48574-1209">function</span></span>|<span data-ttu-id="48574-1210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1210">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1211">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1212">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1212">Requirements</span></span>

|<span data-ttu-id="48574-1213">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1213">Requirement</span></span>|<span data-ttu-id="48574-1214">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1215">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1216">1.0</span></span>|
|[<span data-ttu-id="48574-1217">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1218">ReadItem</span></span>|
|[<span data-ttu-id="48574-1219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1220">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1220">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1221">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1221">Examples</span></span>

<span data-ttu-id="48574-1222">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="48574-1222">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="48574-1223">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="48574-1223">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="48574-1224">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="48574-1224">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="48574-1225">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="48574-1225">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="48574-1226">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="48574-1226">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="48574-1227">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1227">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="48574-1228">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1228">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="48574-1229">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="48574-1229">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1230">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-1231">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="48574-1231">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="48574-1232">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="48574-1232">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="48574-p161">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="48574-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1236">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1236">Parameters</span></span>

|<span data-ttu-id="48574-1237">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1237">Name</span></span>|<span data-ttu-id="48574-1238">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1238">Type</span></span>|<span data-ttu-id="48574-1239">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1239">Attributes</span></span>|<span data-ttu-id="48574-1240">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1240">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="48574-1241">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="48574-1241">String &#124; Object</span></span>||<span data-ttu-id="48574-p162">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="48574-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="48574-1244">**OU**</span><span class="sxs-lookup"><span data-stu-id="48574-1244">**OR**</span></span><br/><span data-ttu-id="48574-p163">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="48574-1247">String</span><span class="sxs-lookup"><span data-stu-id="48574-1247">String</span></span>|<span data-ttu-id="48574-1248">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1248">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-p164">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="48574-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="48574-1251">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1251">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="48574-1252">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1252">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1253">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="48574-1253">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="48574-1254">String</span><span class="sxs-lookup"><span data-stu-id="48574-1254">String</span></span>||<span data-ttu-id="48574-p165">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="48574-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="48574-1257">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-1257">String</span></span>||<span data-ttu-id="48574-1258">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="48574-1258">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="48574-1259">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="48574-1259">String</span></span>||<span data-ttu-id="48574-p166">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="48574-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="48574-1262">Booliano</span><span class="sxs-lookup"><span data-stu-id="48574-1262">Boolean</span></span>||<span data-ttu-id="48574-p167">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="48574-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="48574-1265">String</span><span class="sxs-lookup"><span data-stu-id="48574-1265">String</span></span>||<span data-ttu-id="48574-p168">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="48574-1269">function</span><span class="sxs-lookup"><span data-stu-id="48574-1269">function</span></span>|<span data-ttu-id="48574-1270">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1271">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1272">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1272">Requirements</span></span>

|<span data-ttu-id="48574-1273">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1273">Requirement</span></span>|<span data-ttu-id="48574-1274">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1274">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1275">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1275">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1276">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1276">1.0</span></span>|
|[<span data-ttu-id="48574-1277">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1277">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1278">ReadItem</span></span>|
|[<span data-ttu-id="48574-1279">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1279">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1280">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1280">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1281">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1281">Examples</span></span>

<span data-ttu-id="48574-1282">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="48574-1282">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="48574-1283">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="48574-1283">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="48574-1284">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="48574-1284">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="48574-1285">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="48574-1285">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="48574-1286">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="48574-1286">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="48574-1287">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1287">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="48574-1288">getAllInternetHeadersAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1288">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="48574-1289">Obtém todos os cabeçalhos de Internet da mensagem como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-1289">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="48574-1290">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="48574-1290">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1291">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1291">Parameters</span></span>

|<span data-ttu-id="48574-1292">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1292">Name</span></span>|<span data-ttu-id="48574-1293">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1293">Type</span></span>|<span data-ttu-id="48574-1294">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1294">Attributes</span></span>|<span data-ttu-id="48574-1295">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1295">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1296">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1296">Object</span></span>|<span data-ttu-id="48574-1297">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1297">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1298">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1298">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1299">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1299">Object</span></span>|<span data-ttu-id="48574-1300">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1301">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1301">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1302">function</span><span class="sxs-lookup"><span data-stu-id="48574-1302">function</span></span>|<span data-ttu-id="48574-1303">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1304">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1304">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="48574-1305">Com êxito, os dados de cabeçalhos de Internet são fornecidos na propriedade asyncResult. Value como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-1305">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="48574-1306">Consulte [RFC 2183](https://tools.ietf.org/html/rfc2183) para obter as informações de formatação do valor de cadeia de caracteres retornado.</span><span class="sxs-lookup"><span data-stu-id="48574-1306">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="48574-1307">Se a chamada falhar, a propriedade asyncResult. Error conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="48574-1307">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1308">Requirements</span></span>

|<span data-ttu-id="48574-1309">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1309">Requirement</span></span>|<span data-ttu-id="48574-1310">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1310">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1312">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1312">1.8</span></span>|
|[<span data-ttu-id="48574-1313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1314">ReadItem</span></span>|
|[<span data-ttu-id="48574-1315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1316">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1316">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1317">Returns:</span></span>

<span data-ttu-id="48574-1318">A Internet cabeçalhos dados como uma cadeia de caracteres formatada de acordo com a [RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="48574-1318">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="48574-1319">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="48574-1319">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1320">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1320">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="48574-1321">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="48574-1321">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="48574-1322">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="48574-1322">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="48574-1323">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="48574-1323">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="48574-1324">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="48574-1324">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="48574-1325">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-1325">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="48574-1326">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="48574-1326">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1327">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1327">Parameters</span></span>

|<span data-ttu-id="48574-1328">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1328">Name</span></span>|<span data-ttu-id="48574-1329">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1329">Type</span></span>|<span data-ttu-id="48574-1330">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1330">Attributes</span></span>|<span data-ttu-id="48574-1331">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1331">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="48574-1332">String</span><span class="sxs-lookup"><span data-stu-id="48574-1332">String</span></span>||<span data-ttu-id="48574-1333">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="48574-1333">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="48574-1334">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1334">Object</span></span>|<span data-ttu-id="48574-1335">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1335">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1336">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1336">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1337">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1337">Object</span></span>|<span data-ttu-id="48574-1338">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1338">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1339">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1339">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1340">function</span><span class="sxs-lookup"><span data-stu-id="48574-1340">function</span></span>|<span data-ttu-id="48574-1341">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1342">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1343">Requirements</span></span>

|<span data-ttu-id="48574-1344">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1344">Requirement</span></span>|<span data-ttu-id="48574-1345">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1347">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1347">1.8</span></span>|
|[<span data-ttu-id="48574-1348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1349">ReadItem</span></span>|
|[<span data-ttu-id="48574-1350">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-1350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1351">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-1351">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1352">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1352">Returns:</span></span>

<span data-ttu-id="48574-1353">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="48574-1353">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1354">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1354">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="48574-1355">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48574-1355">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="48574-1356">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="48574-1356">Gets the item's attachments as an array.</span></span> <span data-ttu-id="48574-1357">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="48574-1357">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1358">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1358">Parameters</span></span>

|<span data-ttu-id="48574-1359">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1359">Name</span></span>|<span data-ttu-id="48574-1360">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1360">Type</span></span>|<span data-ttu-id="48574-1361">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1361">Attributes</span></span>|<span data-ttu-id="48574-1362">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1362">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1363">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1363">Object</span></span>|<span data-ttu-id="48574-1364">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1364">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1365">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1365">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1366">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1366">Object</span></span>|<span data-ttu-id="48574-1367">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1367">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1368">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1368">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1369">function</span><span class="sxs-lookup"><span data-stu-id="48574-1369">function</span></span>|<span data-ttu-id="48574-1370">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1371">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1371">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1372">Requirements</span></span>

|<span data-ttu-id="48574-1373">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1373">Requirement</span></span>|<span data-ttu-id="48574-1374">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1374">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1376">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1376">1.8</span></span>|
|[<span data-ttu-id="48574-1377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1378">ReadItem</span></span>|
|[<span data-ttu-id="48574-1379">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1380">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1380">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1381">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1381">Returns:</span></span>

<span data-ttu-id="48574-1382">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48574-1382">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="48574-1383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1383">Example</span></span>

<span data-ttu-id="48574-1384">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-1384">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="48574-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="48574-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="48574-1386">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="48574-1386">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1387">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1387">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-1388">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1388">Requirements</span></span>

|<span data-ttu-id="48574-1389">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1389">Requirement</span></span>|<span data-ttu-id="48574-1390">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1390">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1391">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1392">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1392">1.0</span></span>|
|[<span data-ttu-id="48574-1393">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1393">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1394">ReadItem</span></span>|
|[<span data-ttu-id="48574-1395">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1395">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1396">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1396">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1397">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1397">Returns:</span></span>

<span data-ttu-id="48574-1398">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="48574-1398">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1399">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1399">Example</span></span>

<span data-ttu-id="48574-1400">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-1400">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="48574-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="48574-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="48574-1402">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="48574-1402">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1403">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1403">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1404">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1404">Parameters</span></span>

|<span data-ttu-id="48574-1405">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1405">Name</span></span>|<span data-ttu-id="48574-1406">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1406">Type</span></span>|<span data-ttu-id="48574-1407">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1407">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="48574-1408">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="48574-1408">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="48574-1409">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="48574-1409">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1410">Requirements</span></span>

|<span data-ttu-id="48574-1411">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1411">Requirement</span></span>|<span data-ttu-id="48574-1412">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1412">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1414">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1414">1.0</span></span>|
|[<span data-ttu-id="48574-1415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1416">Restrito</span><span class="sxs-lookup"><span data-stu-id="48574-1416">Restricted</span></span>|
|[<span data-ttu-id="48574-1417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1418">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1418">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1419">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1419">Returns:</span></span>

<span data-ttu-id="48574-1420">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="48574-1420">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="48574-1421">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="48574-1421">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="48574-1422">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="48574-1422">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="48574-1423">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1423">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="48574-1424">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="48574-1424">Value of `entityType`</span></span>|<span data-ttu-id="48574-1425">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="48574-1425">Type of objects in returned array</span></span>|<span data-ttu-id="48574-1426">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="48574-1426">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="48574-1427">String</span><span class="sxs-lookup"><span data-stu-id="48574-1427">String</span></span>|<span data-ttu-id="48574-1428">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="48574-1428">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="48574-1429">Contato</span><span class="sxs-lookup"><span data-stu-id="48574-1429">Contact</span></span>|<span data-ttu-id="48574-1430">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48574-1430">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="48574-1431">String</span><span class="sxs-lookup"><span data-stu-id="48574-1431">String</span></span>|<span data-ttu-id="48574-1432">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48574-1432">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="48574-1433">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="48574-1433">MeetingSuggestion</span></span>|<span data-ttu-id="48574-1434">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48574-1434">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="48574-1435">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="48574-1435">PhoneNumber</span></span>|<span data-ttu-id="48574-1436">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="48574-1436">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="48574-1437">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="48574-1437">TaskSuggestion</span></span>|<span data-ttu-id="48574-1438">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48574-1438">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="48574-1439">String</span><span class="sxs-lookup"><span data-stu-id="48574-1439">String</span></span>|<span data-ttu-id="48574-1440">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="48574-1440">**Restricted**</span></span>|

<span data-ttu-id="48574-1441">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="48574-1441">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="48574-1442">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1442">Example</span></span>

<span data-ttu-id="48574-1443">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="48574-1443">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="48574-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="48574-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="48574-1445">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="48574-1445">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1446">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1446">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-1447">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="48574-1447">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1448">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1448">Parameters</span></span>

|<span data-ttu-id="48574-1449">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1449">Name</span></span>|<span data-ttu-id="48574-1450">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1450">Type</span></span>|<span data-ttu-id="48574-1451">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1451">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="48574-1452">String</span><span class="sxs-lookup"><span data-stu-id="48574-1452">String</span></span>|<span data-ttu-id="48574-1453">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="48574-1453">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1454">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1454">Requirements</span></span>

|<span data-ttu-id="48574-1455">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1455">Requirement</span></span>|<span data-ttu-id="48574-1456">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1456">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1457">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1457">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1458">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1458">1.0</span></span>|
|[<span data-ttu-id="48574-1459">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1459">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1460">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1460">ReadItem</span></span>|
|[<span data-ttu-id="48574-1461">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1461">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1462">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1462">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1463">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1463">Returns:</span></span>

<span data-ttu-id="48574-p174">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="48574-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="48574-1466">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="48574-1466">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="48574-1467">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1467">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="48574-1468">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="48574-1468">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1469">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="48574-1469">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1470">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1470">Parameters</span></span>

|<span data-ttu-id="48574-1471">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1471">Name</span></span>|<span data-ttu-id="48574-1472">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1472">Type</span></span>|<span data-ttu-id="48574-1473">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1473">Attributes</span></span>|<span data-ttu-id="48574-1474">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1474">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1475">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1475">Object</span></span>|<span data-ttu-id="48574-1476">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1476">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1477">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1477">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1478">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1478">Object</span></span>|<span data-ttu-id="48574-1479">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1479">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1480">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1480">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1481">function</span><span class="sxs-lookup"><span data-stu-id="48574-1481">function</span></span>|<span data-ttu-id="48574-1482">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1482">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1483">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1483">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48574-1484">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="48574-1484">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="48574-1485">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="48574-1485">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1486">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1486">Requirements</span></span>

|<span data-ttu-id="48574-1487">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1487">Requirement</span></span>|<span data-ttu-id="48574-1488">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1489">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1490">Visualização</span><span class="sxs-lookup"><span data-stu-id="48574-1490">Preview</span></span>|
|[<span data-ttu-id="48574-1491">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1492">ReadItem</span></span>|
|[<span data-ttu-id="48574-1493">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1494">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1494">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1495">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1495">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="48574-1496">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="48574-1496">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="48574-1497">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="48574-1497">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="48574-1498">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="48574-1498">Compose mode only.</span></span>

<span data-ttu-id="48574-1499">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1499">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1500">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-1500">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="48574-1501">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="48574-1501">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1502">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1502">Parameters</span></span>

|<span data-ttu-id="48574-1503">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1503">Name</span></span>|<span data-ttu-id="48574-1504">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1504">Type</span></span>|<span data-ttu-id="48574-1505">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1505">Attributes</span></span>|<span data-ttu-id="48574-1506">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1506">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1507">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1507">Object</span></span>|<span data-ttu-id="48574-1508">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1508">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1509">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1509">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1510">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1510">Object</span></span>|<span data-ttu-id="48574-1511">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1511">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1512">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1512">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1513">function</span><span class="sxs-lookup"><span data-stu-id="48574-1513">function</span></span>||<span data-ttu-id="48574-1514">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1514">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48574-1515">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1515">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48574-1516">Erros</span><span class="sxs-lookup"><span data-stu-id="48574-1516">Errors</span></span>

|<span data-ttu-id="48574-1517">Código de erro</span><span class="sxs-lookup"><span data-stu-id="48574-1517">Error code</span></span>|<span data-ttu-id="48574-1518">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1518">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="48574-1519">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="48574-1519">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1520">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1520">Requirements</span></span>

|<span data-ttu-id="48574-1521">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1521">Requirement</span></span>|<span data-ttu-id="48574-1522">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1522">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1523">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1524">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1524">1.8</span></span>|
|[<span data-ttu-id="48574-1525">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1526">ReadItem</span></span>|
|[<span data-ttu-id="48574-1527">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1528">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1528">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1529">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1529">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="48574-1530">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1530">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="48574-1531">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="48574-1531">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="48574-1532">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="48574-1532">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="48574-1533">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="48574-1533">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1534">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1534">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-p178">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="48574-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="48574-1538">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="48574-1538">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="48574-1539">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="48574-1539">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="48574-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="48574-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-1543">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1543">Requirements</span></span>

|<span data-ttu-id="48574-1544">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1544">Requirement</span></span>|<span data-ttu-id="48574-1545">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1547">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1547">1.0</span></span>|
|[<span data-ttu-id="48574-1548">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1549">ReadItem</span></span>|
|[<span data-ttu-id="48574-1550">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1551">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1551">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1552">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1552">Returns:</span></span>

<span data-ttu-id="48574-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="48574-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="48574-1555">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="48574-1555">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="48574-1556">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1556">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="48574-1557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1557">Example</span></span>

<span data-ttu-id="48574-1558">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="48574-1558">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="48574-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="48574-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="48574-1560">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="48574-1560">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1561">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1561">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-1562">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="48574-1562">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="48574-p181">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="48574-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1565">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1565">Parameters</span></span>

|<span data-ttu-id="48574-1566">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1566">Name</span></span>|<span data-ttu-id="48574-1567">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1567">Type</span></span>|<span data-ttu-id="48574-1568">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1568">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="48574-1569">String</span><span class="sxs-lookup"><span data-stu-id="48574-1569">String</span></span>|<span data-ttu-id="48574-1570">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="48574-1570">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1571">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1571">Requirements</span></span>

|<span data-ttu-id="48574-1572">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1572">Requirement</span></span>|<span data-ttu-id="48574-1573">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1573">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1574">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1575">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1575">1.0</span></span>|
|[<span data-ttu-id="48574-1576">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1577">ReadItem</span></span>|
|[<span data-ttu-id="48574-1578">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1579">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1579">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1580">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1580">Returns:</span></span>

<span data-ttu-id="48574-1581">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="48574-1581">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="48574-1582">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="48574-1582">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="48574-1583">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1583">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="48574-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="48574-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="48574-1585">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-1585">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="48574-p182">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna uma cadeia de caracteres vazia para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="48574-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1588">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1588">Parameters</span></span>

|<span data-ttu-id="48574-1589">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1589">Name</span></span>|<span data-ttu-id="48574-1590">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1590">Type</span></span>|<span data-ttu-id="48574-1591">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1591">Attributes</span></span>|<span data-ttu-id="48574-1592">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1592">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="48574-1593">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="48574-1593">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="48574-p183">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="48574-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="48574-1597">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1597">Object</span></span>|<span data-ttu-id="48574-1598">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1599">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1599">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1600">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1600">Object</span></span>|<span data-ttu-id="48574-1601">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1602">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1602">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1603">function</span><span class="sxs-lookup"><span data-stu-id="48574-1603">function</span></span>||<span data-ttu-id="48574-1604">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1604">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48574-1605">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="48574-1605">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="48574-1606">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="48574-1606">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1607">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1607">Requirements</span></span>

|<span data-ttu-id="48574-1608">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1608">Requirement</span></span>|<span data-ttu-id="48574-1609">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1609">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1610">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1611">1.2</span><span class="sxs-lookup"><span data-stu-id="48574-1611">1.2</span></span>|
|[<span data-ttu-id="48574-1612">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1613">ReadItem</span></span>|
|[<span data-ttu-id="48574-1614">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1615">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1615">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1616">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1616">Returns:</span></span>

<span data-ttu-id="48574-1617">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="48574-1617">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="48574-1618">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="48574-1618">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1619">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1619">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="48574-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="48574-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="48574-1621">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="48574-1621">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="48574-1622">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="48574-1622">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1623">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1623">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-1624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1624">Requirements</span></span>

|<span data-ttu-id="48574-1625">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1625">Requirement</span></span>|<span data-ttu-id="48574-1626">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1626">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1627">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1628">1.6</span><span class="sxs-lookup"><span data-stu-id="48574-1628">1.6</span></span>|
|[<span data-ttu-id="48574-1629">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1630">ReadItem</span></span>|
|[<span data-ttu-id="48574-1631">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1632">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1632">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1633">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1633">Returns:</span></span>

<span data-ttu-id="48574-1634">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="48574-1634">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1635">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1635">Example</span></span>

<span data-ttu-id="48574-1636">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="48574-1636">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="48574-1637">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="48574-1637">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="48574-p186">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="48574-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1640">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="48574-1640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48574-p187">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="48574-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="48574-1644">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="48574-1644">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="48574-1645">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="48574-1645">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="48574-p188">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="48574-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48574-1649">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1649">Requirements</span></span>

|<span data-ttu-id="48574-1650">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1650">Requirement</span></span>|<span data-ttu-id="48574-1651">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1651">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1652">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1653">1.6</span><span class="sxs-lookup"><span data-stu-id="48574-1653">1.6</span></span>|
|[<span data-ttu-id="48574-1654">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1655">ReadItem</span></span>|
|[<span data-ttu-id="48574-1656">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1657">Read</span><span class="sxs-lookup"><span data-stu-id="48574-1657">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48574-1658">Retorna:</span><span class="sxs-lookup"><span data-stu-id="48574-1658">Returns:</span></span>

<span data-ttu-id="48574-p189">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="48574-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="48574-1661">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1661">Example</span></span>

<span data-ttu-id="48574-1662">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="48574-1662">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="48574-1663">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="48574-1663">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="48574-1664">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="48574-1664">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1665">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1665">Parameters</span></span>

|<span data-ttu-id="48574-1666">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1666">Name</span></span>|<span data-ttu-id="48574-1667">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1667">Type</span></span>|<span data-ttu-id="48574-1668">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1668">Attributes</span></span>|<span data-ttu-id="48574-1669">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1669">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1670">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1670">Object</span></span>|<span data-ttu-id="48574-1671">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1671">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1672">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1672">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1673">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1673">Object</span></span>|<span data-ttu-id="48574-1674">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1674">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1675">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1675">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1676">function</span><span class="sxs-lookup"><span data-stu-id="48574-1676">function</span></span>||<span data-ttu-id="48574-1677">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48574-1678">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="48574-1678">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="48574-1679">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="48574-1679">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1680">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1680">Requirements</span></span>

|<span data-ttu-id="48574-1681">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1681">Requirement</span></span>|<span data-ttu-id="48574-1682">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1682">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1683">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1684">1,8</span><span class="sxs-lookup"><span data-stu-id="48574-1684">1.8</span></span>|
|[<span data-ttu-id="48574-1685">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1686">ReadItem</span></span>|
|[<span data-ttu-id="48574-1687">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-1687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1688">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-1688">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1689">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1689">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="48574-1690">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="48574-1690">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="48574-1691">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="48574-1691">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="48574-p191">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="48574-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1695">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1695">Parameters</span></span>

|<span data-ttu-id="48574-1696">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1696">Name</span></span>|<span data-ttu-id="48574-1697">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1697">Type</span></span>|<span data-ttu-id="48574-1698">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1698">Attributes</span></span>|<span data-ttu-id="48574-1699">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1699">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="48574-1700">function</span><span class="sxs-lookup"><span data-stu-id="48574-1700">function</span></span>||<span data-ttu-id="48574-1701">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48574-1702">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1702">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="48574-1703">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-1703">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="48574-1704">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1704">Object</span></span>|<span data-ttu-id="48574-1705">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1705">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1706">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1706">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="48574-1707">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1707">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1708">Requirements</span></span>

|<span data-ttu-id="48574-1709">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1709">Requirement</span></span>|<span data-ttu-id="48574-1710">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1710">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1712">1.0</span><span class="sxs-lookup"><span data-stu-id="48574-1712">1.0</span></span>|
|[<span data-ttu-id="48574-1713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1714">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1714">ReadItem</span></span>|
|[<span data-ttu-id="48574-1715">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-1715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1716">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-1716">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1717">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1717">Example</span></span>

<span data-ttu-id="48574-p194">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="48574-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="48574-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="48574-1722">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="48574-1722">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="48574-1723">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="48574-1723">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="48574-1724">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-1724">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="48574-1725">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="48574-1725">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="48574-1726">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="48574-1726">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1727">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1727">Parameters</span></span>

|<span data-ttu-id="48574-1728">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1728">Name</span></span>|<span data-ttu-id="48574-1729">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1729">Type</span></span>|<span data-ttu-id="48574-1730">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1730">Attributes</span></span>|<span data-ttu-id="48574-1731">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1731">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="48574-1732">String</span><span class="sxs-lookup"><span data-stu-id="48574-1732">String</span></span>||<span data-ttu-id="48574-1733">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="48574-1733">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="48574-1734">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1734">Object</span></span>|<span data-ttu-id="48574-1735">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1735">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1736">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1736">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1737">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1737">Object</span></span>|<span data-ttu-id="48574-1738">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1738">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1739">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1739">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1740">function</span><span class="sxs-lookup"><span data-stu-id="48574-1740">function</span></span>|<span data-ttu-id="48574-1741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1741">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1742">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1742">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48574-1743">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="48574-1743">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48574-1744">Erros</span><span class="sxs-lookup"><span data-stu-id="48574-1744">Errors</span></span>

|<span data-ttu-id="48574-1745">Código de erro</span><span class="sxs-lookup"><span data-stu-id="48574-1745">Error code</span></span>|<span data-ttu-id="48574-1746">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1746">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="48574-1747">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="48574-1747">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1748">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1748">Requirements</span></span>

|<span data-ttu-id="48574-1749">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1749">Requirement</span></span>|<span data-ttu-id="48574-1750">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1750">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1751">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1752">1.1</span><span class="sxs-lookup"><span data-stu-id="48574-1752">1.1</span></span>|
|[<span data-ttu-id="48574-1753">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1754">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1754">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1755">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1756">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1756">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1757">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1757">Example</span></span>

<span data-ttu-id="48574-1758">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="48574-1758">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="48574-1759">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48574-1759">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="48574-1760">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="48574-1760">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="48574-1761">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="48574-1761">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1762">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1762">Parameters</span></span>

| <span data-ttu-id="48574-1763">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1763">Name</span></span> | <span data-ttu-id="48574-1764">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1764">Type</span></span> | <span data-ttu-id="48574-1765">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1765">Attributes</span></span> | <span data-ttu-id="48574-1766">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1766">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48574-1767">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48574-1767">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48574-1768">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="48574-1768">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="48574-1769">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1769">Object</span></span> | <span data-ttu-id="48574-1770">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1770">&lt;optional&gt;</span></span> | <span data-ttu-id="48574-1771">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1771">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48574-1772">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1772">Object</span></span> | <span data-ttu-id="48574-1773">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1773">&lt;optional&gt;</span></span> | <span data-ttu-id="48574-1774">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1774">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48574-1775">function</span><span class="sxs-lookup"><span data-stu-id="48574-1775">function</span></span>| <span data-ttu-id="48574-1776">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1776">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1777">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1778">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1778">Requirements</span></span>

|<span data-ttu-id="48574-1779">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1779">Requirement</span></span>| <span data-ttu-id="48574-1780">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1780">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1781">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48574-1782">1.7</span><span class="sxs-lookup"><span data-stu-id="48574-1782">1.7</span></span> |
|[<span data-ttu-id="48574-1783">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48574-1784">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48574-1784">ReadItem</span></span> |
|[<span data-ttu-id="48574-1785">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="48574-1785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48574-1786">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="48574-1786">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="48574-1787">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="48574-1787">saveAsync([options], callback)</span></span>

<span data-ttu-id="48574-1788">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="48574-1788">Asynchronously saves an item.</span></span>

<span data-ttu-id="48574-1789">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1789">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="48574-1790">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-1790">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="48574-1791">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="48574-1791">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1792">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="48574-1792">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="48574-1793">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="48574-1793">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="48574-p198">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="48574-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="48574-1797">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="48574-1797">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="48574-1798">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="48574-1798">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="48574-1799">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="48574-1799">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="48574-1800">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="48574-1800">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="48574-1801">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="48574-1801">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1802">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1802">Parameters</span></span>

|<span data-ttu-id="48574-1803">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1803">Name</span></span>|<span data-ttu-id="48574-1804">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1804">Type</span></span>|<span data-ttu-id="48574-1805">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1805">Attributes</span></span>|<span data-ttu-id="48574-1806">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1806">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48574-1807">Object</span><span class="sxs-lookup"><span data-stu-id="48574-1807">Object</span></span>|<span data-ttu-id="48574-1808">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1808">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1809">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1809">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1810">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1810">Object</span></span>|<span data-ttu-id="48574-1811">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1811">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1812">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1812">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48574-1813">function</span><span class="sxs-lookup"><span data-stu-id="48574-1813">function</span></span>||<span data-ttu-id="48574-1814">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1814">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48574-1815">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48574-1815">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1816">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1816">Requirements</span></span>

|<span data-ttu-id="48574-1817">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1817">Requirement</span></span>|<span data-ttu-id="48574-1818">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1818">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1819">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1819">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1820">1.3</span><span class="sxs-lookup"><span data-stu-id="48574-1820">1.3</span></span>|
|[<span data-ttu-id="48574-1821">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1821">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1822">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1822">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1823">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1823">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1824">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1824">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48574-1825">Exemplos</span><span class="sxs-lookup"><span data-stu-id="48574-1825">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="48574-p200">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="48574-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="48574-1828">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="48574-1828">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="48574-1829">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="48574-1829">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="48574-p201">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="48574-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48574-1833">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="48574-1833">Parameters</span></span>

|<span data-ttu-id="48574-1834">Nome</span><span class="sxs-lookup"><span data-stu-id="48574-1834">Name</span></span>|<span data-ttu-id="48574-1835">Tipo</span><span class="sxs-lookup"><span data-stu-id="48574-1835">Type</span></span>|<span data-ttu-id="48574-1836">Atributos</span><span class="sxs-lookup"><span data-stu-id="48574-1836">Attributes</span></span>|<span data-ttu-id="48574-1837">Descrição</span><span class="sxs-lookup"><span data-stu-id="48574-1837">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="48574-1838">String</span><span class="sxs-lookup"><span data-stu-id="48574-1838">String</span></span>||<span data-ttu-id="48574-p202">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="48574-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="48574-1842">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1842">Object</span></span>|<span data-ttu-id="48574-1843">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1843">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1844">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="48574-1844">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48574-1845">Objeto</span><span class="sxs-lookup"><span data-stu-id="48574-1845">Object</span></span>|<span data-ttu-id="48574-1846">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1846">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1847">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="48574-1847">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="48574-1848">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="48574-1848">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="48574-1849">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="48574-1849">&lt;optional&gt;</span></span>|<span data-ttu-id="48574-1850">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="48574-1850">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="48574-1851">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="48574-1851">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="48574-1852">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="48574-1852">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="48574-1853">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="48574-1853">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="48574-1854">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="48574-1854">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="48574-1855">function</span><span class="sxs-lookup"><span data-stu-id="48574-1855">function</span></span>||<span data-ttu-id="48574-1856">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48574-1856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48574-1857">Requisitos</span><span class="sxs-lookup"><span data-stu-id="48574-1857">Requirements</span></span>

|<span data-ttu-id="48574-1858">Requisito</span><span class="sxs-lookup"><span data-stu-id="48574-1858">Requirement</span></span>|<span data-ttu-id="48574-1859">Valor</span><span class="sxs-lookup"><span data-stu-id="48574-1859">Value</span></span>|
|---|---|
|[<span data-ttu-id="48574-1860">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="48574-1860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48574-1861">1.2</span><span class="sxs-lookup"><span data-stu-id="48574-1861">1.2</span></span>|
|[<span data-ttu-id="48574-1862">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="48574-1862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48574-1863">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48574-1863">ReadWriteItem</span></span>|
|[<span data-ttu-id="48574-1864">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="48574-1864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48574-1865">Escrever</span><span class="sxs-lookup"><span data-stu-id="48574-1865">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48574-1866">Exemplo</span><span class="sxs-lookup"><span data-stu-id="48574-1866">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
