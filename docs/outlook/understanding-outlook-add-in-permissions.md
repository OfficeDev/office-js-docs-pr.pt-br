---
title: Noções básicas sobre permissões de suplemento do Outlook
description: Suplementos do Outlook especificam o nível de permissão necessário em seu manifesto que incluem o modo restrito, ReadItem, ReadWriteItem ou ReadWriteMailbox.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 58d21a33034475b8c33b8449ece24c9dafc84e2b
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165691"
---
# <a name="understanding-outlook-add-in-permissions"></a><span data-ttu-id="05e56-103">Noções básicas sobre permissões de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="05e56-103">Understanding Outlook add-in permissions</span></span>

<span data-ttu-id="05e56-p101">Os suplementos do Outlook especificam o nível de permissão necessário nos seus manifestos. Os níveis disponíveis são **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox**. Esses níveis de permissões são cumulativos: **Restricted** é o nível mais baixo, e cada nível mais alto inclui as permissões dos níveis mais baixos. **ReadWriteMailbox** inclui todas as permissões com suporte.</span><span class="sxs-lookup"><span data-stu-id="05e56-p101">Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.</span></span>

<span data-ttu-id="05e56-p102">Você pode ver as permissões solicitadas por um suplemento de email antes de instalá-lo da [AppSource](https://appsource.microsoft.com). Também pode ver as permissões necessárias de suplementos instalados no Centro de Administração do Exchange.</span><span class="sxs-lookup"><span data-stu-id="05e56-p102">You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.</span></span>

## <a name="restricted-permission"></a><span data-ttu-id="05e56-110">Permissão restrita</span><span class="sxs-lookup"><span data-stu-id="05e56-110">Restricted permission</span></span>

<span data-ttu-id="05e56-p103">A permissão **Restricted** é o nível mais básico de permissão. Especifique a **Restricted** no elemento [Permissions](../reference/manifest/permissions.md), no manifesto, para solicitar essa permissão. O Outlook atribui essa permissão a um suplemento de email por padrão se o suplemento não solicitar uma permissão específica em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="05e56-p103">The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.</span></span>

### <a name="can-do"></a><span data-ttu-id="05e56-114">Pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-114">Can do</span></span>

- <span data-ttu-id="05e56-115">[Obter somente entidades específicas](match-strings-in-an-item-as-well-known-entities.md) (número de telefone, endereço, URL) do assunto ou corpo do item.</span><span class="sxs-lookup"><span data-stu-id="05e56-115">[Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.</span></span>

- <span data-ttu-id="05e56-116">Especificar uma [regra de ativação ItemIs](activation-rules.md#itemis-rule) que exige que o item atual em um formulário de leitura ou de redação seja um tipo de item específico, ou uma regra [ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md) que corresponde a um subconjunto menor de entidades conhecidas com suporte (número de telefone, endereço, URL) no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="05e56-116">Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.</span></span>

- <span data-ttu-id="05e56-117">Acessar quaisquer propriedades e métodos que **não** pertencem às informações específicas sobre o usuário ou o item (confira a próxima seção para ver a lista de membros que fazem isso).</span><span class="sxs-lookup"><span data-stu-id="05e56-117">Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).</span></span>

### <a name="cant-do"></a><span data-ttu-id="05e56-118">Não pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-118">Can't do</span></span>

- <span data-ttu-id="05e56-119">Usar uma regra [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) nas entidades de contato, endereço de email, sugestão de reunião ou sugestão de tarefa.</span><span class="sxs-lookup"><span data-stu-id="05e56-119">Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entitiy.</span></span>

- <span data-ttu-id="05e56-120">Usar a regra [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) ou [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule).</span><span class="sxs-lookup"><span data-stu-id="05e56-120">Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.</span></span>

- <span data-ttu-id="05e56-p104">Acessar os membros na lista a seguir que pertencem às informações do usuário ou do item. A tentativa de acessar os membros nessa lista retorna **null** e resulta em uma mensagem de erro informando que o Outlook que o suplemento de email tenha permissões elevadas.</span><span class="sxs-lookup"><span data-stu-id="05e56-p104">Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.</span></span>

    - [<span data-ttu-id="05e56-123">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-123">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-124">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-124">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-125">item.attachments</span><span class="sxs-lookup"><span data-stu-id="05e56-125">item.attachments</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-126">item.bcc</span><span class="sxs-lookup"><span data-stu-id="05e56-126">item.bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-127">item.body</span><span class="sxs-lookup"><span data-stu-id="05e56-127">item.body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-128">item.cc</span><span class="sxs-lookup"><span data-stu-id="05e56-128">item.cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-129">item.from</span><span class="sxs-lookup"><span data-stu-id="05e56-129">item.from</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-130">item.getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="05e56-130">item.getRegExMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-131">item.getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="05e56-131">item.getRegExMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-132">item.optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="05e56-132">item.optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-133">item.organizer</span><span class="sxs-lookup"><span data-stu-id="05e56-133">item.organizer</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-134">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-134">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-135">item.requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="05e56-135">item.requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-136">item.sender</span><span class="sxs-lookup"><span data-stu-id="05e56-136">item.sender</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-137">item.to</span><span class="sxs-lookup"><span data-stu-id="05e56-137">item.to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="05e56-138">mailbox.getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-138">mailbox.getCallbackTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="05e56-139">mailbox.getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-139">mailbox.getUserIdentityTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="05e56-140">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-140">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="05e56-141">mailbox.userProfile</span><span class="sxs-lookup"><span data-stu-id="05e56-141">mailbox.userProfile</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - <span data-ttu-id="05e56-142">[Body](/javascript/api/outlook/office.body) e todos os seus membros filhos</span><span class="sxs-lookup"><span data-stu-id="05e56-142">[Body](/javascript/api/outlook/office.body) and all its child members</span></span>
    - <span data-ttu-id="05e56-143">[Location](/javascript/api/outlook/office.location) e todos os seus membros filhos</span><span class="sxs-lookup"><span data-stu-id="05e56-143">[Location](/javascript/api/outlook/office.location) and all its child members</span></span>
    - <span data-ttu-id="05e56-144">[Recipients](/javascript/api/outlook/office.recipients) e todos os seus membros filhos</span><span class="sxs-lookup"><span data-stu-id="05e56-144">[Recipients](/javascript/api/outlook/office.recipients) and all its child members</span></span>
    - <span data-ttu-id="05e56-145">[Subject](/javascript/api/outlook/office.subject) e todos os seus membros filhos</span><span class="sxs-lookup"><span data-stu-id="05e56-145">[Subject](/javascript/api/outlook/office.subject) and all its child members</span></span>
    - <span data-ttu-id="05e56-146">[Time](/javascript/api/outlook/office.time) e todos os seus membros filhos</span><span class="sxs-lookup"><span data-stu-id="05e56-146">[Time](/javascript/api/outlook/office.time) and all its child members</span></span>

## <a name="readitem-permission"></a><span data-ttu-id="05e56-147">Permissão ReadItem</span><span class="sxs-lookup"><span data-stu-id="05e56-147">ReadItem permission</span></span>

<span data-ttu-id="05e56-p105">A permissão **ReadItem** é o nível seguinte de permissões no modelo de permissões. Especifique a **ReadItem** no elemento **Permissions**, no manifesto, para solicitar essa permissão.</span><span class="sxs-lookup"><span data-stu-id="05e56-p105">The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="05e56-150">Pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-150">Can do</span></span>

- <span data-ttu-id="05e56-151">[Ler todas as propriedades](item-data.md) do item atual em um formulário de leitura ou [de redação](get-and-set-item-data-in-a-compose-form.md), por exemplo, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) em um formulário de leitura e [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) em um formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="05e56-151">[Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.</span></span>

- <span data-ttu-id="05e56-152">[Obter um token de retorno de chamada para obter anexos do item](get-attachments-of-an-outlook-item.md) ou o item completo com os Serviços Web do Exchange (EWS) ou as [APIs REST do Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="05e56-152">[Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).</span></span>

- <span data-ttu-id="05e56-153">[Gravar propriedades personalizadas](/javascript/api/outlook/office.CustomProperties) definidas pelo suplemento nesse item.</span><span class="sxs-lookup"><span data-stu-id="05e56-153">[Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.</span></span>

- <span data-ttu-id="05e56-154">[Obter todas as entidades conhecidas existentes](match-strings-in-an-item-as-well-known-entities.md) do assunto ou do corpo do item, e não apenas um subconjunto.</span><span class="sxs-lookup"><span data-stu-id="05e56-154">[Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.</span></span>

- <span data-ttu-id="05e56-p106">Usar todas as [entidades conhecidas](activation-rules.md#itemhasknownentity-rule) nas regras [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ou [expressões regulares](activation-rules.md#itemhasregularexpressionmatch-rule) nas regras [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule). O exemplo a seguir segue a versão 1.1 do esquema. Ele mostra uma regra que ativa o suplemento se uma ou mais entidades conhecidas são encontradas no assunto ou no corpo da mensagem selecionada:</span><span class="sxs-lookup"><span data-stu-id="05e56-p106">Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a><span data-ttu-id="05e56-158">Não pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-158">Can't do</span></span>

- <span data-ttu-id="05e56-159">Usar o token fornecido por **mailbox.getCallbackTokenAsync** para:</span><span class="sxs-lookup"><span data-stu-id="05e56-159">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="05e56-160">Atualizar ou excluir o item atual usando a API REST do Outlook ou acessar outros itens na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="05e56-160">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="05e56-161">Obter o item de evento de calendário atual usando a API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="05e56-161">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="05e56-162">Usar qualquer uma das seguintes APIs:</span><span class="sxs-lookup"><span data-stu-id="05e56-162">Use any of the following APIs:</span></span>
    - [<span data-ttu-id="05e56-163">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-163">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="05e56-164">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-164">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-165">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-165">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-166">item.bcc.addAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-166">item.bcc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-167">item.bcc.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-167">item.bcc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-168">item.body.prependAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-168">item.body.prependAsync</span></span>](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [<span data-ttu-id="05e56-169">item.body.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-169">item.body.setAsync</span></span>](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [<span data-ttu-id="05e56-170">item.body.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-170">item.body.setSelectedDataAsync</span></span>](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [<span data-ttu-id="05e56-171">item.cc.addAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-171">item.cc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-172">item.cc.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-172">item.cc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-173">item.end.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-173">item.end.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="05e56-174">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-174">item.location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [<span data-ttu-id="05e56-175">item.optionalAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-175">item.optionalAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-176">item.optionalAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-176">item.optionalAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-177">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-177">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="05e56-178">item.requiredAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-178">item.requiredAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-179">item.requiredAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-179">item.requiredAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-180">item.start.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-180">item.start.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="05e56-181">item.subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-181">item.subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [<span data-ttu-id="05e56-182">item.to.addAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-182">item.to.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="05e56-183">item.to.setAsync</span><span class="sxs-lookup"><span data-stu-id="05e56-183">item.to.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a><span data-ttu-id="05e56-184">Permissão ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="05e56-184">ReadWriteItem permission</span></span>

<span data-ttu-id="05e56-185">Especifique o **ReadWriteItem** no elemento **Permissions**, no manifesto, para solicitar essa permissão.</span><span class="sxs-lookup"><span data-stu-id="05e56-185">Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission.</span></span> <span data-ttu-id="05e56-186">Os suplementos de email ativados nos formulários de redação que usam métodos de gravação (**Message.to.addAsync** ou **Message.to.setAsync**) devem usar pelo menos esse nível de permissão.</span><span class="sxs-lookup"><span data-stu-id="05e56-186">Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="05e56-187">Pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-187">Can do</span></span>

- <span data-ttu-id="05e56-188">[Ler e gravar todas as propriedades no nível do item](item-data.md) que está sendo visualizado ou redigido no Outlook.</span><span class="sxs-lookup"><span data-stu-id="05e56-188">[Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.</span></span>

- <span data-ttu-id="05e56-189">[Adicionar ou remover anexos](add-and-remove-attachments-to-an-item-in-a-compose-form.md) desse item.</span><span class="sxs-lookup"><span data-stu-id="05e56-189">[Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.</span></span>

- <span data-ttu-id="05e56-190">Usar todos os outros membros da API JavaScript para o Office que se aplicam a suplementos de email, exceto **Mailbox.makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="05e56-190">Use all other members of the JavaScript API for Office that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.</span></span>

### <a name="cant-do"></a><span data-ttu-id="05e56-191">Não pode ser feito</span><span class="sxs-lookup"><span data-stu-id="05e56-191">Can't do</span></span>

- <span data-ttu-id="05e56-192">Usar o token fornecido por **mailbox.getCallbackTokenAsync** para:</span><span class="sxs-lookup"><span data-stu-id="05e56-192">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="05e56-193">Atualizar ou excluir o item atual usando a API REST do Outlook ou acessar outros itens na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="05e56-193">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="05e56-194">Obter o item de evento de calendário atual usando a API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="05e56-194">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="05e56-195">Usar **mailbox.makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="05e56-195">Use **mailbox.makeEWSRequestAsync**.</span></span>

## <a name="readwritemailbox-permission"></a><span data-ttu-id="05e56-196">Permissão ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="05e56-196">ReadWriteMailbox permission</span></span>

<span data-ttu-id="05e56-p108">A permissão **ReadWriteMailbox** é o mais alto nível de permissão. Especifique a **ReadWriteMailbox** no elemento **Permissions**, no manifesto, para solicitar essa permissão.</span><span class="sxs-lookup"><span data-stu-id="05e56-p108">The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.</span></span>

<span data-ttu-id="05e56-199">Além do suporte que a permissão **ReadWriteItem** oferece, o token fornecido pela **mailbox.getCallbackTokenAsync** fornece acesso para usar as operações dos Serviços Web do Exchange (EWS) ou as APIs REST do Outlook para fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="05e56-199">In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:</span></span>

- <span data-ttu-id="05e56-200">Ler e gravar todas as propriedades de qualquer item na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="05e56-200">Read and write all properties of any item in the user's mailbox.</span></span>
- <span data-ttu-id="05e56-201">Criar, ler e gravar em qualquer pasta ou item nessa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="05e56-201">Create, read, and write to any folder or item in that mailbox.</span></span>
- <span data-ttu-id="05e56-202">Enviar um item dessa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="05e56-202">Send an item from that mailbox</span></span>

<span data-ttu-id="05e56-203">Por meio da **mailbox.makeEWSRequestAsync**, é possível acessar as seguintes operações dos EWS:</span><span class="sxs-lookup"><span data-stu-id="05e56-203">Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:</span></span>

- [<span data-ttu-id="05e56-204">CopyItem</span><span class="sxs-lookup"><span data-stu-id="05e56-204">CopyItem</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)
- [<span data-ttu-id="05e56-205">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="05e56-205">CreateFolder</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)
- [<span data-ttu-id="05e56-206">CreateItem</span><span class="sxs-lookup"><span data-stu-id="05e56-206">CreateItem</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)
- [<span data-ttu-id="05e56-207">FindConversation</span><span class="sxs-lookup"><span data-stu-id="05e56-207">FindConversation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)
- [<span data-ttu-id="05e56-208">FindFolder</span><span class="sxs-lookup"><span data-stu-id="05e56-208">FindFolder</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)
- [<span data-ttu-id="05e56-209">FindItem</span><span class="sxs-lookup"><span data-stu-id="05e56-209">FindItem</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)
- [<span data-ttu-id="05e56-210">GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="05e56-210">GetConversationItems</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [<span data-ttu-id="05e56-211">GetFolder</span><span class="sxs-lookup"><span data-stu-id="05e56-211">GetFolder</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)
- [<span data-ttu-id="05e56-212">GetItem</span><span class="sxs-lookup"><span data-stu-id="05e56-212">GetItem</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)
- [<span data-ttu-id="05e56-213">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="05e56-213">MarkAsJunk</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [<span data-ttu-id="05e56-214">MoveItem</span><span class="sxs-lookup"><span data-stu-id="05e56-214">MoveItem</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)
- [<span data-ttu-id="05e56-215">SendItem</span><span class="sxs-lookup"><span data-stu-id="05e56-215">SendItem</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)
- [<span data-ttu-id="05e56-216">UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="05e56-216">UpdateFolder</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [<span data-ttu-id="05e56-217">UpdateItem</span><span class="sxs-lookup"><span data-stu-id="05e56-217">UpdateItem</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)

<span data-ttu-id="05e56-218">A tentativa de usar uma operação sem suporte resulta em uma resposta de erro.</span><span class="sxs-lookup"><span data-stu-id="05e56-218">Attempting to use an unsupported operation will result in an error response.</span></span>

## <a name="see-also"></a><span data-ttu-id="05e56-219">Confira também</span><span class="sxs-lookup"><span data-stu-id="05e56-219">See also</span></span>

- [<span data-ttu-id="05e56-220">Privacidade, permissões e segurança de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="05e56-220">Privacy, permissions, and security for Outlook add-ins</span></span>](../develop/privacy-and-security.md)
- [<span data-ttu-id="05e56-221">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="05e56-221">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
