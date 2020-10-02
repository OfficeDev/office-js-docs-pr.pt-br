---
title: Habilitar cenários de acesso de representante em um suplemento do Outlook
description: Descreve brevemente o acesso de representante e discute como configurar o suporte a suplementos.
ms.date: 09/30/2020
localization_priority: Normal
ms.openlocfilehash: 68e9c8003f8d223a591283fd1a73f0a38bd3c8a4
ms.sourcegitcommit: 6c3a04acde57832feeaaa599148f93af7e3e36ea
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/02/2020
ms.locfileid: "48336416"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="08edb-103">Habilitar cenários de acesso de representante em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="08edb-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="08edb-104">Um proprietário de caixa de correio pode usar o recurso de acesso de representante para [permitir que outra pessoa gerencie seus emails e calendários](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="08edb-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="08edb-105">Este artigo especifica a quais permissões de representante a API JavaScript do Office oferece suporte e descreve como habilitar cenários de acesso de representante no suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="08edb-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="08edb-106">O acesso de representante não está disponível atualmente no Outlook no Android e iOS.</span><span class="sxs-lookup"><span data-stu-id="08edb-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="08edb-107">Além disso, esse recurso não está disponível atualmente com [caixas de correio compartilhadas em grupo](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="08edb-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="08edb-108">Essa funcionalidade pode ser disponibilizada no futuro.</span><span class="sxs-lookup"><span data-stu-id="08edb-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="08edb-109">O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="08edb-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="08edb-110">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="08edb-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="08edb-111">Permissões com suporte para acesso de representante</span><span class="sxs-lookup"><span data-stu-id="08edb-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="08edb-112">A tabela a seguir descreve as permissões de representante que a API JavaScript do Office suporta.</span><span class="sxs-lookup"><span data-stu-id="08edb-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="08edb-113">Permissão</span><span class="sxs-lookup"><span data-stu-id="08edb-113">Permission</span></span>|<span data-ttu-id="08edb-114">Valor</span><span class="sxs-lookup"><span data-stu-id="08edb-114">Value</span></span>|<span data-ttu-id="08edb-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="08edb-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="08edb-116">Ler</span><span class="sxs-lookup"><span data-stu-id="08edb-116">Read</span></span>|<span data-ttu-id="08edb-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="08edb-117">1 (000001)</span></span>|<span data-ttu-id="08edb-118">Pode ler itens.</span><span class="sxs-lookup"><span data-stu-id="08edb-118">Can read items.</span></span>|
|<span data-ttu-id="08edb-119">Gravação</span><span class="sxs-lookup"><span data-stu-id="08edb-119">Write</span></span>|<span data-ttu-id="08edb-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="08edb-120">2 (000010)</span></span>|<span data-ttu-id="08edb-121">Pode criar itens.</span><span class="sxs-lookup"><span data-stu-id="08edb-121">Can create items.</span></span>|
|<span data-ttu-id="08edb-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="08edb-122">DeleteOwn</span></span>|<span data-ttu-id="08edb-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="08edb-123">4 (000100)</span></span>|<span data-ttu-id="08edb-124">Só pode excluir os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="08edb-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="08edb-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="08edb-125">DeleteAll</span></span>|<span data-ttu-id="08edb-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="08edb-126">8 (001000)</span></span>|<span data-ttu-id="08edb-127">Pode excluir qualquer item.</span><span class="sxs-lookup"><span data-stu-id="08edb-127">Can delete any items.</span></span>|
|<span data-ttu-id="08edb-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="08edb-128">EditOwn</span></span>|<span data-ttu-id="08edb-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="08edb-129">16 (010000)</span></span>|<span data-ttu-id="08edb-130">Só pode editar os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="08edb-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="08edb-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="08edb-131">EditAll</span></span>|<span data-ttu-id="08edb-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="08edb-132">32 (100000)</span></span>|<span data-ttu-id="08edb-133">Pode editar qualquer item.</span><span class="sxs-lookup"><span data-stu-id="08edb-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="08edb-134">Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="08edb-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="08edb-135">O objeto [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma bitmask para indicar as permissões do representante.</span><span class="sxs-lookup"><span data-stu-id="08edb-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="08edb-136">Cada posição na bitmask representa uma permissão específica e, se estiver definida como `1` , o representante tem a respectiva permissão.</span><span class="sxs-lookup"><span data-stu-id="08edb-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="08edb-137">Por exemplo, se o segundo bit à direita é `1` , o representante tem permissão de **gravação** .</span><span class="sxs-lookup"><span data-stu-id="08edb-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="08edb-138">Você pode ver um exemplo de como verificar se há uma permissão específica na seção [executar uma operação como representante,](#perform-an-operation-as-delegate) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="08edb-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="08edb-139">Sincronizar entre clientes de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="08edb-139">Sync across mailbox clients</span></span>

<span data-ttu-id="08edb-140">As atualizações de um representante para a caixa de correio do proprietário costumam ser sincronizadas imediatamente nas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="08edb-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="08edb-141">No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações poderão levar algumas horas para a sincronização. Em vez disso, recomendamos usar o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso.</span><span class="sxs-lookup"><span data-stu-id="08edb-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="08edb-142">Para saber mais, confira a [seção Propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "obter e definir metadados em um suplemento do Outlook".</span><span class="sxs-lookup"><span data-stu-id="08edb-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="08edb-143">Em um cenário de representante, não é possível usar o EWS com os tokens atualmente fornecidos pela API office.js.</span><span class="sxs-lookup"><span data-stu-id="08edb-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="08edb-144">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="08edb-144">Configure the manifest</span></span>

<span data-ttu-id="08edb-145">Para habilitar cenários de acesso de representante no suplemento, você deve definir o elemento [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` no manifesto no elemento pai `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="08edb-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="08edb-146">No momento, não há suporte para outros fatores de formulário.</span><span class="sxs-lookup"><span data-stu-id="08edb-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="08edb-147">Para oferecer suporte a chamadas REST de um representante, defina o nó de [permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="08edb-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="08edb-148">O exemplo a seguir mostra o `SupportsSharedFolders` elemento definido como `true` em uma seção do manifesto.</span><span class="sxs-lookup"><span data-stu-id="08edb-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="08edb-149">Executar uma operação como representante</span><span class="sxs-lookup"><span data-stu-id="08edb-149">Perform an operation as delegate</span></span>

<span data-ttu-id="08edb-150">Você pode obter as propriedades compartilhadas de um item no modo de redação ou leitura chamando o método [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="08edb-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="08edb-151">Isso retorna um objeto [SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.</span><span class="sxs-lookup"><span data-stu-id="08edb-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="08edb-152">O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão de **gravação** e fazer uma chamada REST.</span><span class="sxs-lookup"><span data-stu-id="08edb-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="08edb-153">Como representante, você pode usar o REST para [obter o conteúdo de uma mensagem do Outlook anexada a um item do Outlook ou a uma postagem de grupo](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span><span class="sxs-lookup"><span data-stu-id="08edb-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="08edb-154">Gerenciar o resto de chamadas em itens compartilhados e não compartilhados</span><span class="sxs-lookup"><span data-stu-id="08edb-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="08edb-155">Se você quiser chamar uma operação REST em um item, se o item é ou não compartilhado, você pode usar a `getSharedPropertiesAsync` API para determinar se o item é compartilhado.</span><span class="sxs-lookup"><span data-stu-id="08edb-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="08edb-156">Depois, você pode construir a URL REST para a operação usando o objeto apropriado.</span><span class="sxs-lookup"><span data-stu-id="08edb-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="08edb-157">Limitações</span><span class="sxs-lookup"><span data-stu-id="08edb-157">Limitations</span></span>

<span data-ttu-id="08edb-158">Dependendo dos cenários do seu suplemento, há algumas limitações a serem consideradas ao lidar com as situações de representante.</span><span class="sxs-lookup"><span data-stu-id="08edb-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="08edb-159">REST e EWS</span><span class="sxs-lookup"><span data-stu-id="08edb-159">REST and EWS</span></span>

<span data-ttu-id="08edb-160">O suplemento pode usar REST, mas não EWS, e a permissão do suplemento deve ser definida para `ReadWriteMailbox` habilitar o acesso REST à caixa de correio do proprietário.</span><span class="sxs-lookup"><span data-stu-id="08edb-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="08edb-161">Modo de redação de mensagem</span><span class="sxs-lookup"><span data-stu-id="08edb-161">Message Compose mode</span></span>

<span data-ttu-id="08edb-162">No modo de redação de mensagens, o [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) não é suportado no Outlook na Web ou no Windows, a menos que as seguintes condições sejam atendidas.</span><span class="sxs-lookup"><span data-stu-id="08edb-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="08edb-163">O proprietário compartilha pelo menos uma pasta de caixa de correio com o representante.</span><span class="sxs-lookup"><span data-stu-id="08edb-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="08edb-164">O representante esboça uma mensagem na pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="08edb-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="08edb-165">Exemplos:</span><span class="sxs-lookup"><span data-stu-id="08edb-165">Examples:</span></span>

    - <span data-ttu-id="08edb-166">O representante responde ou encaminha um email na pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="08edb-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="08edb-167">O representante salva uma mensagem de rascunho e, em seguida, move-a da pasta **rascunhos** para a pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="08edb-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="08edb-168">O representante abre o rascunho da pasta compartilhada e continua a redigir.</span><span class="sxs-lookup"><span data-stu-id="08edb-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="08edb-169">Depois que a mensagem foi enviada, ela geralmente é encontrada na pasta **itens enviados** do representante.</span><span class="sxs-lookup"><span data-stu-id="08edb-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="08edb-170">Confira também</span><span class="sxs-lookup"><span data-stu-id="08edb-170">See also</span></span>

- [<span data-ttu-id="08edb-171">Permitir que outra pessoa Gerencie seu email e calendário</span><span class="sxs-lookup"><span data-stu-id="08edb-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="08edb-172">Compartilhamento de calendário no Office 365</span><span class="sxs-lookup"><span data-stu-id="08edb-172">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="08edb-173">Como solicitar elementos de manifesto</span><span class="sxs-lookup"><span data-stu-id="08edb-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="08edb-174">[Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="08edb-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="08edb-175">Operadores de JavaScript bit a bit</span><span class="sxs-lookup"><span data-stu-id="08edb-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)