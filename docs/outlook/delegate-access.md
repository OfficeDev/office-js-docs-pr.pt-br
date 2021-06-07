---
title: Habilitar cenários de acesso de representante em um Outlook de entrada
description: Descreve brevemente o acesso de representantes e discute como configurar o suporte ao complemento.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783488"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="1bcb3-103">Habilitar cenários de acesso de representante em um Outlook de entrada</span><span class="sxs-lookup"><span data-stu-id="1bcb3-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="1bcb3-104">Um proprietário de caixa de correio pode usar o recurso de acesso de representante [para permitir que outra pessoa gerencie seus emails e calendários.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="1bcb3-105">Este artigo especifica quais permissões delegar Office API JavaScript suporta e descreve como habilitar cenários de acesso de representante em seu Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1bcb3-106">O acesso de representante não está disponível no Outlook android e iOS.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="1bcb3-107">Além disso, esse recurso não está disponível no momento com caixas de correio [compartilhadas](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) em grupo Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="1bcb3-108">Essa funcionalidade pode ser disponibilizada no futuro.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="1bcb3-109">O suporte para esse recurso foi introduzido no conjunto de requisitos 1.8.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="1bcb3-110">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="1bcb3-111">Permissões com suporte para acesso de representante</span><span class="sxs-lookup"><span data-stu-id="1bcb3-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="1bcb3-112">A tabela a seguir descreve as permissões de representante suportadas Office API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="1bcb3-113">Permissão</span><span class="sxs-lookup"><span data-stu-id="1bcb3-113">Permission</span></span>|<span data-ttu-id="1bcb3-114">Valor</span><span class="sxs-lookup"><span data-stu-id="1bcb3-114">Value</span></span>|<span data-ttu-id="1bcb3-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="1bcb3-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="1bcb3-116">Leitura</span><span class="sxs-lookup"><span data-stu-id="1bcb3-116">Read</span></span>|<span data-ttu-id="1bcb3-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-117">1 (000001)</span></span>|<span data-ttu-id="1bcb3-118">Pode ler itens.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-118">Can read items.</span></span>|
|<span data-ttu-id="1bcb3-119">Gravar</span><span class="sxs-lookup"><span data-stu-id="1bcb3-119">Write</span></span>|<span data-ttu-id="1bcb3-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-120">2 (000010)</span></span>|<span data-ttu-id="1bcb3-121">Pode criar itens.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-121">Can create items.</span></span>|
|<span data-ttu-id="1bcb3-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="1bcb3-122">DeleteOwn</span></span>|<span data-ttu-id="1bcb3-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-123">4 (000100)</span></span>|<span data-ttu-id="1bcb3-124">Pode excluir apenas os itens criados.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="1bcb3-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="1bcb3-125">DeleteAll</span></span>|<span data-ttu-id="1bcb3-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-126">8 (001000)</span></span>|<span data-ttu-id="1bcb3-127">Pode excluir qualquer item.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-127">Can delete any items.</span></span>|
|<span data-ttu-id="1bcb3-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="1bcb3-128">EditOwn</span></span>|<span data-ttu-id="1bcb3-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-129">16 (010000)</span></span>|<span data-ttu-id="1bcb3-130">Pode editar apenas os itens criados.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="1bcb3-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="1bcb3-131">EditAll</span></span>|<span data-ttu-id="1bcb3-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-132">32 (100000)</span></span>|<span data-ttu-id="1bcb3-133">Pode editar todos os itens.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="1bcb3-134">Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="1bcb3-135">O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões do representante.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="1bcb3-136">Cada posição na máscara de bits representa uma permissão específica e, se estiver definida como, o `1` representante terá a respectiva permissão.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="1bcb3-137">Por exemplo, se o segundo bit da direita for `1` , o representante terá permissão **Gravar.**</span><span class="sxs-lookup"><span data-stu-id="1bcb3-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="1bcb3-138">Você pode ver um exemplo de como verificar se há uma permissão específica na seção [Executar](#perform-an-operation-as-delegate) uma operação como representante posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="1bcb3-139">Sincronizar entre clientes de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1bcb3-139">Sync across mailbox clients</span></span>

<span data-ttu-id="1bcb3-140">As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="1bcb3-141">No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações podem levar algumas horas para sincronizar. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="1bcb3-142">Para saber mais, consulte a seção [propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um Outlook de complemento".</span><span class="sxs-lookup"><span data-stu-id="1bcb3-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1bcb3-143">Em um cenário de representante, você não pode usar o EWS com os tokens atualmente fornecidos pela API office.js.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="1bcb3-144">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="1bcb3-144">Configure the manifest</span></span>

<span data-ttu-id="1bcb3-145">Para habilitar cenários de acesso de representante no seu add-in, você deve definir o [elemento SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) como no `true` manifesto sob o elemento pai `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="1bcb3-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="1bcb3-146">Atualmente, outros fatores de formulário não são suportados.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="1bcb3-147">Para dar suporte a chamadas REST de um representante, de definir o nó [Permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="1bcb3-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="1bcb3-148">O exemplo a seguir mostra `SupportsSharedFolders` o elemento definido como em uma seção do `true` manifesto.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="1bcb3-149">Executar uma operação como representante</span><span class="sxs-lookup"><span data-stu-id="1bcb3-149">Perform an operation as delegate</span></span>

<span data-ttu-id="1bcb3-150">Você pode obter as propriedades compartilhadas de um item no modo Redação ou Leitura chamando o [método item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="1bcb3-151">Isso retorna um [objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="1bcb3-152">O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão **Gravar** e fazer uma chamada REST.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="1bcb3-153">Como representante, você pode usar REST para obter o conteúdo de uma mensagem Outlook anexada a um item Outlook [ou postagem de grupo.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="1bcb3-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="1bcb3-154">Manipular a chamada REST em itens compartilhados e não compartilhados</span><span class="sxs-lookup"><span data-stu-id="1bcb3-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="1bcb3-155">Se você quiser chamar uma operação REST em um item, se o item é compartilhado ou não, você pode usar a API para determinar se o `getSharedPropertiesAsync` item é compartilhado.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="1bcb3-156">Depois disso, você pode construir a URL REST para a operação usando o objeto apropriado.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="1bcb3-157">Limitações</span><span class="sxs-lookup"><span data-stu-id="1bcb3-157">Limitations</span></span>

<span data-ttu-id="1bcb3-158">Dependendo dos cenários do seu complemento, há algumas limitações a considerar ao lidar com situações de representante.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="1bcb3-159">REST e EWS</span><span class="sxs-lookup"><span data-stu-id="1bcb3-159">REST and EWS</span></span>

<span data-ttu-id="1bcb3-160">Seu complemento pode usar REST, mas não EWS, e a permissão do add-in deve ser definida para habilitar o acesso REST à caixa de correio `ReadWriteMailbox` do proprietário.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="1bcb3-161">Modo De composição de Mensagens</span><span class="sxs-lookup"><span data-stu-id="1bcb3-161">Message Compose mode</span></span>

<span data-ttu-id="1bcb3-162">No modo Redação de Mensagem, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) não é suportado Outlook na Web ou Windows a menos que as seguintes condições sejam atendidas.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="1bcb3-163">O proprietário compartilha pelo menos uma pasta de caixa de correio com o representante.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="1bcb3-164">O representante esboça uma mensagem na pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="1bcb3-165">Exemplos:</span><span class="sxs-lookup"><span data-stu-id="1bcb3-165">Examples:</span></span>

    - <span data-ttu-id="1bcb3-166">O representante responde ou encaminha um email na pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="1bcb3-167">O representante salva uma mensagem de rascunho e a move de sua própria pasta **Rascunhos** para a pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="1bcb3-168">O representante abre o rascunho da pasta compartilhada e continua compondo.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="1bcb3-169">Depois que a mensagem é enviada, ela geralmente é encontrada na pasta Itens **Enviados do** representante.</span><span class="sxs-lookup"><span data-stu-id="1bcb3-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="1bcb3-170">Confira também</span><span class="sxs-lookup"><span data-stu-id="1bcb3-170">See also</span></span>

- [<span data-ttu-id="1bcb3-171">Permitir que outra pessoa gerencie seu email e calendário</span><span class="sxs-lookup"><span data-stu-id="1bcb3-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="1bcb3-172">Compartilhamento de calendário em Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="1bcb3-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="1bcb3-173">Como solicitar elementos de manifesto</span><span class="sxs-lookup"><span data-stu-id="1bcb3-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="1bcb3-174">[Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="1bcb3-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="1bcb3-175">Operadores de bit do JavaScript</span><span class="sxs-lookup"><span data-stu-id="1bcb3-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)