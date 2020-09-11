---
title: Habilitar cenários de acesso de representante em um suplemento do Outlook
description: Descreve brevemente o acesso de representante e discute como configurar o suporte a suplementos.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 68b912d35f68cbf1177dd0b809994840092330a9
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430979"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="40e15-103">Habilitar cenários de acesso de representante em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="40e15-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="40e15-104">Um proprietário de caixa de correio pode usar o recurso de acesso de representante para [permitir que outra pessoa gerencie seus emails e calendários](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="40e15-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="40e15-105">Este artigo especifica a quais permissões de representante a API JavaScript do Office oferece suporte e descreve como habilitar cenários de acesso de representante no suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="40e15-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="40e15-106">O acesso de representante não está disponível atualmente no Outlook no Android e iOS.</span><span class="sxs-lookup"><span data-stu-id="40e15-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="40e15-107">Além disso, esse recurso não está disponível atualmente com [caixas de correio compartilhadas em grupo](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="40e15-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="40e15-108">Essa funcionalidade pode ser disponibilizada no futuro.</span><span class="sxs-lookup"><span data-stu-id="40e15-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="40e15-109">O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="40e15-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="40e15-110">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="40e15-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="40e15-111">Permissões com suporte para acesso de representante</span><span class="sxs-lookup"><span data-stu-id="40e15-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="40e15-112">A tabela a seguir descreve as permissões de representante que a API JavaScript do Office suporta.</span><span class="sxs-lookup"><span data-stu-id="40e15-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="40e15-113">Permissão</span><span class="sxs-lookup"><span data-stu-id="40e15-113">Permission</span></span>|<span data-ttu-id="40e15-114">Valor</span><span class="sxs-lookup"><span data-stu-id="40e15-114">Value</span></span>|<span data-ttu-id="40e15-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="40e15-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="40e15-116">Ler</span><span class="sxs-lookup"><span data-stu-id="40e15-116">Read</span></span>|<span data-ttu-id="40e15-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="40e15-117">1 (000001)</span></span>|<span data-ttu-id="40e15-118">Pode ler itens.</span><span class="sxs-lookup"><span data-stu-id="40e15-118">Can read items.</span></span>|
|<span data-ttu-id="40e15-119">Gravar</span><span class="sxs-lookup"><span data-stu-id="40e15-119">Write</span></span>|<span data-ttu-id="40e15-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="40e15-120">2 (000010)</span></span>|<span data-ttu-id="40e15-121">Pode criar itens.</span><span class="sxs-lookup"><span data-stu-id="40e15-121">Can create items.</span></span>|
|<span data-ttu-id="40e15-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="40e15-122">DeleteOwn</span></span>|<span data-ttu-id="40e15-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="40e15-123">4 (000100)</span></span>|<span data-ttu-id="40e15-124">Só pode excluir os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="40e15-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="40e15-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="40e15-125">DeleteAll</span></span>|<span data-ttu-id="40e15-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="40e15-126">8 (001000)</span></span>|<span data-ttu-id="40e15-127">Pode excluir qualquer item.</span><span class="sxs-lookup"><span data-stu-id="40e15-127">Can delete any items.</span></span>|
|<span data-ttu-id="40e15-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="40e15-128">EditOwn</span></span>|<span data-ttu-id="40e15-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="40e15-129">16 (010000)</span></span>|<span data-ttu-id="40e15-130">Só pode editar os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="40e15-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="40e15-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="40e15-131">EditAll</span></span>|<span data-ttu-id="40e15-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="40e15-132">32 (100000)</span></span>|<span data-ttu-id="40e15-133">Pode editar qualquer item.</span><span class="sxs-lookup"><span data-stu-id="40e15-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="40e15-134">Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="40e15-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="40e15-135">O objeto [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma bitmask para indicar as permissões do representante.</span><span class="sxs-lookup"><span data-stu-id="40e15-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="40e15-136">Cada posição na bitmask representa uma permissão específica e, se estiver definida como `1` , o representante tem a respectiva permissão.</span><span class="sxs-lookup"><span data-stu-id="40e15-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="40e15-137">Por exemplo, se o segundo bit à direita é `1` , o representante tem permissão de **gravação** .</span><span class="sxs-lookup"><span data-stu-id="40e15-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="40e15-138">Você pode ver um exemplo de como verificar se há uma permissão específica na seção [executar uma operação como representante,](#perform-an-operation-as-delegate) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="40e15-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="40e15-139">Sincronizar entre clientes de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="40e15-139">Sync across mailbox clients</span></span>

<span data-ttu-id="40e15-140">As atualizações de um representante para a caixa de correio do proprietário costumam ser sincronizadas imediatamente nas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="40e15-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="40e15-141">No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações poderão levar algumas horas para a sincronização. Em vez disso, recomendamos usar o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso.</span><span class="sxs-lookup"><span data-stu-id="40e15-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="40e15-142">Para saber mais, confira a [seção Propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "obter e definir metadados em um suplemento do Outlook".</span><span class="sxs-lookup"><span data-stu-id="40e15-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="40e15-143">Em um cenário de representante, não é possível usar o EWS com os tokens atualmente fornecidos pela API office.js.</span><span class="sxs-lookup"><span data-stu-id="40e15-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="40e15-144">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="40e15-144">Configure the manifest</span></span>

<span data-ttu-id="40e15-145">Para habilitar cenários de acesso de representante no suplemento, você deve definir o elemento [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` no manifesto no elemento pai `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="40e15-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="40e15-146">No momento, não há suporte para outros fatores de formulário.</span><span class="sxs-lookup"><span data-stu-id="40e15-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="40e15-147">Para oferecer suporte a chamadas REST de um representante, defina o nó de [permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="40e15-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="40e15-148">O exemplo a seguir mostra o `SupportsSharedFolders` elemento definido como `true` em uma seção do manifesto.</span><span class="sxs-lookup"><span data-stu-id="40e15-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="40e15-149">Executar uma operação como representante</span><span class="sxs-lookup"><span data-stu-id="40e15-149">Perform an operation as delegate</span></span>

<span data-ttu-id="40e15-150">Você pode obter as propriedades compartilhadas de um item no modo de redação ou leitura chamando o método [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="40e15-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="40e15-151">Isso retorna um objeto [SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.</span><span class="sxs-lookup"><span data-stu-id="40e15-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="40e15-152">Em um cenário de representante, o suplemento pode usar REST, mas não EWS, e a permissão do suplemento deve ser definida para `ReadWriteMailbox` habilitar o acesso REST à caixa de correio do proprietário.</span><span class="sxs-lookup"><span data-stu-id="40e15-152">In a delegate scenario, your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

<span data-ttu-id="40e15-153">O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão de **gravação** e fazer uma chamada REST.</span><span class="sxs-lookup"><span data-stu-id="40e15-153">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="40e15-154">Como representante, você pode usar o REST para [obter o conteúdo de uma mensagem do Outlook anexada a um item do Outlook ou a uma postagem de grupo](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span><span class="sxs-lookup"><span data-stu-id="40e15-154">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="see-also"></a><span data-ttu-id="40e15-155">Confira também</span><span class="sxs-lookup"><span data-stu-id="40e15-155">See also</span></span>

- [<span data-ttu-id="40e15-156">Permitir que outra pessoa Gerencie seu email e calendário</span><span class="sxs-lookup"><span data-stu-id="40e15-156">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="40e15-157">Compartilhamento de calendário no Office 365</span><span class="sxs-lookup"><span data-stu-id="40e15-157">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="40e15-158">Como solicitar elementos de manifesto</span><span class="sxs-lookup"><span data-stu-id="40e15-158">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="40e15-159">[Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="40e15-159">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="40e15-160">Operadores de JavaScript bit a bit</span><span class="sxs-lookup"><span data-stu-id="40e15-160">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)