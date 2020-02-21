---
title: Habilitar cenários de acesso de representante em um suplemento do Outlook
description: Descreve brevemente o acesso de representante e discute como configurar o suporte a suplementos.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6cee68af9efc02bbb474effaba1a898511aea531
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165828"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="0a16e-103">Habilitar cenários de acesso de representante em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="0a16e-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="0a16e-104">Um proprietário de caixa de correio pode usar o recurso de acesso de representante para [permitir que outra pessoa gerencie seus emails e calendários](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="0a16e-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="0a16e-105">Este artigo especifica a quais permissões de representante a API JavaScript do Office oferece suporte e descreve como habilitar cenários de acesso de representante no suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="0a16e-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0a16e-106">O acesso de representante não está disponível no Outlook no Mac, no Android e no iOS.</span><span class="sxs-lookup"><span data-stu-id="0a16e-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="0a16e-107">Essa funcionalidade pode ser disponibilizada no futuro.</span><span class="sxs-lookup"><span data-stu-id="0a16e-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="0a16e-108">O suporte para esse recurso foi introduzido no conjunto de requisitos 1,8.</span><span class="sxs-lookup"><span data-stu-id="0a16e-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="0a16e-109">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="0a16e-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="0a16e-110">Permissões com suporte para acesso de representante</span><span class="sxs-lookup"><span data-stu-id="0a16e-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="0a16e-111">A tabela a seguir descreve as permissões de representante que a API JavaScript do Office suporta.</span><span class="sxs-lookup"><span data-stu-id="0a16e-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="0a16e-112">Permissão</span><span class="sxs-lookup"><span data-stu-id="0a16e-112">Permission</span></span>|<span data-ttu-id="0a16e-113">Valor</span><span class="sxs-lookup"><span data-stu-id="0a16e-113">Value</span></span>|<span data-ttu-id="0a16e-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="0a16e-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="0a16e-115">Ler</span><span class="sxs-lookup"><span data-stu-id="0a16e-115">Read</span></span>|<span data-ttu-id="0a16e-116">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="0a16e-116">1 (000001)</span></span>|<span data-ttu-id="0a16e-117">Pode ler itens.</span><span class="sxs-lookup"><span data-stu-id="0a16e-117">Can read items.</span></span>|
|<span data-ttu-id="0a16e-118">Gravar</span><span class="sxs-lookup"><span data-stu-id="0a16e-118">Write</span></span>|<span data-ttu-id="0a16e-119">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="0a16e-119">2 (000010)</span></span>|<span data-ttu-id="0a16e-120">Pode criar itens.</span><span class="sxs-lookup"><span data-stu-id="0a16e-120">Can create items.</span></span>|
|<span data-ttu-id="0a16e-121">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="0a16e-121">DeleteOwn</span></span>|<span data-ttu-id="0a16e-122">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="0a16e-122">4 (000100)</span></span>|<span data-ttu-id="0a16e-123">Só pode excluir os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="0a16e-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="0a16e-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="0a16e-124">DeleteAll</span></span>|<span data-ttu-id="0a16e-125">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="0a16e-125">8 (001000)</span></span>|<span data-ttu-id="0a16e-126">Pode excluir qualquer item.</span><span class="sxs-lookup"><span data-stu-id="0a16e-126">Can delete any items.</span></span>|
|<span data-ttu-id="0a16e-127">EditOwn</span><span class="sxs-lookup"><span data-stu-id="0a16e-127">EditOwn</span></span>|<span data-ttu-id="0a16e-128">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="0a16e-128">16 (010000)</span></span>|<span data-ttu-id="0a16e-129">Só pode editar os itens que eles criaram.</span><span class="sxs-lookup"><span data-stu-id="0a16e-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="0a16e-130">EditAll</span><span class="sxs-lookup"><span data-stu-id="0a16e-130">EditAll</span></span>|<span data-ttu-id="0a16e-131">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="0a16e-131">32 (100000)</span></span>|<span data-ttu-id="0a16e-132">Pode editar qualquer item.</span><span class="sxs-lookup"><span data-stu-id="0a16e-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="0a16e-133">Atualmente, a API oferece suporte para obter permissões de representante existentes, mas não definir permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="0a16e-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="0a16e-134">O objeto [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma bitmask para indicar as permissões do representante.</span><span class="sxs-lookup"><span data-stu-id="0a16e-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="0a16e-135">Cada posição na bitmask representa uma permissão específica e, se estiver definida como `1` , o representante tem a respectiva permissão.</span><span class="sxs-lookup"><span data-stu-id="0a16e-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="0a16e-136">Por exemplo, se o segundo bit à direita é `1`, o representante tem permissão de **gravação** .</span><span class="sxs-lookup"><span data-stu-id="0a16e-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="0a16e-137">Você pode ver um exemplo de como verificar se há uma permissão específica na seção [executar uma operação como representante,](#perform-an-operation-as-delegate) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="0a16e-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="0a16e-138">Sincronizar entre clientes de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0a16e-138">Sync across mailbox clients</span></span>

<span data-ttu-id="0a16e-139">As atualizações de um representante para a caixa de correio do proprietário costumam ser sincronizadas imediatamente nas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="0a16e-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="0a16e-140">No entanto, se o suplemento usar operações REST ou EWS para definir uma propriedade estendida em um item, essas alterações poderão levar algumas horas para a sincronização. Em vez disso, recomendamos usar o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso.</span><span class="sxs-lookup"><span data-stu-id="0a16e-140">However, if the add-in uses REST or EWS operations to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="0a16e-141">Para saber mais, confira a [seção Propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "obter e definir metadados em um suplemento do Outlook".</span><span class="sxs-lookup"><span data-stu-id="0a16e-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="0a16e-142">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="0a16e-142">Configure the manifest</span></span>

<span data-ttu-id="0a16e-143">Para habilitar cenários de acesso de representante no suplemento, você deve definir o elemento [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` no manifesto no elemento `DesktopFormFactor`pai.</span><span class="sxs-lookup"><span data-stu-id="0a16e-143">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="0a16e-144">No momento, não há suporte para outros fatores de formulário.</span><span class="sxs-lookup"><span data-stu-id="0a16e-144">At present, other form factors are not supported.</span></span>

<span data-ttu-id="0a16e-145">O exemplo a seguir mostra `SupportsSharedFolders` o elemento definido `true` como em uma seção do manifesto.</span><span class="sxs-lookup"><span data-stu-id="0a16e-145">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="0a16e-146">Executar uma operação como representante</span><span class="sxs-lookup"><span data-stu-id="0a16e-146">Perform an operation as delegate</span></span>

<span data-ttu-id="0a16e-147">Você pode obter as propriedades compartilhadas de um item no modo de redação ou leitura chamando o método [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="0a16e-147">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="0a16e-148">Isso retorna um objeto [SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do representante, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.</span><span class="sxs-lookup"><span data-stu-id="0a16e-148">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="0a16e-149">O exemplo a seguir mostra como obter as propriedades compartilhadas de uma mensagem ou compromisso, verificar se o representante tem permissão de **gravação** e fazer uma chamada REST.</span><span class="sxs-lookup"><span data-stu-id="0a16e-149">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="0a16e-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="0a16e-150">See also</span></span>

- [<span data-ttu-id="0a16e-151">Permitir que outra pessoa Gerencie seu email e calendário</span><span class="sxs-lookup"><span data-stu-id="0a16e-151">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="0a16e-152">Compartilhamento de calendário no Office 365</span><span class="sxs-lookup"><span data-stu-id="0a16e-152">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="0a16e-153">Como solicitar elementos de manifesto</span><span class="sxs-lookup"><span data-stu-id="0a16e-153">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="0a16e-154">[Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="0a16e-154">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="0a16e-155">Operadores de JavaScript bit a bit</span><span class="sxs-lookup"><span data-stu-id="0a16e-155">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)