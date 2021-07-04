---
title: Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada
description: Discute como configurar o suporte ao complemento para pastas compartilhadas (a.k.a. acesso delegado) e caixas de correio compartilhadas.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 70578f2c78a9dd88efc9ba70d5599a13e121df53
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290709"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="9dfc5-104">Habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas em um Outlook de entrada</span><span class="sxs-lookup"><span data-stu-id="9dfc5-104">Enable shared folders and shared mailbox scenarios in an Outlook add-in</span></span>

<span data-ttu-id="9dfc5-105">Este artigo descreve como habilitar pastas compartilhadas (também conhecidas como acesso de representante) e cenários de caixa de correio compartilhada (agora em visualização [)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)no seu Outlook add-in, incluindo quais permissões a API JavaScript Office suporta.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-105">This article describes how to enable shared folders (also known as delegate access) and shared mailbox (now in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9dfc5-106">O suporte a esse recurso foi introduzido no [conjunto de requisitos 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="9dfc5-106">Support for this feature was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="9dfc5-107">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-setups"></a><span data-ttu-id="9dfc5-108">Configurações com suporte</span><span class="sxs-lookup"><span data-stu-id="9dfc5-108">Supported setups</span></span>

<span data-ttu-id="9dfc5-109">As seções a seguir descrevem configurações com suporte para caixas de correio compartilhadas (agora em visualização) e pastas compartilhadas.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-109">The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders.</span></span> <span data-ttu-id="9dfc5-110">As APIs de recurso podem não funcionar conforme o esperado em outras configurações.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-110">The feature APIs may not work as expected in other configurations.</span></span> <span data-ttu-id="9dfc5-111">Selecione a plataforma que você gostaria de aprender a configurar.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-111">Select the platform you'd like to learn how to configure.</span></span>

### <a name="windows"></a>[<span data-ttu-id="9dfc5-112">Windows</span><span class="sxs-lookup"><span data-stu-id="9dfc5-112">Windows</span></span>](#tab/windows)

#### <a name="shared-folders"></a><span data-ttu-id="9dfc5-113">Pastas compartilhadas</span><span class="sxs-lookup"><span data-stu-id="9dfc5-113">Shared folders</span></span>

<span data-ttu-id="9dfc5-114">O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="9dfc5-114">The mailbox owner must first [provide access to a delegate](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="9dfc5-115">O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa ao seu perfil" do artigo Gerenciar itens de calendário e email de [outra pessoa.](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-115">The delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="9dfc5-116">Caixas de correio compartilhadas (visualização)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-116">Shared mailboxes (preview)</span></span>

<span data-ttu-id="9dfc5-117">Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-117">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="9dfc5-118">No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-118">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="9dfc5-119">Um recurso Exchange Server conhecido como "automapping" está ativado por [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) padrão, o que significa que, posteriormente, a caixa de correio compartilhada deve aparecer automaticamente no aplicativo Outlook do usuário depois que o Outlook tiver sido fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-119">An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened.</span></span> <span data-ttu-id="9dfc5-120">No entanto, se um administrador tiver desabilitado a automação, o usuário deverá seguir as etapas manuais descritas na seção "Adicionar uma caixa de correio compartilhada ao Outlook" do artigo Abrir e usar uma caixa de correio compartilhada no [Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span><span class="sxs-lookup"><span data-stu-id="9dfc5-120">However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span></span>

> [!WARNING]
> <span data-ttu-id="9dfc5-121">Não **entre** na caixa de correio compartilhada com uma senha.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-121">Do **NOT** sign into the shared mailbox with a password.</span></span> <span data-ttu-id="9dfc5-122">As APIs de recurso não funcionarão nesse caso.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-122">The feature APIs won't work in that case.</span></span>

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="9dfc5-123">Navegador da Web – Outlook moderno</span><span class="sxs-lookup"><span data-stu-id="9dfc5-123">Web browser - modern Outlook</span></span>](#tab/modern)

#### <a name="shared-folders"></a><span data-ttu-id="9dfc5-124">Pastas compartilhadas</span><span class="sxs-lookup"><span data-stu-id="9dfc5-124">Shared folders</span></span>

<span data-ttu-id="9dfc5-125">O proprietário da caixa de correio [deve primeiro fornecer acesso a um representante](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) atualizando as permissões de pasta de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-125">The mailbox owner must first [provide access to a delegate](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) by updating the mailbox folder permissions.</span></span> <span data-ttu-id="9dfc5-126">O representante deve seguir as instruções descritas na seção "Adicionar caixa de correio de outra pessoa à sua lista de pastas Outlook Web App" do artigo Acessar a caixa de correio [de outra pessoa](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span><span class="sxs-lookup"><span data-stu-id="9dfc5-126">The delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="9dfc5-127">Caixas de correio compartilhadas (visualização)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-127">Shared mailboxes (preview)</span></span>

<span data-ttu-id="9dfc5-128">Exchange administradores de servidor podem criar e gerenciar caixas de correio compartilhadas para conjuntos de usuários acessarem.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-128">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="9dfc5-129">No momento, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) é a única versão de servidor com suporte para esse recurso.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-129">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="9dfc5-130">Depois de receber acesso, um usuário de caixa de correio compartilhada deve seguir as etapas descritas na seção "Adicionar a caixa de correio compartilhada para que ela seja exibida em sua caixa de correio principal" do artigo Abrir e usar uma caixa de correio compartilhada no [Outlook na Web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span><span class="sxs-lookup"><span data-stu-id="9dfc5-130">After receiving access, a shared mailbox user must follow the steps outlined in the "Add the shared mailbox so it displays under your primary mailbox" section of the article [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span></span>

> [!WARNING]
> <span data-ttu-id="9dfc5-131">NÃO **use** outras opções como "Abrir outra caixa de correio".</span><span class="sxs-lookup"><span data-stu-id="9dfc5-131">Do **NOT** use other options like "Open another mailbox".</span></span> <span data-ttu-id="9dfc5-132">As APIs de recurso podem não funcionar corretamente.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-132">The feature APIs may not work properly then.</span></span>

---

<span data-ttu-id="9dfc5-133">Para saber mais sobre onde os complementos fazem e não são ativados em geral, consulte a seção Itens de Caixa de Correio disponíveis para os [complementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-133">To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.</span></span>

## <a name="supported-permissions"></a><span data-ttu-id="9dfc5-134">Permissões com suporte</span><span class="sxs-lookup"><span data-stu-id="9dfc5-134">Supported permissions</span></span>

<span data-ttu-id="9dfc5-135">A tabela a seguir descreve as permissões que a API JavaScript Office suporta para representantes e usuários de caixa de correio compartilhados.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-135">The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.</span></span>

|<span data-ttu-id="9dfc5-136">Permissão</span><span class="sxs-lookup"><span data-stu-id="9dfc5-136">Permission</span></span>|<span data-ttu-id="9dfc5-137">Valor</span><span class="sxs-lookup"><span data-stu-id="9dfc5-137">Value</span></span>|<span data-ttu-id="9dfc5-138">Descrição</span><span class="sxs-lookup"><span data-stu-id="9dfc5-138">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="9dfc5-139">Read</span><span class="sxs-lookup"><span data-stu-id="9dfc5-139">Read</span></span>|<span data-ttu-id="9dfc5-140">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-140">1 (000001)</span></span>|<span data-ttu-id="9dfc5-141">Pode ler itens.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-141">Can read items.</span></span>|
|<span data-ttu-id="9dfc5-142">Gravar</span><span class="sxs-lookup"><span data-stu-id="9dfc5-142">Write</span></span>|<span data-ttu-id="9dfc5-143">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-143">2 (000010)</span></span>|<span data-ttu-id="9dfc5-144">Pode criar itens.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-144">Can create items.</span></span>|
|<span data-ttu-id="9dfc5-145">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="9dfc5-145">DeleteOwn</span></span>|<span data-ttu-id="9dfc5-146">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-146">4 (000100)</span></span>|<span data-ttu-id="9dfc5-147">Pode excluir apenas os itens criados.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-147">Can delete only the items they created.</span></span>|
|<span data-ttu-id="9dfc5-148">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="9dfc5-148">DeleteAll</span></span>|<span data-ttu-id="9dfc5-149">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-149">8 (001000)</span></span>|<span data-ttu-id="9dfc5-150">Pode excluir qualquer item.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-150">Can delete any items.</span></span>|
|<span data-ttu-id="9dfc5-151">EditOwn</span><span class="sxs-lookup"><span data-stu-id="9dfc5-151">EditOwn</span></span>|<span data-ttu-id="9dfc5-152">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-152">16 (010000)</span></span>|<span data-ttu-id="9dfc5-153">Pode editar apenas os itens criados.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-153">Can edit only the items they created.</span></span>|
|<span data-ttu-id="9dfc5-154">EditAll</span><span class="sxs-lookup"><span data-stu-id="9dfc5-154">EditAll</span></span>|<span data-ttu-id="9dfc5-155">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-155">32 (100000)</span></span>|<span data-ttu-id="9dfc5-156">Pode editar todos os itens.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-156">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="9dfc5-157">Atualmente, a API oferece suporte para obter permissões existentes, mas não para definir permissões.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-157">Currently the API supports getting existing permissions, but not setting permissions.</span></span>

<span data-ttu-id="9dfc5-158">O [objeto DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) é implementado usando uma máscara de bits para indicar as permissões.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-158">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions.</span></span> <span data-ttu-id="9dfc5-159">Cada posição na máscara de bits representa uma permissão específica e, se estiver definida como, o `1` usuário terá a respectiva permissão.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-159">Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission.</span></span> <span data-ttu-id="9dfc5-160">Por exemplo, se o segundo bit da direita for `1` , o usuário terá permissão **Gravar.**</span><span class="sxs-lookup"><span data-stu-id="9dfc5-160">For example, if the second bit from the right is `1`, then the user has **Write** permission.</span></span> <span data-ttu-id="9dfc5-161">Você pode ver um exemplo de como verificar uma permissão específica na seção Executar uma operação como representante ou usuário de caixa de correio [compartilhada](#perform-an-operation-as-delegate-or-shared-mailbox-user) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-161">You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.</span></span>

## <a name="sync-across-shared-folder-clients"></a><span data-ttu-id="9dfc5-162">Sincronizar entre clientes de pasta compartilhada</span><span class="sxs-lookup"><span data-stu-id="9dfc5-162">Sync across shared folder clients</span></span>

<span data-ttu-id="9dfc5-163">As atualizações de um representante para a caixa de correio do proprietário geralmente são sincronizadas entre caixas de correio imediatamente.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-163">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="9dfc5-164">No entanto, se as operações REST ou Exchange Web Services (EWS) foram usadas para definir uma propriedade estendida em um item, essas alterações podem levar algumas horas para sincronizar. Em vez disso, recomendamos que você use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) e APIs relacionadas para evitar esse atraso.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-164">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="9dfc5-165">Para saber mais, consulte a seção [propriedades personalizadas](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) do artigo "Obter e definir metadados em um Outlook de complemento".</span><span class="sxs-lookup"><span data-stu-id="9dfc5-165">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9dfc5-166">Em um cenário de representante, você não pode usar o EWS com os tokens atualmente fornecidos pela API office.js.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-166">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="9dfc5-167">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="9dfc5-167">Configure the manifest</span></span>

<span data-ttu-id="9dfc5-168">Para habilitar pastas compartilhadas e cenários de caixa de correio compartilhadas no seu complemento, você deve definir o [elemento SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) como no manifesto sob `true` o elemento pai `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="9dfc5-168">To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="9dfc5-169">Atualmente, outros fatores de formulário não são suportados.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-169">At present, other form factors are not supported.</span></span>

<span data-ttu-id="9dfc5-170">Para dar suporte a chamadas REST de um representante, de definir o nó [Permissões](../reference/manifest/permissions.md) no manifesto como `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="9dfc5-170">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="9dfc5-171">O exemplo a seguir mostra `SupportsSharedFolders` o elemento definido como em uma seção do `true` manifesto.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-171">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a><span data-ttu-id="9dfc5-172">Executar uma operação como representante ou usuário de caixa de correio compartilhada</span><span class="sxs-lookup"><span data-stu-id="9dfc5-172">Perform an operation as delegate or shared mailbox user</span></span>

<span data-ttu-id="9dfc5-173">Você pode obter as propriedades compartilhadas de um item no modo Redação ou Leitura chamando o [método item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-173">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="9dfc5-174">Isso retorna um [objeto SharedProperties](/javascript/api/outlook/office.sharedproperties) que atualmente fornece as permissões do usuário, o endereço de email do proprietário, a URL base da API REST e a caixa de correio de destino.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-174">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="9dfc5-175">O exemplo a seguir mostra como obter as propriedades compartilhadas de uma  mensagem ou compromisso, verificar se o representante ou usuário de caixa de correio compartilhada tem permissão Gravar e fazer uma chamada REST.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-175">The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="9dfc5-176">Como representante, você pode usar REST para obter o conteúdo de uma mensagem Outlook anexada a um item Outlook [ou postagem de grupo.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="9dfc5-176">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="9dfc5-177">Manipular a chamada REST em itens compartilhados e não compartilhados</span><span class="sxs-lookup"><span data-stu-id="9dfc5-177">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="9dfc5-178">Se você quiser chamar uma operação REST em um item, se o item é compartilhado ou não, você pode usar a API para determinar se o `getSharedPropertiesAsync` item é compartilhado.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-178">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="9dfc5-179">Depois disso, você pode construir a URL REST para a operação usando o objeto apropriado.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-179">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="9dfc5-180">Limitações</span><span class="sxs-lookup"><span data-stu-id="9dfc5-180">Limitations</span></span>

<span data-ttu-id="9dfc5-181">Dependendo dos cenários do seu complemento, há algumas limitações a considerar ao lidar com situações de pasta compartilhada ou de caixa de correio compartilhada.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-181">Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="9dfc5-182">Modo De composição de Mensagens</span><span class="sxs-lookup"><span data-stu-id="9dfc5-182">Message Compose mode</span></span>

<span data-ttu-id="9dfc5-183">No modo Redação de Mensagem, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) não é suportado no Outlook na Web ou no Windows a menos que as seguintes condições sejam atendidas.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-183">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.</span></span>

<span data-ttu-id="9dfc5-184">a.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-184">a.</span></span> <span data-ttu-id="9dfc5-185">**Delegar acesso/pastas compartilhadas**</span><span class="sxs-lookup"><span data-stu-id="9dfc5-185">**Delegate access/Shared folders**</span></span>

1. <span data-ttu-id="9dfc5-186">O proprietário da caixa de correio inicia uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-186">The mailbox owner starts a message.</span></span> <span data-ttu-id="9dfc5-187">Pode ser uma nova mensagem, uma resposta ou um encaminhamento.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-187">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="9dfc5-188">Eles salvam a mensagem e a movem de sua própria pasta **Rascunhos** para uma pasta compartilhada com o representante.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-188">They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.</span></span>
1. <span data-ttu-id="9dfc5-189">O representante abre o rascunho da pasta compartilhada e continua compondo.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-189">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="9dfc5-190">b.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-190">b.</span></span> <span data-ttu-id="9dfc5-191">**Caixa de correio compartilhada**</span><span class="sxs-lookup"><span data-stu-id="9dfc5-191">**Shared mailbox**</span></span>

1. <span data-ttu-id="9dfc5-192">Um usuário de caixa de correio compartilhado inicia uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-192">A shared mailbox user starts a message.</span></span> <span data-ttu-id="9dfc5-193">Pode ser uma nova mensagem, uma resposta ou um encaminhamento.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-193">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="9dfc5-194">Eles salvam a mensagem e a movem de sua própria pasta **Rascunhos** para uma pasta na caixa de correio compartilhada.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-194">They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.</span></span>
1. <span data-ttu-id="9dfc5-195">Outro usuário de caixa de correio compartilhada abre o rascunho da caixa de correio compartilhada e continua compondo.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-195">Another shared mailbox user opens the draft from the shared mailbox then continues composing.</span></span>

<span data-ttu-id="9dfc5-196">A mensagem agora está em um contexto compartilhado e os complementos que suportam esses cenários compartilhados podem obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-196">The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.</span></span> <span data-ttu-id="9dfc5-197">Depois que a mensagem é enviada, ela geralmente é encontrada na pasta Itens **Enviados do** remetente.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-197">After the message has been sent, it's usually found in the sender's **Sent Items** folder.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="9dfc5-198">REST e EWS</span><span class="sxs-lookup"><span data-stu-id="9dfc5-198">REST and EWS</span></span>

<span data-ttu-id="9dfc5-199">Seu complemento pode usar REST e a permissão do complemento deve ser definida como para habilitar o acesso REST à caixa de correio do proprietário ou à caixa de correio compartilhada conforme `ReadWriteMailbox` aplicável.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-199">Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox or to the shared mailbox as applicable.</span></span> <span data-ttu-id="9dfc5-200">Não há suporte para EWS.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-200">EWS is not supported.</span></span>

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a><span data-ttu-id="9dfc5-201">Usuário ou caixa de correio compartilhada oculta de uma lista de endereços</span><span class="sxs-lookup"><span data-stu-id="9dfc5-201">User or shared mailbox hidden from an address list</span></span>

<span data-ttu-id="9dfc5-202">Se um administrador ocultou um usuário ou endereço de caixa de correio compartilhado de uma lista de endereços, como a GAL (lista de endereços global), os itens de email afetados abriram no relatório de caixa de correio `Office.context.mailbox.item` como nulos.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-202">If an admin hid a user or shared mailbox address from an address list like the global address list (GAL), affected mail items opened in the mailbox report `Office.context.mailbox.item` as null.</span></span> <span data-ttu-id="9dfc5-203">Por exemplo, se o usuário abrir um item de email em uma caixa de correio compartilhada oculta da GAL, representar esse `Office.context.mailbox.item` item de email será nulo.</span><span class="sxs-lookup"><span data-stu-id="9dfc5-203">For example, if the user opens a mail item in a shared mailbox that's hidden from the GAL, `Office.context.mailbox.item` representing that mail item is null.</span></span>

## <a name="see-also"></a><span data-ttu-id="9dfc5-204">Confira também</span><span class="sxs-lookup"><span data-stu-id="9dfc5-204">See also</span></span>

- [<span data-ttu-id="9dfc5-205">Permitir que outra pessoa gerencie seu email e calendário</span><span class="sxs-lookup"><span data-stu-id="9dfc5-205">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="9dfc5-206">Compartilhamento de calendário em Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="9dfc5-206">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="9dfc5-207">Adicionar uma caixa de correio compartilhada Outlook</span><span class="sxs-lookup"><span data-stu-id="9dfc5-207">Add a shared mailbox to Outlook</span></span>](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [<span data-ttu-id="9dfc5-208">Como solicitar elementos de manifesto</span><span class="sxs-lookup"><span data-stu-id="9dfc5-208">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="9dfc5-209">[Máscara (computação)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="9dfc5-209">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="9dfc5-210">Operadores de bit do JavaScript</span><span class="sxs-lookup"><span data-stu-id="9dfc5-210">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)