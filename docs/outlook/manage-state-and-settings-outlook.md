---
title: Gerenciar o estado e as configurações de um suplemento do Outlook
description: Saiba como persistir o estado e as configurações do suplemento para um suplemento do Outlook.
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 3cb4f7d6e31fd4d37e01939f743682f60f24f959
ms.sourcegitcommit: 9da68c00ecc00a2f307757e0f5a903a8e31b7769
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/22/2020
ms.locfileid: "43785778"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="1feb2-103">Gerenciar o estado e as configurações de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="1feb2-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="1feb2-104">Revise o [estado e as configurações do suplemento persistentes](../develop/persisting-add-in-state-and-settings.md) na seção **principais conceitos** desta documentação antes de ler este artigo.</span><span class="sxs-lookup"><span data-stu-id="1feb2-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="1feb2-105">Para suplementos do Outlook, a API JavaScript do Office fornece objetos [RoamingSettings](/javascript/api/outlook/office.roamingsettings) e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado do suplemento entre as sessões, conforme descrito na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1feb2-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="1feb2-106">Em todos os casos, os valores de configurações salvos são associados à [Id](../reference/manifest/id.md) do suplemento que os criou.</span><span class="sxs-lookup"><span data-stu-id="1feb2-106">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="1feb2-107">**Objeto**</span><span class="sxs-lookup"><span data-stu-id="1feb2-107">**Object**</span></span>|<span data-ttu-id="1feb2-108">**Local de armazenamento**</span><span class="sxs-lookup"><span data-stu-id="1feb2-108">**Storage location**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="1feb2-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1feb2-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="1feb2-110">A caixa de correio do Exchange Server do usuário onde o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="1feb2-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="1feb2-111">Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem "mover-se" com o usuário e estão disponíveis para o suplemento quando ele estiver sendo executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessar a caixa de correio desse usuário.</span><span class="sxs-lookup"><span data-stu-id="1feb2-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="1feb2-112">As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="1feb2-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="1feb2-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="1feb2-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="1feb2-p103">A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.</span><span class="sxs-lookup"><span data-stu-id="1feb2-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="1feb2-116">Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis</span><span class="sxs-lookup"><span data-stu-id="1feb2-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="1feb2-117">Um suplemento do Outlook pode usar o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) para salvar o estado e os dados de configurações do suplemento específico da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="1feb2-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="1feb2-118">Esses dados são acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento.</span><span class="sxs-lookup"><span data-stu-id="1feb2-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="1feb2-119">Os dados são armazenados na caixa de correio do usuário do Exchange Server e ficam acessíveis quando esse usuário faz logon em sua conta e executa o suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1feb2-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="1feb2-120">Carregar configurações de roaming</span><span class="sxs-lookup"><span data-stu-id="1feb2-120">Loading roaming settings</span></span>

<span data-ttu-id="1feb2-p105">Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](/javascript/api/office). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.</span><span class="sxs-lookup"><span data-stu-id="1feb2-p105">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="1feb2-123">Criar ou atribuir uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="1feb2-123">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="1feb2-p106">Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="1feb2-p106">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

<span data-ttu-id="1feb2-126">O método **saveAsync** salva as configurações móveis de forma assíncrona e utiliza uma função de retorno de chamada opcional.</span><span class="sxs-lookup"><span data-stu-id="1feb2-126">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="1feb2-127">Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="1feb2-127">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="1feb2-128">Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](/javascript/api/office/office.asyncresult) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="1feb2-128">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/office/office.asyncresult) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="1feb2-129">Remover uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="1feb2-129">Removing a roaming setting</span></span>

<span data-ttu-id="1feb2-130">Também estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) para remover a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1feb2-130">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="1feb2-131">Como salvar configurações por item para suplementos do Outlook como propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="1feb2-131">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="1feb2-p108">As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.</span><span class="sxs-lookup"><span data-stu-id="1feb2-p108">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="1feb2-p109">Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicitação de reunião específico, você deve carregar as propriedades na memória chamando o método [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](/javascript/api/outlook/office.customproperties#set-name--value-) e [get](/javascript/api/outlook/office.roamingsettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na memória. Para salvar as alterações feitas nas propriedades personalizadas do item, você deve usar o método [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) para persistir as alterações no item no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1feb2-p109">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="1feb2-139">Exemplo de propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="1feb2-139">Custom properties example</span></span>

<span data-ttu-id="1feb2-p110">O exemplo a seguir mostra um conjunto simplificado de funções para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1feb2-p110">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="1feb2-142">Um suplemento do Outlook que usa essas funções recupera as propriedades personalizadas chamando o método **obter** na variável `_customProps`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="1feb2-142">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="1feb2-143">Este exemplo inclui as seguintes funções:</span><span class="sxs-lookup"><span data-stu-id="1feb2-143">This example includes the following functions:</span></span>

|<span data-ttu-id="1feb2-144">**Nome da função**</span><span class="sxs-lookup"><span data-stu-id="1feb2-144">**Function name**</span></span>|<span data-ttu-id="1feb2-145">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="1feb2-145">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="1feb2-146">Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1feb2-146">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="1feb2-147">Obtém as propriedades personalizadas que são retornadas do servidor Exchange e as salva para uso posterior.</span><span class="sxs-lookup"><span data-stu-id="1feb2-147">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="1feb2-148">Define ou atualiza uma propriedade específica e salva a alteração no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1feb2-148">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="1feb2-149">Remove uma propriedade específica e persiste a remoção no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1feb2-149">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="1feb2-150">Retorno de chamada para chamadas ao método **saveAsync** nas funções `updateProperty` e `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="1feb2-150">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

## <a name="see-also"></a><span data-ttu-id="1feb2-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="1feb2-151">See also</span></span>

- [<span data-ttu-id="1feb2-152">Persistir o estado e as configurações do suplemento</span><span class="sxs-lookup"><span data-stu-id="1feb2-152">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="1feb2-153">Inicialize seu suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="1feb2-153">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)