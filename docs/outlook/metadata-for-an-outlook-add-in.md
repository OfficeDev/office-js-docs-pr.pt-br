---
title: Obter e definir metadados em um suplemento do Outlook
description: Gerencie dados personalizados no suplemento do Outlook usando configurações de roaming ou propriedades personalizadas.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: abcae0766079f090ec15b9d11ec66c43355bfb0f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431238"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a><span data-ttu-id="33ab6-103">Obter e definir metadados de suplemento para um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="33ab6-103">Get and set add-in metadata for an Outlook add-in</span></span>

<span data-ttu-id="33ab6-104">Você pode gerenciar dados personalizados em seu suplemento do Outlook usando um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="33ab6-104">You can manage custom data in your Outlook add-in by using either of the following:</span></span>

- <span data-ttu-id="33ab6-105">As configurações de roaming, que gerenciam dados personalizados para uma caixa de correio de usuário.</span><span class="sxs-lookup"><span data-stu-id="33ab6-105">Roaming settings, which manage custom data for a user's mailbox.</span></span>
- <span data-ttu-id="33ab6-106">Propriedades personalizadas, que gerenciam dados personalizados para um item na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="33ab6-106">Custom properties, which manage custom data for an item in a user's mailbox.</span></span>

<span data-ttu-id="33ab6-p101">Ambos dão acesso a dados personalizados que só podem ser acessados por seu suplemento do Outlook, mas cada método armazena os dados separadamente. Ou seja, os dados armazenados por meio de configurações de roaming não podem ser acessados por propriedades personalizadas e vice-versa. Os dados são armazenados no servidor dessa caixa de correio e podem ser acessados nas sessões subsequentes do Outlook em todos os fatores forma a que o suplemento dê suporte.</span><span class="sxs-lookup"><span data-stu-id="33ab6-p101">Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports.</span></span>

## <a name="custom-data-per-mailbox-roaming-settings"></a><span data-ttu-id="33ab6-110">Dados personalizados por caixa de correio: configurações de roaming</span><span class="sxs-lookup"><span data-stu-id="33ab6-110">Custom data per mailbox: roaming settings</span></span>

<span data-ttu-id="33ab6-p102">Você pode especificar dados específicos para uma caixa de correio do Exchange de um usuário usando o objeto [RoamingSettings](/javascript/api/outlook/office.RoamingSettings). Exemplos desses dados incluem os dados pessoais e as preferências do usuário. O suplemento de email pode acessar as configurações de roaming quando faz roaming em qualquer dispositivo no qual deva ser executado (área de trabalho, tablet ou smartphone).</span><span class="sxs-lookup"><span data-stu-id="33ab6-p102">You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).</span></span>

<span data-ttu-id="33ab6-p103">As mudanças nesses dados são armazenadas em uma cópia na memória dessas configurações para a sessão atual do Outlook. Você deve salvar explicitamente todas as configurações de roaming após a atualização para que elas fiquem disponíveis na próxima vez em que o usuário abrir o suplemento, no mesmo ou em qualquer outro dispositivo com suporte.</span><span class="sxs-lookup"><span data-stu-id="33ab6-p103">Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.</span></span>


### <a name="roaming-settings-format"></a><span data-ttu-id="33ab6-116">Formato das configurações de roaming</span><span class="sxs-lookup"><span data-stu-id="33ab6-116">Roaming settings format</span></span>

<span data-ttu-id="33ab6-117">Os dados de um objeto **RoamingSettings** são armazenados como uma cadeia de caracteres serializada JavaScript Object Notation (JSON).</span><span class="sxs-lookup"><span data-stu-id="33ab6-117">The data in a **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string.</span></span> 

<span data-ttu-id="33ab6-118">Abaixo temos um exemplo da estrutura, supondo que existam três configurações de roaming definidas chamadas `add-in_setting_name_0`, `add-in_setting_name_1` e `add-in_setting_name_2`.</span><span class="sxs-lookup"><span data-stu-id="33ab6-118">The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.</span></span>


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a><span data-ttu-id="33ab6-119">Carregar configurações de roaming</span><span class="sxs-lookup"><span data-stu-id="33ab6-119">Loading roaming settings</span></span>

<span data-ttu-id="33ab6-120">Um suplemento de email normalmente carrega configurações de roaming no manipulador de eventos [Office.initialize](/javascript/api/office#office-initialize-reason-).</span><span class="sxs-lookup"><span data-stu-id="33ab6-120">A mail add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office#office-initialize-reason-) event handler.</span></span> <span data-ttu-id="33ab6-121">O exemplo de código JavaScript a seguir mostra como carregar configurações de roaming existentes e obter os valores de duas configurações, **customerName** e **customerBalance**:</span><span class="sxs-lookup"><span data-stu-id="33ab6-121">The following JavaScript code example shows how to load existing roaming settings and get the values of 2 settings, **customerName** and **customerBalance**:</span></span>


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="33ab6-122">Criar ou atribuir uma configuração de roaming</span><span class="sxs-lookup"><span data-stu-id="33ab6-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="33ab6-123">Continuando com o exemplo anterior, a função JavaScript a seguir, `setAddInSetting`, mostra como usar o método [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) para definir uma configuração denominada `cookie` com a data de hoje e manter os dados usando o método [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) para salvar todas as configurações de roaming de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="33ab6-123">Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) method to save all the roaming settings back to the server.</span></span>

<span data-ttu-id="33ab6-124">O `set` método cria a configuração se a configuração ainda não existir e atribui a configuração ao valor especificado.</span><span class="sxs-lookup"><span data-stu-id="33ab6-124">The `set` method creates the setting if the setting does not already exist, and assigns the setting to the specified value.</span></span> <span data-ttu-id="33ab6-125">O `saveAsync` método salva as configurações de roaming de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="33ab6-125">The `saveAsync` method saves roaming settings asynchronously.</span></span> <span data-ttu-id="33ab6-126">Este exemplo de código passa um método de retorno de chamada, `saveMyAddInSettingsCallback` , para `saveAsync` quando a chamada assíncrona é concluída,  `saveMyAddInSettingsCallback` é chamado usando um parâmetro, _AsyncResult_.</span><span class="sxs-lookup"><span data-stu-id="33ab6-126">This code sample passes a callback method, `saveMyAddInSettingsCallback`, to `saveAsync` When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_.</span></span> <span data-ttu-id="33ab6-127">Esse parâmetro é um objeto [AsyncResult](/javascript/api/office/office.asyncresult) que contém o resultado e detalhes sobre a chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="33ab6-127">This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call.</span></span> <span data-ttu-id="33ab6-128">Você pode usar o parâmetro opcional _userContext_ para passar as informações de estado de chamada assíncrona à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="33ab6-128">You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.</span></span>

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="33ab6-129">Remover uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="33ab6-129">Removing a roaming setting</span></span>

<span data-ttu-id="33ab6-130">Estendendo também os exemplos anteriores, a função JavaScript a seguir, `removeAddInSetting`, mostra como usar o método [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) para remover a definição `cookie` e salvar todas as configurações de roaming de volta no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="33ab6-130">Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a><span data-ttu-id="33ab6-131">Dados personalizados por item em uma caixa de correio: propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="33ab6-131">Custom data per item in a mailbox: custom properties</span></span>

<span data-ttu-id="33ab6-p106">Você pode especificar dados específicos de um item na caixa de correio do usuário usando o objeto [CustomProperties](/javascript/api/outlook/office.CustomProperties). Por exemplo, seu suplemento de e-mail poderia categorizar determinadas mensagens e anotar a categoria usando uma propriedade personalizada `messageCategory`. Ou, se seu suplemento de e-mail cria compromissos de sugestões de reunião em uma mensagem, você pode usar uma propriedade personalizada para controlar cada um desses compromissos. Isso garante que se o usuário abrir a mensagem novamente, o suplemento de e-mail não se oferecerá para criar o compromisso uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="33ab6-p106">You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.CustomProperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.</span></span>

<span data-ttu-id="33ab6-p107">Semelhante às configurações de roaming, as mudanças nas propriedades personalizadas são armazenadas em cópias na memória das propriedades para a sessão atual do Outlook. Para garantir que essas propriedades personalizadas estarão disponíveis na próxima sessão, use[CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span><span class="sxs-lookup"><span data-stu-id="33ab6-p107">Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span></span>

<span data-ttu-id="33ab6-138">Essas propriedades personalizadas de suplemento específicas do item podem ser acessadas apenas usando o `CustomProperties` objeto.</span><span class="sxs-lookup"><span data-stu-id="33ab6-138">These add-in-specific, item-specific custom properties can only be accessed by using the `CustomProperties` object.</span></span> <span data-ttu-id="33ab6-139">Essas propriedades são diferentes das propriedades personalizadas, baseadas em MAPI ( [UserProperties](/office/vba/api/Outlook.UserProperties) ) no modelo de objeto do Outlook e de propriedades estendidas no EWS (serviços Web do Exchange).</span><span class="sxs-lookup"><span data-stu-id="33ab6-139">These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS).</span></span> <span data-ttu-id="33ab6-140">Você não pode acessar diretamente `CustomProperties` usando o modelo de objeto do Outlook, EWS ou REST.</span><span class="sxs-lookup"><span data-stu-id="33ab6-140">You cannot directly access `CustomProperties` by using the Outlook object model, EWS, or REST.</span></span> <span data-ttu-id="33ab6-141">Para saber como acessar `CustomProperties` o usando EWS ou REST, confira a seção [obter propriedades personalizadas usando EWS ou REST](#get-custom-properties-using-ews-or-rest).</span><span class="sxs-lookup"><span data-stu-id="33ab6-141">To learn how to access `CustomProperties` using EWS or REST, see the section [Get custom properties using EWS or REST](#get-custom-properties-using-ews-or-rest).</span></span>

### <a name="using-custom-properties"></a><span data-ttu-id="33ab6-142">Usar propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="33ab6-142">Using custom properties</span></span>

<span data-ttu-id="33ab6-143">Antes de poder usar propriedades personalizadas, você precisa carregá-las chamando o método [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods).</span><span class="sxs-lookup"><span data-stu-id="33ab6-143">Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="33ab6-144">Após ter criado o conjunto de propriedades, você poderá usar os métodos [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) e [get](/javascript/api/outlook/office.CustomProperties) para adicionar e recuperar propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-144">After you have created the property bag, you can use the [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) and [get](/javascript/api/outlook/office.CustomProperties) methods to add and retrieve custom properties.</span></span> <span data-ttu-id="33ab6-145">Você deve usar o [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) método para salvar as alterações feitas no conjunto de propriedades.</span><span class="sxs-lookup"><span data-stu-id="33ab6-145">You must use the [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) method to save any changes that you make to the property bag.</span></span>


 > [!NOTE]
 > <span data-ttu-id="33ab6-146">Como o Outlook no Mac não armazena propriedades personalizadas em cache, se a rede do usuário é desativada, os suplementos de e-mail no Outlook no Mac não conseguem acessar suas propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-146">Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook on Mac would not be able to access their custom properties.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="33ab6-147">Exemplo de propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="33ab6-147">Custom properties example</span></span>


<span data-ttu-id="33ab6-p110">O exemplo a seguir mostra um conjunto de métodos simplificado para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar este exemplo como ponto de partida para o seu suplemento que usa propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-p110">The following example shows a simplified set of methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties.</span></span>

<span data-ttu-id="33ab6-150">Este exemplo inclui os seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="33ab6-150">This example includes the following methods:</span></span>


- <span data-ttu-id="33ab6-151">[Office.initialize](/javascript/api/office#office-initialize-reason-): inicializa o suplemento e carrega o conjunto de propriedades personalizadas do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="33ab6-151">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- Initializes the add-in and loads the custom property bag from the Exchange server.</span></span>

- <span data-ttu-id="33ab6-152">**customPropsCallback**: obtém o recipiente de propriedades personalizadas que é retornado do servidor e o salva para uso posterior.</span><span class="sxs-lookup"><span data-stu-id="33ab6-152">**customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.</span></span>

- <span data-ttu-id="33ab6-153">**updateProperty**: define ou atualiza uma propriedade específica e salva a alteração no servidor.</span><span class="sxs-lookup"><span data-stu-id="33ab6-153">**updateProperty** -- Sets or updates a specific property, and then saves the change to the server.</span></span>

- <span data-ttu-id="33ab6-154">**removeProperty**: remove uma propriedade específica do recipiente de propriedades e salva a remoção no servidor.</span><span class="sxs-lookup"><span data-stu-id="33ab6-154">**removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.</span></span>


```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a><span data-ttu-id="33ab6-155">Obtenha propriedades personalizadas usando EWS ou REST</span><span class="sxs-lookup"><span data-stu-id="33ab6-155">Get custom properties using EWS or REST</span></span>

<span data-ttu-id="33ab6-156">Para obter **CustomProperties** usando EWS ou restante, você deverá primeiro determinar o nome do seu MAPI baseado propriedade estendida.</span><span class="sxs-lookup"><span data-stu-id="33ab6-156">To get **CustomProperties** using EWS or REST, you should first determine the name of its MAPI-based extended property.</span></span> <span data-ttu-id="33ab6-157">Você pode obter propriedade da mesma forma que você teria qualquer propriedade com base MAPI estendida.</span><span class="sxs-lookup"><span data-stu-id="33ab6-157">You can then get that property in the same way you would get any MAPI-based extended property.</span></span>

#### <a name="how-custom-properties-are-stored-on-an-item"></a><span data-ttu-id="33ab6-158">Como as propriedades personalizadas são armazenadas em um item</span><span class="sxs-lookup"><span data-stu-id="33ab6-158">How custom properties are stored on an item</span></span>

<span data-ttu-id="33ab6-159">Propriedades personalizadas definidas por um suplemento não são equivalentes normal MAPI com base em Propriedades.</span><span class="sxs-lookup"><span data-stu-id="33ab6-159">Custom properties set by an add-in are not equivalent to normal MAPI-based properties.</span></span> <span data-ttu-id="33ab6-160">As APIs de suplemento serializam todos os suplementos `CustomProperties` como uma carga JSON e, em seguida, os salvam em uma única propriedade estendida com base em MAPI cujo nome é `cecp-<app-guid>` ( `<app-guid>` sua ID do suplemento) e o GUID do conjunto de propriedades é `{00020329-0000-0000-C000-000000000046}` .</span><span class="sxs-lookup"><span data-stu-id="33ab6-160">Add-in APIs serialize all your add-in's `CustomProperties` as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`.</span></span> <span data-ttu-id="33ab6-161">(Para saber mais sobre esse objeto, confira [MS-OXCEXT 2.2.5 propriedades personalizadas do aplicativo de e-mail](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx).) Em seguida, você pode usar EWS ou REST para obter essa propriedade com base MAPI.</span><span class="sxs-lookup"><span data-stu-id="33ab6-161">(For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx).) You can then use EWS or REST to get this MAPI-based property.</span></span>

#### <a name="get-custom-properties-using-ews"></a><span data-ttu-id="33ab6-162">Obtenha propriedades personalizadas usando EWS</span><span class="sxs-lookup"><span data-stu-id="33ab6-162">Get custom properties using EWS</span></span>

<span data-ttu-id="33ab6-163">O suplemento de email pode obter a `CustomProperties` propriedade estendida baseada em MAPI usando a operação EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) .</span><span class="sxs-lookup"><span data-stu-id="33ab6-163">Your mail add-in can get the `CustomProperties` MAPI-based extended property by using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation.</span></span> <span data-ttu-id="33ab6-164">Acesso ao `GetItem` lado do servidor usando um token de retorno de chamada ou no lado do cliente usando o método [Mailbox. makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="33ab6-164">Access `GetItem` on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span> <span data-ttu-id="33ab6-165">Na `GetItem` solicitação, especifique a `CustomProperties` Propriedade baseada em MAPI em seu conjunto de propriedades usando os detalhes fornecidos na seção anterior [como as propriedades personalizadas são armazenadas em um item](#how-custom-properties-are-stored-on-an-item).</span><span class="sxs-lookup"><span data-stu-id="33ab6-165">In the `GetItem` request, specify the `CustomProperties` MAPI-based property in its property set using the details provided in the preceding section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="33ab6-166">O exemplo a seguir mostra como acessar um item e suas propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-166">The following example shows how to get an item and its custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33ab6-167">No exemplo a seguir, substituir `<app-guid>` com ID do suplemento.</span><span class="sxs-lookup"><span data-stu-id="33ab6-167">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

<span data-ttu-id="33ab6-168">Você também pode obter mais propriedades personalizadas se especificar na cadeia de caracteres solicitação como outros [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elementos.</span><span class="sxs-lookup"><span data-stu-id="33ab6-168">You can also get more custom properties if you specify them in the request string as other [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elements.</span></span>

#### <a name="get-custom-properties-using-rest"></a><span data-ttu-id="33ab6-169">Obtenha propriedades personalizadas usando REST</span><span class="sxs-lookup"><span data-stu-id="33ab6-169">Get custom properties using REST</span></span>

<span data-ttu-id="33ab6-170">No seu suplemento, você pode criar sua consulta REST para mensagens e eventos para obter as que já têm propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-170">In your add-in, you can construct your REST query against messages and events to get the ones that already have custom properties.</span></span> <span data-ttu-id="33ab6-171">Em sua consulta, você deve incluir o **CustomProperties** propriedades de MAPI baseados na sua propriedade definida utilizando os detalhes fornecidos na seção [como as propriedades personalizadas são armazenadas em um item](#how-custom-properties-are-stored-on-an-item).</span><span class="sxs-lookup"><span data-stu-id="33ab6-171">In your query, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in the section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="33ab6-172">O exemplo a seguir mostra como obter todos os eventos com as propriedades personalizadas definidos pelo seu suplemento e certifique-se que a resposta inclui o valor da propriedade para que você possa aplicar mais filtragem lógica.</span><span class="sxs-lookup"><span data-stu-id="33ab6-172">The following example shows how to get all events that have any custom properties set by your add-in and ensure that the response includes the value of the property so you can apply further filtering logic.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33ab6-173">No exemplo a seguir substituir `<app-guid>` com ID do suplemento.</span><span class="sxs-lookup"><span data-stu-id="33ab6-173">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

<span data-ttu-id="33ab6-174">Outros exemplos que usam o REST para obter um único valor com base MAPI estendida, confira [Obter singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="33ab6-174">For other examples that use REST to get single-value MAPI-based extended properties, see [Get singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).</span></span>

<span data-ttu-id="33ab6-175">O exemplo a seguir mostra como acessar um item e suas propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-175">The following example shows how to get an item and its custom properties.</span></span> <span data-ttu-id="33ab6-176">Na função retorno de chamada para o `done` método `item.SingleValueExtendedProperties` contém uma lista das propriedades personalizadas solicitadas.</span><span class="sxs-lookup"><span data-stu-id="33ab6-176">In the callback function for the `done` method, `item.SingleValueExtendedProperties` contains a list of the requested custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="33ab6-177">No exemplo a seguir, substituir `<app-guid>` com ID do suplemento.</span><span class="sxs-lookup"><span data-stu-id="33ab6-177">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a><span data-ttu-id="33ab6-178">Confira também</span><span class="sxs-lookup"><span data-stu-id="33ab6-178">See also</span></span>

- [<span data-ttu-id="33ab6-179">Visão geral da propriedade MAPI</span><span class="sxs-lookup"><span data-stu-id="33ab6-179">MAPI Property Overview</span></span>](/office/client-developer/outlook/mapi/mapi-property-overview)
- [<span data-ttu-id="33ab6-180">Visão geral das propriedades do Outlook</span><span class="sxs-lookup"><span data-stu-id="33ab6-180">Outlook Properties Overview</span></span>](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [<span data-ttu-id="33ab6-181">Chamar APIs REST do Outlook de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="33ab6-181">Call Outlook REST APIs from an Outlook add-in</span></span>](use-rest-api.md)
- [<span data-ttu-id="33ab6-182">Chamar serviços Web de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="33ab6-182">Call web services from an Outlook add-in</span></span>](web-services.md)
- [<span data-ttu-id="33ab6-183">Propriedades e propriedades estendidas no EWS no Exchange</span><span class="sxs-lookup"><span data-stu-id="33ab6-183">Properties and extended properties in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [<span data-ttu-id="33ab6-184">Conjuntos de propriedades e formas de resposta no EWS no Exchange</span><span class="sxs-lookup"><span data-stu-id="33ab6-184">Property sets and response shapes in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)
