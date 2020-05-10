---
title: Persistir o estado e as configurações do suplemento
description: Saiba como manter dados nos aplicativos Web de suplemento do Office em execução no ambiente sem estado de um controle de navegador.
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: 0162bc17897cba99f4ce2457cea08d0da70f4341
ms.sourcegitcommit: 7e6faf3dc144400a7b7e5a42adecbbec0bd4602d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/09/2020
ms.locfileid: "44180221"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="d7808-103">Persistir o estado e as configurações do suplemento</span><span class="sxs-lookup"><span data-stu-id="d7808-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="d7808-p101">Essencialmente, os suplementos do Office são aplicativos Web em execução no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:</span><span class="sxs-lookup"><span data-stu-id="d7808-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="d7808-108">Use membros da API JavaScript do Office que armazenam dados como:</span><span class="sxs-lookup"><span data-stu-id="d7808-108">Use members of the Office JavaScript API that store data as either:</span></span>
    -  <span data-ttu-id="d7808-109">Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="d7808-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="d7808-110">XML personalizado armazenado no documento.</span><span class="sxs-lookup"><span data-stu-id="d7808-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="d7808-111">Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="d7808-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="d7808-112">Este artigo se concentra em como usar a API JavaScript do Office para persistir o estado do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d7808-112">This article focuses on how to use the Office JavaScript API to persist add-in state.</span></span> <span data-ttu-id="d7808-113">Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="d7808-113">For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="d7808-114">Persistir o estado e as configurações do suplemento com a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="d7808-114">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="d7808-115">A API JavaScript do Office fornece os objetos [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings)e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado do suplemento nas sessões, conforme descrito na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="d7808-115">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="d7808-116">Em todos os casos, os valores de configurações salvos são associados à [Id](../reference/manifest/id.md) do suplemento que os criou.</span><span class="sxs-lookup"><span data-stu-id="d7808-116">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="d7808-117">**Object**</span><span class="sxs-lookup"><span data-stu-id="d7808-117">**Object**</span></span>|<span data-ttu-id="d7808-118">**Suporte a tipos de suplementos**</span><span class="sxs-lookup"><span data-stu-id="d7808-118">**Add-in type support**</span></span>|<span data-ttu-id="d7808-119">**Local de armazenamento**</span><span class="sxs-lookup"><span data-stu-id="d7808-119">**Storage location**</span></span>|<span data-ttu-id="d7808-120">**Suporte ao host do Office**</span><span class="sxs-lookup"><span data-stu-id="d7808-120">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d7808-121">Configurações</span><span class="sxs-lookup"><span data-stu-id="d7808-121">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="d7808-122">conteúdo e painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d7808-122">content and task pane</span></span>|<span data-ttu-id="d7808-123">O documento, planilha ou apresentação com o qual o suplemento está trabalhando.</span><span class="sxs-lookup"><span data-stu-id="d7808-123">The document, spreadsheet, or presentation the add-in is working with.</span></span> <span data-ttu-id="d7808-124">As configurações de suplemento de conteúdo e de painel de tarefas estão disponíveis para o suplemento que as criou do documento em que foram salvas.</span><span class="sxs-lookup"><span data-stu-id="d7808-124">Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="d7808-p105">**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="d7808-p105">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="d7808-128">Word, Excel ou PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d7808-128">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="d7808-p106">**Observação:** os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para suplementos em execução no Project (bem como outros aplicativos de host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="d7808-p106">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="d7808-132">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d7808-132">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="d7808-133">Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-133">Outlook</span></span>|<span data-ttu-id="d7808-134">A caixa de correio do Exchange Server do usuário onde o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="d7808-134">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="d7808-135">Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem "mover-se" com o usuário e estão disponíveis para o suplemento quando ele estiver sendo executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessar a caixa de correio desse usuário.</span><span class="sxs-lookup"><span data-stu-id="d7808-135">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="d7808-136">As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="d7808-136">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="d7808-137">Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-137">Outlook</span></span>|
|[<span data-ttu-id="d7808-138">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="d7808-138">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="d7808-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-139">Outlook</span></span>|<span data-ttu-id="d7808-p108">A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.</span><span class="sxs-lookup"><span data-stu-id="d7808-p108">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="d7808-142">Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-142">Outlook</span></span>|
|[<span data-ttu-id="d7808-143">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d7808-143">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="d7808-144">painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d7808-144">task pane</span></span>|<span data-ttu-id="d7808-p109">O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas estão disponíveis para o suplemento que as criou por meio do documento em que são salvos.</span><span class="sxs-lookup"><span data-stu-id="d7808-p109">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="d7808-p110">**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="d7808-p110">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="d7808-150">Word (usando a API comum de JavaScript do Office), Excel (usando a API do JavaScript do Excel com host específico)</span><span class="sxs-lookup"><span data-stu-id="d7808-150">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="d7808-151">Os dados de configurações são gerenciados na memória no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="d7808-151">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="d7808-p111">As duas seções a seguir discutem configurações no contexto da API comum de JavaScript do Office. A API JavaScript do Excel com host específico também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](/javascript/api/excel/excel.settingcollection).</span><span class="sxs-lookup"><span data-stu-id="d7808-p111">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="d7808-156">Internamente, os dados no conjunto de propriedades acessados `Settings`com `CustomProperties`o, `RoamingSettings` , ou objetos, são armazenados como um objeto JSON (JavaScript Object Notation) serializado que contém pares de nome/valor.</span><span class="sxs-lookup"><span data-stu-id="d7808-156">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="d7808-157">O nome (chave) para cada valor deve ser um `string`, e o valor armazenado pode ser um JavaScript `string`, `number` `date`, ou `object`, mas não uma **função**.</span><span class="sxs-lookup"><span data-stu-id="d7808-157">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="d7808-158">Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="d7808-158">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="d7808-159">Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d7808-159">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="d7808-160">Durante a sessão, as configurações são gerenciadas totalmente na memória usando os `get`métodos `set`, e `remove` do objeto que corresponde ao tipo de configuração que você está criando (**Settings**, **CustomProperties**ou **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="d7808-160">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>


> [!IMPORTANT]
> <span data-ttu-id="d7808-161">Para persistir quaisquer adições, atualizações ou exclusões feitas durante a sessão atual do suplemento para o local de armazenamento, você deve chamar o `saveAsync` método do objeto correspondente usado para trabalhar com esse tipo de configuração.</span><span class="sxs-lookup"><span data-stu-id="d7808-161">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="d7808-162">Os `get`métodos `set`, e `remove` operam apenas na cópia na memória do recipiente de propriedades de configurações.</span><span class="sxs-lookup"><span data-stu-id="d7808-162">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="d7808-163">Se o suplemento for fechado sem chamadas `saveAsync`, as alterações feitas nas configurações durante essa sessão serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="d7808-163">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="d7808-164">Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d7808-164">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="d7808-165">Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](/javascript/api/office/office.settings) e seus métodos.</span><span class="sxs-lookup"><span data-stu-id="d7808-165">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="d7808-166">O conjunto de propriedades criado com os métodos do `Settings` objeto está disponível somente para a instância do suplemento de conteúdo ou de painel de tarefas que o criou, e apenas do documento no qual ele foi salvo.</span><span class="sxs-lookup"><span data-stu-id="d7808-166">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="d7808-167">O `Settings` objeto é carregado automaticamente como parte do objeto [Document](/javascript/api/office/office.document) e está disponível quando o painel de tarefas ou suplemento de conteúdo é ativado.</span><span class="sxs-lookup"><span data-stu-id="d7808-167">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="d7808-168">Depois que `Document` o objeto é instanciado, você pode acessar `Settings` o objeto com a propriedade [Settings](/javascript/api/office/office.document#settings) do `Document` objeto.</span><span class="sxs-lookup"><span data-stu-id="d7808-168">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="d7808-169">Durante o tempo de vida da sessão, você só pode usar `Settings.get`os `Settings.set`métodos, `Settings.remove` e para ler, gravar ou remover configurações persistentes e estado de suplemento da cópia na memória do recipiente de propriedades.</span><span class="sxs-lookup"><span data-stu-id="d7808-169">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="d7808-170">Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="d7808-170">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="d7808-171">Criar ou atualizar um valor de configuração</span><span class="sxs-lookup"><span data-stu-id="d7808-171">Creating or updating a setting value</span></span>

<span data-ttu-id="d7808-p117">O exemplo de código a seguir mostra como usar o método [Settings.set](/javascript/api/office/office.settings#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro é o _value_ da configuração.</span><span class="sxs-lookup"><span data-stu-id="d7808-p117">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="d7808-175">A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir.</span><span class="sxs-lookup"><span data-stu-id="d7808-175">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="d7808-176">Use o `Settings.saveAsync` método para manter as configurações novas ou atualizadas no documento.</span><span class="sxs-lookup"><span data-stu-id="d7808-176">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="d7808-177">Obter o valor de uma configuração</span><span class="sxs-lookup"><span data-stu-id="d7808-177">Getting the value of a setting</span></span>

<span data-ttu-id="d7808-178">O exemplo a seguir mostra como usar o método [Settings.get](/javascript/api/office/office.settings#get-name-) para obter o valor de uma configuração chamada "themeColor".</span><span class="sxs-lookup"><span data-stu-id="d7808-178">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="d7808-179">O único parâmetro do `get` método é o _nome_ da configuração que diferencia maiúsculas de minúsculas.</span><span class="sxs-lookup"><span data-stu-id="d7808-179">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="d7808-180">O `get` método retorna o valor que foi salvo anteriormente para o _nome_ da configuração que foi passado.</span><span class="sxs-lookup"><span data-stu-id="d7808-180">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="d7808-181">Se a configuração não existir, o método retornará **null**.</span><span class="sxs-lookup"><span data-stu-id="d7808-181">If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="d7808-182">Remover uma configuração</span><span class="sxs-lookup"><span data-stu-id="d7808-182">Removing a setting</span></span>

<span data-ttu-id="d7808-183">O exemplo a seguir mostra como usar o método [Settings.remove](/javascript/api/office/office.settings#remove-name-) para remover uma configuração com o nome "themeColor".</span><span class="sxs-lookup"><span data-stu-id="d7808-183">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="d7808-184">O único parâmetro do `remove` método é o _nome_ da configuração que diferencia maiúsculas de minúsculas.</span><span class="sxs-lookup"><span data-stu-id="d7808-184">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="d7808-185">Nada acontecerá se a configuração não existir.</span><span class="sxs-lookup"><span data-stu-id="d7808-185">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="d7808-186">Use o `Settings.saveAsync` método para persistir a remoção da configuração do documento.</span><span class="sxs-lookup"><span data-stu-id="d7808-186">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="d7808-187">Salvar suas configurações</span><span class="sxs-lookup"><span data-stu-id="d7808-187">Saving your settings</span></span>

<span data-ttu-id="d7808-188">Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) para armazená-lo no documento.</span><span class="sxs-lookup"><span data-stu-id="d7808-188">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="d7808-189">O único parâmetro do `saveAsync` método é _callback_, que é uma função de retorno de chamada com um único parâmetro.</span><span class="sxs-lookup"><span data-stu-id="d7808-189">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span> 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="d7808-190">A função anônima passada para o `saveAsync` método como o parâmetro _callback_ é executada quando a operação é concluída.</span><span class="sxs-lookup"><span data-stu-id="d7808-190">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="d7808-191">O parâmetro _AsyncResult_ do retorno de chamada fornece acesso a `AsyncResult` um objeto que contém o status da operação.</span><span class="sxs-lookup"><span data-stu-id="d7808-191">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="d7808-192">No exemplo, a função verifica a `AsyncResult.status` propriedade para ver se a operação de salvamento foi bem-sucedida ou falhou e, em seguida, exibe o resultado na página do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d7808-192">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="d7808-193">Como salvar XML personalizado no documento</span><span class="sxs-lookup"><span data-stu-id="d7808-193">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="d7808-p125">Esta seção discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. A API JavaScript do Excel com host específico também fornece acesso a partes XML personalizado. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="d7808-p125">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="d7808-198">Há uma opção de armazenamento adicional caso precise armazenar informações que excedem os limites de tamanho do documento Settings ou que tenham um caractere estruturado.</span><span class="sxs-lookup"><span data-stu-id="d7808-198">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="d7808-199">Você pode manter a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação na parte superior desta seção).</span><span class="sxs-lookup"><span data-stu-id="d7808-199">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="d7808-200">No Word, use o objeto [CustomXmlPart](/javascript/api/office/office.customxmlpart) e seus métodos (novamente, consulte a observação acima do Excel).</span><span class="sxs-lookup"><span data-stu-id="d7808-200">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="d7808-201">O código a seguir cria um componente XML personalizado e exibe sua ID e seu conteúdo no divs na página.</span><span class="sxs-lookup"><span data-stu-id="d7808-201">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="d7808-202">Observe que deverá haver um atributo `xmlns` na cadeia de caracteres de XML.</span><span class="sxs-lookup"><span data-stu-id="d7808-202">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

<span data-ttu-id="d7808-p127">Para recuperar uma parte do XML personalizado, use o método [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mas a ID é um GUID gerado quando parte de XML é criada, portanto, não é possível saber ao codificar qual é a ID. Por esse motivo, ao criar uma parte de XML, é uma prática recomendada armazenar imediatamente a ID da parte de XML como uma configuração e usar uma chave fácil de lembrar. O método a seguir mostra como fazer isso. (Mas confira as seções anteriores deste artigo para obter detalhes e as práticas recomendadas ao trabalhar com configurações personalizadas).</span><span class="sxs-lookup"><span data-stu-id="d7808-p127">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

<span data-ttu-id="d7808-207">O código a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configuração.</span><span class="sxs-lookup"><span data-stu-id="d7808-207">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="d7808-208">Como salvar as configurações em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-208">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="d7808-209">Para obter informações sobre como salvar as configurações em um suplemento do Outlook, consulte [gerenciar o estado e as configurações de um suplemento do Outlook](../outlook/manage-state-and-settings-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="d7808-209">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="d7808-210">Confira também</span><span class="sxs-lookup"><span data-stu-id="d7808-210">See also</span></span>

- [<span data-ttu-id="d7808-211">Entendendo a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="d7808-211">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="d7808-212">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-212">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="d7808-213">Gerenciar o estado e as configurações de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7808-213">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="d7808-214">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="d7808-214">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
