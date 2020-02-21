---
title: Persistir o estado e as configurações do suplemento
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 69fc0b1316a1a4eb0dfe0ebea01ffdbfe88dcd8c
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163504"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="4dac3-102">Persistir o estado e as configurações do suplemento</span><span class="sxs-lookup"><span data-stu-id="4dac3-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="4dac3-p101">Essencialmente, os suplementos do Office são aplicativos Web em execução no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:</span><span class="sxs-lookup"><span data-stu-id="4dac3-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="4dac3-107">Usar os membros da API JavaScript para Office que armazena dados como:</span><span class="sxs-lookup"><span data-stu-id="4dac3-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="4dac3-108">Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="4dac3-109">XML personalizado armazenado no documento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-109">Custom XML stored in the document.</span></span>

- <span data-ttu-id="4dac3-110">Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="4dac3-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="4dac3-p102">Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="4dac3-113">Persistir o estado e as configurações do suplemento com a API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="4dac3-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="4dac3-p103">A API JavaScript para Office fornece os objetos [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado do suplemento entre sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](/office/dev/add-ins/reference/manifest/id) do suplemento que os criou.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p103">The JavaScript API for Office provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](/office/dev/add-ins/reference/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="4dac3-116">**Object**</span><span class="sxs-lookup"><span data-stu-id="4dac3-116">**Object**</span></span>|<span data-ttu-id="4dac3-117">**Suporte a tipos de suplementos**</span><span class="sxs-lookup"><span data-stu-id="4dac3-117">**Add-in type support**</span></span>|<span data-ttu-id="4dac3-118">**Local de armazenamento**</span><span class="sxs-lookup"><span data-stu-id="4dac3-118">**Storage location**</span></span>|<span data-ttu-id="4dac3-119">**Suporte ao host do Office**</span><span class="sxs-lookup"><span data-stu-id="4dac3-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4dac3-120">Configurações</span><span class="sxs-lookup"><span data-stu-id="4dac3-120">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="4dac3-121">conteúdo e painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4dac3-121">content and task pane</span></span>|<span data-ttu-id="4dac3-122">O documento, a planilha ou a apresentação com o qual o suplemento está trabalhando. Configurações de suplementos de conteúdo e de painel de tarefas estão disponíveis para o suplemento que os criou por meio do documento em que são salvos.</span><span class="sxs-lookup"><span data-stu-id="4dac3-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="4dac3-p104">**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="4dac3-126">Word, Excel ou PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4dac3-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="4dac3-p105">**Observação:** os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para suplementos em execução no Project (bem como outros aplicativos de host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="4dac3-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4dac3-130">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="4dac3-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="4dac3-131">Outlook</span></span>|<span data-ttu-id="4dac3-132">A caixa de correio do servidor Exchange do usuário em que o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem se "mover" com o usuário e estão disponíveis para o suplemento quando ele é executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessa a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4dac3-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="4dac3-133">As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="4dac3-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="4dac3-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="4dac3-134">Outlook</span></span>|
|[<span data-ttu-id="4dac3-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="4dac3-135">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="4dac3-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="4dac3-136">Outlook</span></span>|<span data-ttu-id="4dac3-p106">A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="4dac3-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="4dac3-139">Outlook</span></span>|
|[<span data-ttu-id="4dac3-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4dac3-140">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="4dac3-141">painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4dac3-141">task pane</span></span>|<span data-ttu-id="4dac3-p107">O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas estão disponíveis para o suplemento que as criou por meio do documento em que são salvos.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="4dac3-p108">**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="4dac3-147">Word (usando a API comum de JavaScript do Office), Excel (usando a API do JavaScript do Excel com host específico)</span><span class="sxs-lookup"><span data-stu-id="4dac3-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="4dac3-148">Os dados de configurações são gerenciados na memória no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="4dac3-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="4dac3-p109">As duas seções a seguir discutem configurações no contexto da API comum de JavaScript do Office. A API JavaScript do Excel com host específico também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](/javascript/api/excel/excel.settingcollection).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p109">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="4dac3-153">Internamente, os dados no conjunto de propriedades acessados com os objetos **Configurações**, **CustomProperties** ou **RoamingSettings** são armazenados como um objeto JSON (JavaScript Object Notation) serializado que contém pares de nome/valor.</span><span class="sxs-lookup"><span data-stu-id="4dac3-153">Internally, the data in the property bag accessed with the **Settings**, **CustomProperties**, or **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="4dac3-154">O nome (chave) de cada valor deve ser uma **cadeia**, e o valor armazenado pode ser uma **cadeia**, **um número**, **uma data**, ou **objeto** JavaScript, mas não uma **função**.</span><span class="sxs-lookup"><span data-stu-id="4dac3-154">The name (key) for each value must be a **string**, and the stored value can be a JavaScript **string**, **number**, **date**, or **object**, but not a **function**.</span></span>

<span data-ttu-id="4dac3-155">Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="4dac3-155">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="4dac3-156">Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-156">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="4dac3-157">Durante a sessão, as configurações são gerenciadas inteiramente na memória usando os métodos **obter**, **configurar** e **remover** do objeto que corresponde às configurações de tipo que você está criando (**Definições**, **CustomProperties** ou **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="4dac3-157">During the session, the settings are managed in entirely in memory using the **get**, **set**, and **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**, **CustomProperties**, or **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="4dac3-158">Para persistir as adições, atualizações ou exclusões feitas durante a sessão atual do suplemento para o local de armazenamento, você deve chamar o método **saveAsync** do objeto correspondente usado para trabalhar com esse tipo de configurações.</span><span class="sxs-lookup"><span data-stu-id="4dac3-158">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the **saveAsync** method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="4dac3-159">Os métodos **obter**, **definir**, e**remover** operam somente na cópia na memória do conjunto de propriedades de configurações.</span><span class="sxs-lookup"><span data-stu-id="4dac3-159">The **get**, **set**, and **remove** methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="4dac3-160">Se o suplemento for fechado sem chamar **saveAsync**, as alterações feitas nas configurações durante a sessão serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="4dac3-160">If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="4dac3-161">Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4dac3-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="4dac3-p113">Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](/javascript/api/office/office.settings) e seus métodos. O conjunto de propriedades criado com os métodos do objeto **Settings** está disponível apenas para a instância do suplemento de conteúdo ou de painel de tarefas que o criou e apenas por meio do documento no qual é salvo.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="4dac3-164">O objeto **Configurações** é carregado automaticamente como parte do objeto [Documento](/javascript/api/office/office.document) e está disponível quando o suplemento de conteúdo ou de painel de tarefas é ativado.</span><span class="sxs-lookup"><span data-stu-id="4dac3-164">The **Settings** object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="4dac3-165">Depois que o objeto **Documento** é instanciado, você pode acessar o objeto **Configurações** com a propriedade [configurações](/javascript/api/office/office.document#settings) do objeto **Documento**.</span><span class="sxs-lookup"><span data-stu-id="4dac3-165">After the **Document** object is instantiated, you can access the **Settings** object with the [settings](/javascript/api/office/office.document#settings) property of the **Document** object.</span></span> <span data-ttu-id="4dac3-166">Durante o tempo de vida da sessão, você pode simplesmente usar os métodos **Settings.get**, **Settings.set**, e **Settings.remove** para ler, gravar ou remover configurações persistentes e o estado do suplementos da cópia na memória do conjunto de propriedades.</span><span class="sxs-lookup"><span data-stu-id="4dac3-166">During the lifetime of the session, you can just use the **Settings.get**, **Settings.set**, and **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="4dac3-167">Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="4dac3-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="4dac3-168">Criar ou atualizar um valor de configuração</span><span class="sxs-lookup"><span data-stu-id="4dac3-168">Creating or updating a setting value</span></span>

<span data-ttu-id="4dac3-p115">O exemplo de código a seguir mostra como usar o método [Settings.set](/javascript/api/office/office.settings#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro é o _value_ da configuração.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p115">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="4dac3-p116">A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir. Use o método **Settings.saveAsync** para persistir as configurações novas ou atualizadas para o documento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="4dac3-174">Obter o valor de uma configuração</span><span class="sxs-lookup"><span data-stu-id="4dac3-174">Getting the value of a setting</span></span>

<span data-ttu-id="4dac3-p117">O exemplo a seguir mostra como usar o método [Settings.get](/javascript/api/office/office.settings#get-name-) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do método **get** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p117">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="4dac3-p118">O método **get** retorna o valor que foi salvo anteriormente para a configuração _name_ que foi passada. Se a configuração não existir, o método retornará **null**.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="4dac3-179">Remover uma configuração</span><span class="sxs-lookup"><span data-stu-id="4dac3-179">Removing a setting</span></span>

<span data-ttu-id="4dac3-p119">O exemplo a seguir mostra como usar o método [Settings.remove](/javascript/api/office/office.settings#remove-name-) para remover uma configuração com o nome "themeColor". O único parâmetro do método **remove** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p119">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="4dac3-182">Nada acontecerá se a configuração não existir.</span><span class="sxs-lookup"><span data-stu-id="4dac3-182">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="4dac3-183">Use o método **Settings.saveAsync** para persistir a remoção da configuração do documento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-183">Use the **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="4dac3-184">Salvar suas configurações</span><span class="sxs-lookup"><span data-stu-id="4dac3-184">Saving your settings</span></span>

<span data-ttu-id="4dac3-p121">Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) para armazená-lo no documento. O único parâmetro do método **saveAsync** é _callback_, que é uma função de retorno de chamada com um único parâmetro.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="4dac3-187">A função anônima passada ao método **saveAsync** como o parâmetro _callback_ é executada quando a operação é concluída.</span><span class="sxs-lookup"><span data-stu-id="4dac3-187">The anonymous function passed into the **saveAsync** method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="4dac3-188">O parâmetro _asyncResult_ do retorno de chamada fornece acesso a um objeto **AsyncResult** que contém o status da operação.</span><span class="sxs-lookup"><span data-stu-id="4dac3-188">The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation.</span></span> <span data-ttu-id="4dac3-189">No exemplo, a função verifica a propriedade **AsyncResult.status** para ver se a operação de salvamento teve êxito ou falhou e exibe o resultado na página do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-189">In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="4dac3-190">Como salvar XML personalizado no documento</span><span class="sxs-lookup"><span data-stu-id="4dac3-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="4dac3-p123">Esta seção discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. A API JavaScript do Excel com host específico também fornece acesso a partes XML personalizado. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p123">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="4dac3-195">Há uma opção de armazenamento adicional caso precise armazenar informações que excedem os limites de tamanho do documento Settings ou que tenham um caractere estruturado.</span><span class="sxs-lookup"><span data-stu-id="4dac3-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="4dac3-196">Você pode manter a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação na parte superior desta seção).</span><span class="sxs-lookup"><span data-stu-id="4dac3-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="4dac3-197">No Word, use o objeto [CustomXmlPart](/javascript/api/office/office.customxmlpart) e seus métodos (novamente, consulte a observação acima do Excel).</span><span class="sxs-lookup"><span data-stu-id="4dac3-197">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="4dac3-198">O código a seguir cria um componente XML personalizado e exibe sua ID e seu conteúdo no divs na página.</span><span class="sxs-lookup"><span data-stu-id="4dac3-198">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="4dac3-199">Observe que deverá haver um atributo `xmlns` na cadeia de caracteres de XML.</span><span class="sxs-lookup"><span data-stu-id="4dac3-199">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="4dac3-p125">Para recuperar uma parte do XML personalizado, use o método [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mas a ID é um GUID gerado quando parte de XML é criada, portanto, não é possível saber ao codificar qual é a ID. Por esse motivo, ao criar uma parte de XML, é uma prática recomendada armazenar imediatamente a ID da parte de XML como uma configuração e usar uma chave fácil de lembrar. O método a seguir mostra como fazer isso. (Mas confira as seções anteriores deste artigo para obter detalhes e as práticas recomendadas ao trabalhar com configurações personalizadas).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p125">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="4dac3-204">O código a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configuração.</span><span class="sxs-lookup"><span data-stu-id="4dac3-204">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="4dac3-205">Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis</span><span class="sxs-lookup"><span data-stu-id="4dac3-205">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="4dac3-206">Um suplemento do Outlook pode usar o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) para salvar o estado e os dados de configurações do suplemento específico da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4dac3-206">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="4dac3-207">Esses dados são acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento.</span><span class="sxs-lookup"><span data-stu-id="4dac3-207">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="4dac3-208">Os dados são armazenados na caixa de correio do usuário do Exchange Server e ficam acessíveis quando esse usuário faz logon em sua conta e executa o suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4dac3-208">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="4dac3-209">Carregar configurações de roaming</span><span class="sxs-lookup"><span data-stu-id="4dac3-209">Loading roaming settings</span></span>


<span data-ttu-id="4dac3-p127">Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](/javascript/api/office). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="4dac3-212">Criar ou atribuir uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="4dac3-212">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="4dac3-p128">Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="4dac3-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>


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

<span data-ttu-id="4dac3-215">O método **saveAsync** salva as configurações móveis de forma assíncrona e utiliza uma função de retorno de chamada opcional.</span><span class="sxs-lookup"><span data-stu-id="4dac3-215">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="4dac3-216">Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="4dac3-216">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="4dac3-217">Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](/javascript/api/outlook) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="4dac3-217">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/outlook) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="4dac3-218">Remover uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="4dac3-218">Removing a roaming setting</span></span>


<span data-ttu-id="4dac3-219">Também estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) para remover a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4dac3-219">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="4dac3-220">Como salvar configurações por item para suplementos do Outlook como propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="4dac3-220">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="4dac3-p130">As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="4dac3-p131">Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicitação de reunião específico, você deve carregar as propriedades na memória chamando o método [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](/javascript/api/outlook/office.customproperties#set-name--value-) e [get](/javascript/api/outlook/office.roamingsettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na memória. Para salvar as alterações feitas nas propriedades personalizadas do item, você deve usar o método [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) para persistir as alterações no item no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="4dac3-228">Exemplo de propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="4dac3-228">Custom properties example</span></span>

<span data-ttu-id="4dac3-p132">O exemplo a seguir mostra um conjunto simplificado de funções para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4dac3-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="4dac3-231">Um suplemento do Outlook que usa essas funções recupera as propriedades personalizadas chamando o método **obter** na variável `_customProps`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="4dac3-231">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="4dac3-232">Este exemplo inclui as seguintes funções:</span><span class="sxs-lookup"><span data-stu-id="4dac3-232">This example includes the following functions:</span></span>



|<span data-ttu-id="4dac3-233">**Nome da função**</span><span class="sxs-lookup"><span data-stu-id="4dac3-233">**Function name**</span></span>|<span data-ttu-id="4dac3-234">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="4dac3-234">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="4dac3-235">Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="4dac3-235">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="4dac3-236">Obtém as propriedades personalizadas que são retornadas do servidor Exchange e as salva para uso posterior.</span><span class="sxs-lookup"><span data-stu-id="4dac3-236">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="4dac3-237">Define ou atualiza uma propriedade específica e salva a alteração no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="4dac3-237">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="4dac3-238">Remove uma propriedade específica e persiste a remoção no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="4dac3-238">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="4dac3-239">Retorno de chamada para chamadas ao método **saveAsync** nas funções `updateProperty` e `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="4dac3-239">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="4dac3-240">Confira também</span><span class="sxs-lookup"><span data-stu-id="4dac3-240">See also</span></span>

- [<span data-ttu-id="4dac3-241">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="4dac3-241">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="4dac3-242">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="4dac3-242">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="4dac3-243">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="4dac3-243">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
