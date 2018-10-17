---
title: Persistência do estado e das configurações do suplemento
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f0f2333e3b4ab7148a86b5aa376598c46155883c
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506109"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="b0e61-102">Persistência do estado e das configurações do suplemento</span><span class="sxs-lookup"><span data-stu-id="b0e61-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="b0e61-p101">Essencialmente, os suplementos do Office são aplicativos Web executados no ambiente sem estado do controle de um navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:</span><span class="sxs-lookup"><span data-stu-id="b0e61-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="b0e61-107">Usar os membros da API JavaScript para Office que armazena dados como:</span><span class="sxs-lookup"><span data-stu-id="b0e61-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="b0e61-108">Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="b0e61-109">XML personalizado armazenado no documento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-109">Custom XML stored in the document.</span></span>
    
- <span data-ttu-id="b0e61-110">Usar técnicas fornecidas pelo controle do navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="b0e61-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>
    
<span data-ttu-id="b0e61-p102">Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="b0e61-113">Persistir o estado e as configurações do suplemento com a API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="b0e61-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="b0e61-p103">A API JavaScript para Office fornece os objetos [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) e [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) para salvar o estado do suplemento entre sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configuração salvos são associados ao [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id?view=office-js) do suplemento que os criou.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p103">The JavaScript API for Office provides the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js), and [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id?view=office-js) of the add-in that created them.</span></span>

|<span data-ttu-id="b0e61-116">**Objeto**</span><span class="sxs-lookup"><span data-stu-id="b0e61-116">**Object**</span></span>|<span data-ttu-id="b0e61-117">**Tipo de suplemento compatível**</span><span class="sxs-lookup"><span data-stu-id="b0e61-117">**Add-in type support**</span></span>|<span data-ttu-id="b0e61-118">**Local de armazenamento**</span><span class="sxs-lookup"><span data-stu-id="b0e61-118">**Storage location**</span></span>|<span data-ttu-id="b0e61-119">**Host do Office compatível**</span><span class="sxs-lookup"><span data-stu-id="b0e61-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b0e61-120">Configurações</span><span class="sxs-lookup"><span data-stu-id="b0e61-120">Settings</span></span>](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js)|<span data-ttu-id="b0e61-121">conteúdo e painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b0e61-121">content and task pane</span></span>|<span data-ttu-id="b0e61-122">O documento, a planilha ou a apresentação com o qual o suplemento está trabalhando. As configurações dos suplementos de conteúdo e de painel de tarefas são disponibilizadas para o suplemento que os criou por meio do documento onde são salvos.</span><span class="sxs-lookup"><span data-stu-id="b0e61-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="b0e61-p104">**Importante:** não armazene senhas e outras informações de identificação pessoal (PII) confidenciais junto com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer PII necessárias ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido de usuário.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="b0e61-126">Word, Excel ou PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b0e61-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="b0e61-p105">**Observação:** os suplementos de painel de tarefas para o Project 2013 não são compatíveis com a API **Settings** para armazenamento de estado e configurações do suplemento. No entanto, para suplementos executados no Project (bem como outros aplicativos host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="b0e61-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b0e61-130">RoamingSettings</span></span>](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js)|<span data-ttu-id="b0e61-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-131">Outlook</span></span>|<span data-ttu-id="b0e61-132">A caixa de correio do Exchange Server do usuário onde o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem acompanhar o usuário e são disponibilizadas para o suplemento quando ele for executado no contexto de qualquer aplicativo host de cliente compatível ou navegador que acessa a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="b0e61-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="b0e61-133">As configurações móveis de um suplemento do Outlook são disponibilizadas apenas para o suplemento que as criou e somente por meio da caixa de correio onde o suplemento está instalado.</span><span class="sxs-lookup"><span data-stu-id="b0e61-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="b0e61-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-134">Outlook</span></span>|
|[<span data-ttu-id="b0e61-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="b0e61-135">CustomProperties</span></span>](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js)|<span data-ttu-id="b0e61-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-136">Outlook</span></span>|<span data-ttu-id="b0e61-p106">A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas dos itens de um suplemento do Outlook são disponibilizadas apenas para o suplemento que as criou e apenas por meio do item onde foram salvas.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="b0e61-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-139">Outlook</span></span>|
|[<span data-ttu-id="b0e61-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b0e61-140">CustomXmlParts</span></span>](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)|<span data-ttu-id="b0e61-141">painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b0e61-141">task pane</span></span>|<span data-ttu-id="b0e61-p107">O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas são disponibilizadas para o suplemento que as criou por meio do documento onde foram salvas.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="b0e61-p108">**Importante:** não armazene senhas e outras informações de identificação pessoal (PII) confidenciais em uma parte XML personalizada. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer PII necessárias ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido de usuário.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="b0e61-147">Word (usando a API JavaScript Comum do Office), Excel (usando a API JavaScript do Excel específica do host)</span><span class="sxs-lookup"><span data-stu-id="b0e61-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="b0e61-148">Os dados de configuração são gerenciados na memória em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="b0e61-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="b0e61-p109">As duas seções seguintes abordam as configurações no contexto da API JavaScript Comum do Office. A API JavaScript do Excel também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para obter mais informações, confira [SettingCollection do Excel](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p109">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).</span></span>

<span data-ttu-id="b0e61-153">Internamente, os dados no recipiente de propriedades acessado com as objetos **Settings**, **CustomProperties** ou **RoamingSettings** são armazenados como objetos JSON (JacaScript Object Notation) serializados contendo pares nome/valor.</span><span class="sxs-lookup"><span data-stu-id="b0e61-153">Internally, the data in the property bag accessed with the **Settings**, **CustomProperties**, or **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="b0e61-154">O nome (chave) de cada valor deve ser uma **sequência de caracteres** e o valor armazenado pode ser uma **sequência de caracteres**, um **número**, uma **data** ou um **objeto** Javascript, mas não uma **função**.</span><span class="sxs-lookup"><span data-stu-id="b0e61-154">The name (key) for each value must be a **string**, and the stored value can be a JavaScript **string**, **number**, **date**, or **object**, but not a **function**.</span></span>

<span data-ttu-id="b0e61-155">Este exemplo de estrutura do recipiente de propriedades contém três valores definidos de **sequência de caracteres** chamados `firstName`, `location` e `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="b0e61-155">This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="b0e61-156">Após o recipiente de propriedades de configuração ter sido salvo na sessão anterior do suplemento, ele pode ser carregado na inicialização do suplemento ou em qualquer outro momento durante a sessão atual do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-156">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="b0e61-157">Durante a sessão, as configurações são totalmente gerenciadas em memória usando os métodos **get**, **set** e **remove** do objeto que corresponde ao tipo de configuração que você criou ( **Settings**, **CustomProperties** ou **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="b0e61-157">During the session, the settings are managed in entirely in memory using the **get**, **set**, and **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**, **CustomProperties**, or **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="b0e61-158">Para persistir quaisquer adições, atualizações ou exclusões realizadas durante a sessão atual do suplemento no local de armazenamento, você deve chamar o método **saveAsync** do objeto correspondente ao que foi usado para trabalhar com esse tipo de configuração.</span><span class="sxs-lookup"><span data-stu-id="b0e61-158">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the **saveAsync** method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="b0e61-159">Os métodos **get**, **set** e **remove** operam somente na cópia em memória do recipiente de propriedades de configuração.</span><span class="sxs-lookup"><span data-stu-id="b0e61-159">The **get**, **set**, and **remove** methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="b0e61-160">Se o seu suplemento for fechado sem chamar **saveAsync**, quaisquer alterações realizadas nas configurações durante a sessão serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="b0e61-160">If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="b0e61-161">Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b0e61-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="b0e61-p113">Para persistir o estado ou as configurações personalizadas de um suplemento de conteúdo ou de painel de tarefas do Word, Excel ou PowerPoint, use o objeto [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) e seus métodos. O recipiente de propriedades criado com os métodos do objeto **Settings** está disponível apenas para a instância do suplemento de conteúdo ou de painel de tarefas que o criou e apenas por meio do documento onde é salvo.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="b0e61-164">O objeto **Settings** é carregado automaticamente como parte do objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) e fica disponível quando suplemento de painel de tarefas ou conteúdo é ativado.</span><span class="sxs-lookup"><span data-stu-id="b0e61-164">The **Settings** object is automatically loaded as part of the [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="b0e61-165">Depois que o objeto **Document** é instanciado, você pode acessar o objeto **Settings** com a propriedade [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) do objeto **Document** .</span><span class="sxs-lookup"><span data-stu-id="b0e61-165">After the **Document** object is instantiated, you can access the **Settings** object with the [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) property of the **Document** object.</span></span> <span data-ttu-id="b0e61-166">Durante o ciclo de vida da sessão, você pode usar apenas os métodos **Settings.get**, **Settings.set** e **Settings.remove** para ler, gravar ou remover as configurações persistentes e o estado do suplemento na cópia em memória do recipiente de propriedades.</span><span class="sxs-lookup"><span data-stu-id="b0e61-166">During the lifetime of the session, you can just use the **Settings.get**, **Settings.set**, and **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="b0e61-167">Como os métodos set e remove operam apenas em relação à cópia em memória do recipiente de propriedades de configuração, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="b0e61-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="b0e61-168">Criação ou atualização de um valor de configuração</span><span class="sxs-lookup"><span data-stu-id="b0e61-168">Creating or updating a setting value</span></span>

<span data-ttu-id="b0e61-p115">O exemplo de código a seguir mostra como usar o método [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro, _value_, é o valor da configuração.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p115">The following code example shows how to use the [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="b0e61-p116">A configuração é criada com o nome especificado se ainda não existir, ou seu valor é atualizado se já existir. Use o método **Settings.saveAsync** para persistir as configurações novas ou atualizadas no documento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="b0e61-174">Obtenção do valor de uma configuração</span><span class="sxs-lookup"><span data-stu-id="b0e61-174">Getting the value of a setting</span></span>

<span data-ttu-id="b0e61-p117">O exemplo a seguir mostra como usar o método [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do método **get** é _name_ , o nome da configuração (que diferencia maiúsculas de minúsculas).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p117">The following example shows how use the [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="b0e61-p118">O método **get** retorna o valor salvo anteriormente da configuração _name_ que foi passada para o método. Se a configuração não existir, o método retornará **null**.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="b0e61-179">Exclusão de uma configuração</span><span class="sxs-lookup"><span data-stu-id="b0e61-179">Removing a setting</span></span>

<span data-ttu-id="b0e61-p119">O exemplo a seguir mostra como usar o método [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) para excluir uma configuração com o nome "themeColor". O único parâmetro do método **remove** é _name_, o nome da configuração (que diferencia maiúsculas de minúsculas).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p119">The following example shows how to use the [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="b0e61-182">Nada acontecerá se a configuração não existir.</span><span class="sxs-lookup"><span data-stu-id="b0e61-182">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="b0e61-183">Use o método **Settings.saveAsync** para persistir a exclusão da configuração no documento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-183">Use the **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="b0e61-184">Gravação das configurações</span><span class="sxs-lookup"><span data-stu-id="b0e61-184">Saving your settings</span></span>

<span data-ttu-id="b0e61-p121">Para salvar adições, alterações ou exclusões que o suplemento fez na cópia em memória do recipiente de propriedades de configuração durante a sessão atual, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) para armazená-lo no documento. O único parâmetro do método **saveAsync** é _callback_, que é uma função de retorno de chamada com um único parâmetro.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="b0e61-187">A função anônima passada para o método **saveAsync** como parâmetro de _retorno de chamada_ será executada quando a operação for concluída.</span><span class="sxs-lookup"><span data-stu-id="b0e61-187">The anonymous function passed into the **saveAsync** method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="b0e61-188">O parâmetro do retorno de chamada  _asyncResult_ dá acesso a um objeto **AsyncResult** que contém o status da operação.</span><span class="sxs-lookup"><span data-stu-id="b0e61-188">The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation.</span></span> <span data-ttu-id="b0e61-189">No exemplo, a função verifica a propriedade **AsyncResult.status** para verificar se a operação de gravação foi bem-sucedida ou falhou e, em seguida, exibe o resultado na página do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0e61-189">In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="b0e61-190">Como salvar XML personalizado no documento</span><span class="sxs-lookup"><span data-stu-id="b0e61-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="b0e61-p123">Esta seção aborda partes XML personalizadas no contexto da API JavaScript comum do Office que é suportada no Word. A API JavaScript do Excel também fornece acesso às partes XML personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para obter mais informações, consulte [CustomXmlPart do Excel](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p123">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).</span></span>

<span data-ttu-id="b0e61-195">Há uma opção de armazenamento adicional caso você precise armazenar informações que excedam os limites de tamanho das configurações do documento ou que possuam um caráter estruturado.</span><span class="sxs-lookup"><span data-stu-id="b0e61-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="b0e61-196">Você pode persistir a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação no início desta seção).</span><span class="sxs-lookup"><span data-stu-id="b0e61-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="b0e61-197">No Word, você pode usar o objeto [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) e seus métodos (novamente, confira a observação acima sobre o Excel).</span><span class="sxs-lookup"><span data-stu-id="b0e61-197">In Word, you use the [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="b0e61-198">O código a seguir cria uma parte XML personalizada, exibe seu ID e, em seguida, seu conteúdo em divs na página.</span><span class="sxs-lookup"><span data-stu-id="b0e61-198">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="b0e61-199">Observe que deve exisitir um atributo `xmlns` na sequência de caracteres XML.</span><span class="sxs-lookup"><span data-stu-id="b0e61-199">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```

<span data-ttu-id="b0e61-p125">Para recuperar uma parte XML personalizada, você deve usar o método [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) , mas o ID é um GUID gerado durante a criação da parte XML , então você não pode referenciá-lo no código. Por esse motivo, uma boa prática é armazenar o ID da parte XML como uma configuração imediatamente após a sua criação e atribuir uma chave fácil de lembrar . O método a seguir mostra como fazer isso. (Mas consulte as seções anteriores deste artigo para obter detalhes e práticas recomendadas ao trabalhar com configurações personalizadas).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p125">To retrieve a custom XML part, you use the [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="b0e61-204">O código a seguir mostra como recuperar a parte do XML obtendo  o seu ID a partir de uma configuração.</span><span class="sxs-lookup"><span data-stu-id="b0e61-204">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID'));
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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="b0e61-205">Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis</span><span class="sxs-lookup"><span data-stu-id="b0e61-205">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="b0e61-p126">Um suplemento do Outlook pode usar o objeto [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) para salvar o estado do suplemento e dados de configurações específicas na caixa de correio do usuário. Esses dados ficam acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento. Os dados são armazenados na caixa de correio do usuário do Exchange Server e pode ser acessados quando o usuário faz logon em sua conta e executa o suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p126">An Outlook add-in can use the [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) object to save add-in state and settings data that is specific to the user's mailbox. This data is accessible only by that Outlook add-in on behalf of the user running the add-in. The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="b0e61-209">Carga de configurações móveis</span><span class="sxs-lookup"><span data-stu-id="b0e61-209">Loading roaming settings</span></span>


<span data-ttu-id="b0e61-p127">Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="b0e61-212">Criação ou atribuição de uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="b0e61-212">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="b0e61-p128">Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="b0e61-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-) method.</span></span>


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

<span data-ttu-id="b0e61-p129">O método **saveAsync** salva as configurações móveis de forma assíncrona e recebe uma função de retorno de chamada opcional. Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**. Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p129">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="b0e61-218">Exclusão de uma configuração móvel</span><span class="sxs-lookup"><span data-stu-id="b0e61-218">Removing a roaming setting</span></span>


<span data-ttu-id="b0e61-219">Ainda estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) para excluir a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b0e61-219">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="b0e61-220">Como salvar configurações por item como propriedades personalizadas em suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-220">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="b0e61-p130">As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook cria um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="b0e61-p131">Para poder usar propriedades personalizadas em um item específico como uma mensagem, um compromisso ou uma solicitação de reunião, você deve carregar as propriedades em memória chamando o método [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) e [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) do objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades em memória. Para salvar quaisquer alterações realizadas nas propriedades personalizadas do item, você deve usar o método [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) para persistir as alterações no item no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) and [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="b0e61-228">Exemplo de propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="b0e61-228">Custom properties example</span></span>

<span data-ttu-id="b0e61-p132">O exemplo a seguir mostra um conjunto simplificado de funções de um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o seu suplemento do Outlook que usa propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b0e61-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="b0e61-231">Um suplemento do Outlook que usa essas funções recupera quaisquer propriedades personalizadas chamando o método **get** na variável `_customProps`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="b0e61-231">An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="b0e61-232">Este exemplo inclui as seguintes funções:</span><span class="sxs-lookup"><span data-stu-id="b0e61-232">This example includes the following functions:</span></span>



|<span data-ttu-id="b0e61-233">**Nome da função**</span><span class="sxs-lookup"><span data-stu-id="b0e61-233">**Function name**</span></span>|<span data-ttu-id="b0e61-234">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="b0e61-234">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="b0e61-235">Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b0e61-235">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="b0e61-236">Obtém as propriedades personalizadas retornadas do Exchange Server e as salva para uso posterior.</span><span class="sxs-lookup"><span data-stu-id="b0e61-236">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="b0e61-237">Define ou atualiza uma propriedade específica e salva a alteração no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b0e61-237">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="b0e61-238">Remove uma propriedade específica e persiste a remoção no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="b0e61-238">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="b0e61-239">Retorno de chamada para chamadas ao método **saveAsync** nas funções `updateProperty` e `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="b0e61-239">Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="b0e61-240">Confira também</span><span class="sxs-lookup"><span data-stu-id="b0e61-240">See also</span></span>

- [<span data-ttu-id="b0e61-241">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="b0e61-241">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="b0e61-242">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="b0e61-242">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="b0e61-243">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="b0e61-243">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
