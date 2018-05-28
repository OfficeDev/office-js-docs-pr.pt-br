---
title: Persistir o estado e as configura??es do suplemento
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b4d1cdf2ce127d140153b6db02bc9a337a37bb5d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="5b746-102">Persistir o estado e as configura??es do suplemento</span><span class="sxs-lookup"><span data-stu-id="5b746-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="5b746-p101">Essencialmente, os suplementos do Office s?o aplicativos Web em execu??o no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou opera??es entre sess?es de uso do suplemento. Por exemplo, o suplemento pode ter configura??es personalizadas ou outros valores que precisa salvar e recarregar na pr?xima vez em que for inicializado, como o modo de exibi??o preferido ou o local padr?o de um usu?rio. Para fazer isso, voc? pode:</span><span class="sxs-lookup"><span data-stu-id="5b746-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="5b746-107">Usar os membros da API JavaScript para Office que armazena dados como:</span><span class="sxs-lookup"><span data-stu-id="5b746-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="5b746-108">Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="5b746-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="5b746-109">XML personalizado armazenado no documento.</span><span class="sxs-lookup"><span data-stu-id="5b746-109">Custom XML stored in the document.</span></span>
    
- <span data-ttu-id="5b746-110">Usar t?cnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="5b746-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/sessionStorage)).</span></span>
    
<span data-ttu-id="5b746-p102">Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="5b746-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="5b746-113">Persistir o estado e as configura??es do suplemento com a API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="5b746-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="5b746-p103">A API JavaScript para Office fornece os objetos [Settings](https://dev.office.com/reference/add-ins/shared/settings), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) e [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) para salvar o estado do suplemento entre sess?es, conforme descrito na tabela a seguir. Em todos os casos, os valores de configura??es salvos s?o associados ? [Id](https://dev.office.com/reference/add-ins/manifest/id) do suplemento que os criou.</span><span class="sxs-lookup"><span data-stu-id="5b746-p103">The JavaScript API for Office provides the [Settings](https://dev.office.com/reference/add-ins/shared/settings), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings), and [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](https://dev.office.com/reference/add-ins/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="5b746-116">**Objeto**</span><span class="sxs-lookup"><span data-stu-id="5b746-116">**Object**</span></span>|<span data-ttu-id="5b746-117">**Suporte a tipos de suplementos**</span><span class="sxs-lookup"><span data-stu-id="5b746-117">**Add-in type support**</span></span>|<span data-ttu-id="5b746-118">**Local de armazenamento**</span><span class="sxs-lookup"><span data-stu-id="5b746-118">**Storage location**</span></span>|<span data-ttu-id="5b746-119">**Suporte ao host do Office**</span><span class="sxs-lookup"><span data-stu-id="5b746-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5b746-120">Configura??es</span><span class="sxs-lookup"><span data-stu-id="5b746-120">Settings</span></span>](https://dev.office.com/reference/add-ins/shared/settings)|<span data-ttu-id="5b746-121">conte?do e painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5b746-121">content and task pane</span></span>|<span data-ttu-id="5b746-122">O documento, a planilha ou a apresenta??o com o qual o suplemento est? trabalhando. Configura??es de suplementos de conte?do e de painel de tarefas est?o dispon?veis para o suplemento que os criou por meio do documento em que s?o salvos.</span><span class="sxs-lookup"><span data-stu-id="5b746-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="5b746-p104">**Importante:** n?o armazene senhas e outras IIP (informa??es de identifica??o pessoal) confidenciais com o objeto **Settings**. Os dados salvos n?o ficam vis?veis para os usu?rios finais, mas s?o armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Voc? deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necess?rios ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5b746-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="5b746-126">Word, Excel ou PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5b746-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="5b746-p105">**Observa??o:** os suplementos de painel de tarefas para o Project 2013 n?o d?o suporte ? API **Settings** para o armazenamento do estado ou das configura??es do suplemento. No entanto, para suplementos em execu??o no Project (bem como outros aplicativos de host do Office), voc? pode usar t?cnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas t?cnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="5b746-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="5b746-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5b746-130">RoamingSettings</span></span>](https://dev.office.com/reference/add-ins/outlook/RoamingSettings)|<span data-ttu-id="5b746-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b746-131">Outlook</span></span>|<span data-ttu-id="5b746-132">A caixa de correio do servidor Exchange do usu?rio em que o suplemento est? instalado. Como essas configura??es s?o armazenadas na caixa de correio do servidor do usu?rio, elas podem se "mover" com o usu?rio e est?o dispon?veis para o suplemento quando ele ? executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessa a caixa de correio do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5b746-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="5b746-133">As configura??es m?veis de suplementos do Outlook est?o dispon?veis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento est? instalado.</span><span class="sxs-lookup"><span data-stu-id="5b746-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="5b746-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b746-134">Outlook</span></span>|
|[<span data-ttu-id="5b746-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="5b746-135">CustomProperties</span></span>](https://dev.office.com/reference/add-ins/outlook/CustomProperties)|<span data-ttu-id="5b746-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b746-136">Outlook</span></span>|<span data-ttu-id="5b746-p106">A mensagem, o compromisso ou o item de solicita??o de reuni?o com o qual o suplemento est? trabalhando. As propriedades personalizadas de itens de suplementos do Outlook est?o dispon?veis apenas para o suplemento que as criou e apenas por meio do item em que est?o salvas.</span><span class="sxs-lookup"><span data-stu-id="5b746-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="5b746-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b746-139">Outlook</span></span>|
|[<span data-ttu-id="5b746-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5b746-140">customXmlParts</span></span>](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)|<span data-ttu-id="5b746-141">painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5b746-141">task pane</span></span>|<span data-ttu-id="5b746-p107">O documento, planilha ou apresenta??o com o qual o suplemento est? trabalhando. As configura??es de suplementos do painel de tarefas est?o dispon?veis para o suplemento que as criou por meio do documento em que s?o salvos.</span><span class="sxs-lookup"><span data-stu-id="5b746-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="5b746-p108">**Importante:** n?o armazene senhas e outras IIP (informa??es de identifica??o pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos n?o ficam vis?veis para os usu?rios finais, mas s?o armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Voc? deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necess?rios ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5b746-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="5b746-147">Word (usando a API comum de JavaScript do Office), Excel (usando a API do JavaScript do Excel com host espec?fico)</span><span class="sxs-lookup"><span data-stu-id="5b746-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="5b746-148">Os dados de configura??es s?o gerenciados na mem?ria no tempo de execu??o</span><span class="sxs-lookup"><span data-stu-id="5b746-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="5b746-149">As duas se??es a seguir discutem configura??es no contexto da API comum de JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="5b746-149">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="5b746-150">A API JavaScript do Excel com host espec?fico tamb?m fornece acesso ?s configura??es personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5b746-150">The host-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="5b746-151">As APIs do Excel e os padr?es de programa??o s?o um pouco diferentes.</span><span class="sxs-lookup"><span data-stu-id="5b746-151">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="5b746-152">Para saber mais, confira [SettingCollection do Excel](https://dev.office.com/reference/add-ins/excel/settingcollection).</span><span class="sxs-lookup"><span data-stu-id="5b746-152">For more information, see [Excel SettingCollection](https://dev.office.com/reference/add-ins/excel/settingcollection).</span></span>

<span data-ttu-id="5b746-p110">Internamente, os dados no conjunto de propriedades acessado com os objetos **Settings**, **CustomProperties** ou **RoamingSettings** s?o armazenados como um objeto JSON (JavaScript Object Notation) serializado que cont?m pares de nome/valor. O nome (chave) de cada valor deve ser uma **cadeia de caracteres**, e o valor armazenado pode ser uma **cadeia de caracteres**, um **n?mero**, uma **data** ou um **objeto** JavaScript, mas n?o uma **fun??o**.</span><span class="sxs-lookup"><span data-stu-id="5b746-p110">Internally, the data in the property bag accessed with the  **Settings**,  **CustomProperties**, or  **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a **string**, and the stored value can be a JavaScript  **string**,  **number**,  **date**, or  **object**, but not a  **function**.</span></span>

<span data-ttu-id="5b746-155">Este exemplo da estrutura do conjunto de propriedades cont?m tr?s valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="5b746-155">This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="5b746-p111">Depois que o conjunto de propriedades de configura??es ? salvo durante a sess?o anterior do suplemento, ele pode ser carregado quando o suplemento ? inicializado ou a qualquer momento depois disso durante a sess?o atual do suplemento. Durante a sess?o, as configura??es s?o gerenciadas inteiramente na mem?ria usando os m?todos **get**, **set** e **remove** do objeto que corresponde ?s configura??es de tipo que voc? est? criando (**Settings**, **CustomProperties** ou **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="5b746-p111">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the  **get**,  **set**, and  **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**,  **CustomProperties**, or  **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="5b746-p112">Para persistir as adi??es, atualiza??es ou exclus?es feitas durante a sess?o atual do suplemento para o local de armazenamento, voc? deve chamar o m?todo **saveAsync** do objeto correspondente usado para trabalhar com esse tipo de configura??es. Os m?todos **get**, **set** e **remove** operam somente na c?pia na mem?ria do conjunto de propriedades de configura??es. Se o suplemento for fechado sem chamar **saveAsync**, as altera??es feitas nas configura??es durante a sess?o ser?o perdidas.</span><span class="sxs-lookup"><span data-stu-id="5b746-p112">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the  **saveAsync** method of the corresponding object used to work with that kind of settings. The **get**,  **set**, and  **remove** methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="5b746-161">Como salvar o estado e as configura??es do suplemento por documento para suplementos de conte?do e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5b746-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="5b746-p113">Para persistir as configura??es de estado ou personalizadas de um suplemento de conte?do ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings) e seus m?todos. O conjunto de propriedades criado com os m?todos do objeto **Settings** est? dispon?vel apenas para a inst?ncia do suplemento de conte?do ou de painel de tarefas que o criou e apenas por meio do documento no qual ? salvo.</span><span class="sxs-lookup"><span data-stu-id="5b746-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](https://dev.office.com/reference/add-ins/shared/settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="5b746-p114">O objeto **Settings** ? carregado automaticamente como parte do objeto [Document](https://dev.office.com/reference/add-ins/shared/document) e est? dispon?vel quando o suplemento de conte?do ou de painel de tarefas ? ativado. Depois que o objeto **Document** ? instanciado, voc? pode acessar o objeto **Settings** com a propriedade [settings](https://dev.office.com/reference/add-ins/shared/document.settings) do objeto **Document**. Durante o tempo de vida da sess?o, voc? pode simplesmente usar os m?todos **Settings.get**, **Settings.set** e **Settings.remove** para ler, gravar ou remover configura??es persistentes e o estado do suplementos da c?pia na mem?ria do conjunto de propriedades.</span><span class="sxs-lookup"><span data-stu-id="5b746-p114">The  **Settings** object is automatically loaded as part of the [Document](https://dev.office.com/reference/add-ins/shared/document) object, and is available when the task pane or content add-in is activated. After the **Document** object is instantiated, you can access the **Settings** object with the [settings](https://dev.office.com/reference/add-ins/shared/document.settings) property of the **Document** object. During the lifetime of the session, you can just use the **Settings.get**,  **Settings.set**, and  **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="5b746-167">Como os m?todos set e remove operam apenas em rela??o ? c?pia na mem?ria do conjunto de propriedades de configura??es, para salvar configura??es novas ou alteradas no documento ao qual o suplemento est? associado, voc? deve chamar o m?todo [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync).</span><span class="sxs-lookup"><span data-stu-id="5b746-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="5b746-168">Criar ou atualizar um valor de configura??o</span><span class="sxs-lookup"><span data-stu-id="5b746-168">Creating or updating a setting value</span></span>

<span data-ttu-id="5b746-p115">O exemplo de c?digo a seguir mostra como usar o m?todo [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) para criar uma configura??o chamada `'themeColor'` com um valor `'green'`. O primeiro par?metro do m?todo set ? _name_ (Id) da configura??o a ser definida ou criada, que diferencia mai?sculas de min?sculas. O segundo par?metro ? o _value_ da configura??o.</span><span class="sxs-lookup"><span data-stu-id="5b746-p115">The following code example shows how to use the [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="5b746-p116">A configura??o com o nome especificado ? criada se ainda n?o existir, ou seu valor ? atualizado se j? existir. Use o m?todo **Settings.saveAsync** para persistir as configura??es novas ou atualizadas para o documento.</span><span class="sxs-lookup"><span data-stu-id="5b746-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="5b746-174">Obter o valor de uma configura??o</span><span class="sxs-lookup"><span data-stu-id="5b746-174">Getting the value of a setting</span></span>

<span data-ttu-id="5b746-p117">O exemplo a seguir mostra como usar o m?todo [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) para obter o valor de uma configura??o chamada "themeColor". O ?nico par?metro do m?todo **get** ? o _name_ da configura??o (que diferencia mai?sculas de min?sculas).</span><span class="sxs-lookup"><span data-stu-id="5b746-p117">The following example shows how use the [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="5b746-p118">O m?todo **get** retorna o valor que foi salvo anteriormente para a configura??o _name_ que foi passada. Se a configura??o n?o existir, o m?todo retornar? **null**.</span><span class="sxs-lookup"><span data-stu-id="5b746-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="5b746-179">Remover uma configura??o</span><span class="sxs-lookup"><span data-stu-id="5b746-179">Removing a setting</span></span>

<span data-ttu-id="5b746-p119">O exemplo a seguir mostra como usar o m?todo [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) para remover uma configura??o com o nome "themeColor". O ?nico par?metro do m?todo **remove** ? o _name_ da configura??o (que diferencia mai?sculas de min?sculas).</span><span class="sxs-lookup"><span data-stu-id="5b746-p119">The following example shows how to use the [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="5b746-p120">Nada acontecer? se a configura??o n?o existir. Use o m?todo **Settings.saveAsync** para persistir a remo??o da configura??o do documento.</span><span class="sxs-lookup"><span data-stu-id="5b746-p120">Nothing will happen if the setting does not exist. Use the  **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="5b746-184">Salvar suas configura??es</span><span class="sxs-lookup"><span data-stu-id="5b746-184">Saving your settings</span></span>

<span data-ttu-id="5b746-p121">Para salvar adi??es, altera??es ou exclus?es que o suplemento fez na c?pia na mem?ria do conjunto de propriedades de configura??es durante a sess?o atual, voc? deve chamar o m?todo [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) para armazen?-lo no documento. O ?nico par?metro do m?todo **saveAsync** ? _callback_, que ? uma fun??o de retorno de chamada com um ?nico par?metro.</span><span class="sxs-lookup"><span data-stu-id="5b746-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="5b746-p122">A fun??o an?nima passada ao m?todo **saveAsync** como o par?metro _callback_ ? executada quando a opera??o ? conclu?da. O par?metro _asyncResult_ do retorno de chamada fornece acesso a um objeto **AsyncResult** que cont?m o status da opera??o. No exemplo, a fun??o verifica a propriedade **AsyncResult.status** para ver se a opera??o de salvamento teve ?xito ou falhou e exibe o resultado na p?gina do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5b746-p122">The anonymous function passed into the  **saveAsync** method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation. In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="5b746-190">Como salvar XML personalizado no documento</span><span class="sxs-lookup"><span data-stu-id="5b746-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="5b746-191">Esta se??o discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word.</span><span class="sxs-lookup"><span data-stu-id="5b746-191">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="5b746-192">A API JavaScript do Excel com host espec?fico tamb?m fornece acesso a partes XML personalizado.</span><span class="sxs-lookup"><span data-stu-id="5b746-192">The host-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="5b746-193">As APIs do Excel e os padr?es de programa??o s?o um pouco diferentes.</span><span class="sxs-lookup"><span data-stu-id="5b746-193">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="5b746-194">Para saber mais, confira [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="5b746-194">For more information, see [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart).</span></span>

<span data-ttu-id="5b746-195">H? uma op??o de armazenamento adicional caso precise armazenar informa??es que excedem os limites de tamanho do documento Settings ou que tenham um caractere estruturado.</span><span class="sxs-lookup"><span data-stu-id="5b746-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="5b746-196">Voc? pode manter a marca??o XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observa??o na parte superior desta se??o).</span><span class="sxs-lookup"><span data-stu-id="5b746-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="5b746-197">No Word, use o objeto [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) e seus m?todos (novamente, consulte a observa??o acima para o Excel). O c?digo a seguir cria um componente XML personalizado e exibe sua ID e seu conte?do em divs na p?gina.</span><span class="sxs-lookup"><span data-stu-id="5b746-197">In Word, you use the [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) object and its methods (Again, see the note above for Excel.) The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="5b746-198">Dever? haver um atributo `xmlns` na cadeia de caracteres de XML.</span><span class="sxs-lookup"><span data-stu-id="5b746-198">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="5b746-199">Para recuperar uma parte do XML personalizado, use o m?todo [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync), mas a ID ? um GUID gerado quando parte de XML ? criada, portanto, n?o ? poss?vel saber ao codificar qual ? a ID.</span><span class="sxs-lookup"><span data-stu-id="5b746-199">To retrieve a custom XML part, you use the [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is.</span></span> <span data-ttu-id="5b746-200">Por esse motivo, ao criar uma parte de XML, ? uma pr?tica recomendada armazenar imediatamente a ID da parte de XML como uma configura??o e usar uma chave f?cil de lembrar.</span><span class="sxs-lookup"><span data-stu-id="5b746-200">For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key.</span></span> <span data-ttu-id="5b746-201">O m?todo a seguir mostra como fazer isso.</span><span class="sxs-lookup"><span data-stu-id="5b746-201">The following method shows how to do this.</span></span> <span data-ttu-id="5b746-202">(Mas confira as se??es anteriores deste artigo para obter detalhes e as pr?ticas recomendadas ao trabalhar com configura??es personalizadas).</span><span class="sxs-lookup"><span data-stu-id="5b746-202">(But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="5b746-203">O c?digo a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configura??o.</span><span class="sxs-lookup"><span data-stu-id="5b746-203">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="5b746-204">Como salvar configura??es na caixa de correio do usu?rio para suplementos do Outlook como configura??es m?veis</span><span class="sxs-lookup"><span data-stu-id="5b746-204">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="5b746-205">Um suplemento do Outlook pode usar o objeto [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para salvar o estado do suplemento e os dados de configura??es espec?ficos da caixa de correio do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="5b746-205">An Outlook add-in can use the [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="5b746-206">Esses dados s?o acess?veis apenas por esse suplemento do Outlook em nome do usu?rio que est? executando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5b746-206">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="5b746-207">Os dados s?o armazenados na caixa de correio do Exchange Server do usu?rio e podem ser acessados ??quando o usu?rio faz logon em sua conta e executa o suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5b746-207">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="5b746-208">Carregar configura??es m?veis</span><span class="sxs-lookup"><span data-stu-id="5b746-208">Loading roaming settings</span></span>


<span data-ttu-id="5b746-p127">Um suplemento do Outlook normalmente carrega configura??es m?veis no manipulador de eventos [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize). O exemplo de c?digo JavaScript a seguir mostra como carregar configura??es m?veis existentes.</span><span class="sxs-lookup"><span data-stu-id="5b746-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="5b746-211">Criar ou atribuir uma configura??o m?vel</span><span class="sxs-lookup"><span data-stu-id="5b746-211">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="5b746-p128">Continuando com o exemplo anterior, a fun??o `setAppSetting` a seguir mostra como usar o m?todo [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para definir ou atualizar uma configura??o chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configura??es m?veis de volta no Exchange Server com o m?todo [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings).</span><span class="sxs-lookup"><span data-stu-id="5b746-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method.</span></span>


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

<span data-ttu-id="5b746-p129">O m?todo **saveAsync** salva as configura??es m?veis de forma ass?ncrona e utiliza uma fun??o de retorno de chamada opcional. Este exemplo de c?digo passa uma fun??o de retorno de chamada denominada `saveMyAppSettingsCallback` para o m?todo **saveAsync**. Quando a chamada ass?ncrona ? retornada, o par?metro _asyncResult_ da fun??o `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) que voc? pode usar para determinar o ?xito ou a falha da opera??o com a propriedade **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="5b746-p129">The  **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="5b746-217">Remover uma configura??o m?vel</span><span class="sxs-lookup"><span data-stu-id="5b746-217">Removing a roaming setting</span></span>


<span data-ttu-id="5b746-218">Tamb?m estendendo os exemplos anteriores, a fun??o `removeAppSetting` a seguir mostra como usar o m?todo [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para remover a configura??o `cookie` e salvar todas as configura??es m?veis de volta no Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="5b746-218">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="5b746-219">Como salvar configura??es por item para suplementos do Outlook como propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="5b746-219">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="5b746-p130">As propriedades personalizadas permitem que o suplemento do Outlook armazene informa??es sobre um item com o qual est? trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugest?o de reuni?o em uma mensagem, voc? pode usar propriedades personalizadas para armazenar o fato de que a reuni?o foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook n?o se ofere?a para criar novamente o compromisso.</span><span class="sxs-lookup"><span data-stu-id="5b746-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="5b746-p131">Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicita??o de reuni?o espec?fico, voc? deve carregar as propriedades na mem?ria chamando o m?todo [loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) do objeto **Item**. Se propriedades personalizadas j? estiverem definidas para o item atual, elas ser?o carregadas do servidor Exchange nesse momento. Ap?s carregar as propriedades, voc? pode usar os m?todos [set](https://dev.office.com/reference/add-ins/outlook/CustomProperties) e [get](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na mem?ria. Para salvar as altera??es feitas nas propriedades personalizadas do item, voc? deve usar o m?todo [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) para persistir as altera??es no item no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="5b746-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](https://dev.office.com/reference/add-ins/outlook/CustomProperties) and [get](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="5b746-227">Exemplo de propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="5b746-227">Custom properties example</span></span>

<span data-ttu-id="5b746-p132">O exemplo a seguir mostra um conjunto simplificado de fun??es para um suplemento do Outlook que usa propriedades personalizadas. Voc? pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5b746-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="5b746-230">Um suplemento do Outlook que usa essas fun??es recupera as propriedades personalizadas chamando o m?todo **get** na vari?vel `_customProps`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="5b746-230">An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="5b746-231">Este exemplo inclui as seguintes fun??es:</span><span class="sxs-lookup"><span data-stu-id="5b746-231">This example includes the following functions:</span></span>



|<span data-ttu-id="5b746-232">**Nome da fun??o**</span><span class="sxs-lookup"><span data-stu-id="5b746-232">**Function name**</span></span>|<span data-ttu-id="5b746-233">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="5b746-233">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="5b746-234">Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="5b746-234">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="5b746-235">Obt?m as propriedades personalizadas que s?o retornadas do servidor Exchange e as salva para uso posterior.</span><span class="sxs-lookup"><span data-stu-id="5b746-235">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="5b746-236">Define ou atualiza uma propriedade espec?fica e salva a altera??o no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="5b746-236">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="5b746-237">Remove uma propriedade espec?fica e persiste a remo??o no servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="5b746-237">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="5b746-238">Retorno de chamada para chamadas ao m?todo **saveAsync** nas fun??es `updateProperty` e `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="5b746-238">Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="5b746-239">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="5b746-239">See also</span></span>

- [<span data-ttu-id="5b746-240">No??es b?sicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="5b746-240">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="5b746-241">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="5b746-241">Outlook add-ins</span></span>](https://docs.microsoft.com/en-us/outlook/add-ins/)
- [<span data-ttu-id="5b746-242">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="5b746-242">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
