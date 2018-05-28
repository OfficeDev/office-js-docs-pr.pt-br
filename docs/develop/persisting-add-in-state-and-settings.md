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
# <a name="persisting-add-in-state-and-settings"></a>Persistir o estado e as configura??es do suplemento

Essencialmente, os suplementos do Office s?o aplicativos Web em execu??o no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou opera??es entre sess?es de uso do suplemento. Por exemplo, o suplemento pode ter configura??es personalizadas ou outros valores que precisa salvar e recarregar na pr?xima vez em que for inicializado, como o modo de exibi??o preferido ou o local padr?o de um usu?rio. Para fazer isso, voc? pode:

- Usar os membros da API JavaScript para Office que armazena dados como:
    -  Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.
    -  XML personalizado armazenado no documento.
    
- Usar t?cnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/sessionStorage)).
    
Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistir o estado e as configura??es do suplemento com a API JavaScript para Office

A API JavaScript para Office fornece os objetos [Settings](https://dev.office.com/reference/add-ins/shared/settings), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) e [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) para salvar o estado do suplemento entre sess?es, conforme descrito na tabela a seguir. Em todos os casos, os valores de configura??es salvos s?o associados ? [Id](https://dev.office.com/reference/add-ins/manifest/id) do suplemento que os criou.

|**Objeto**|**Suporte a tipos de suplementos**|**Local de armazenamento**|**Suporte ao host do Office**|
|:-----|:-----|:-----|:-----|
|[Configura??es](https://dev.office.com/reference/add-ins/shared/settings)|conte?do e painel de tarefas|O documento, a planilha ou a apresenta??o com o qual o suplemento est? trabalhando. Configura??es de suplementos de conte?do e de painel de tarefas est?o dispon?veis para o suplemento que os criou por meio do documento em que s?o salvos.<br/><br/>**Importante:** n?o armazene senhas e outras IIP (informa??es de identifica??o pessoal) confidenciais com o objeto **Settings**. Os dados salvos n?o ficam vis?veis para os usu?rios finais, mas s?o armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Voc? deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necess?rios ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usu?rio.|Word, Excel ou PowerPoint<br/><br/> **Observa??o:** os suplementos de painel de tarefas para o Project 2013 n?o d?o suporte ? API **Settings** para o armazenamento do estado ou das configura??es do suplemento. No entanto, para suplementos em execu??o no Project (bem como outros aplicativos de host do Office), voc? pode usar t?cnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas t?cnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings)|Outlook|A caixa de correio do servidor Exchange do usu?rio em que o suplemento est? instalado. Como essas configura??es s?o armazenadas na caixa de correio do servidor do usu?rio, elas podem se "mover" com o usu?rio e est?o dispon?veis para o suplemento quando ele ? executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessa a caixa de correio do usu?rio.<br/><br/> As configura??es m?veis de suplementos do Outlook est?o dispon?veis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento est? instalado.|Outlook|
|[CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties)|Outlook|A mensagem, o compromisso ou o item de solicita??o de reuni?o com o qual o suplemento est? trabalhando. As propriedades personalizadas de itens de suplementos do Outlook est?o dispon?veis apenas para o suplemento que as criou e apenas por meio do item em que est?o salvas.|Outlook|
|[CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)|painel de tarefas|O documento, planilha ou apresenta??o com o qual o suplemento est? trabalhando. As configura??es de suplementos do painel de tarefas est?o dispon?veis para o suplemento que as criou por meio do documento em que s?o salvos.<br/><br/>**Importante:** n?o armazene senhas e outras IIP (informa??es de identifica??o pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos n?o ficam vis?veis para os usu?rios finais, mas s?o armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Voc? deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necess?rios ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usu?rio.|Word (usando a API comum de JavaScript do Office), Excel (usando a API do JavaScript do Excel com host espec?fico)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configura??es s?o gerenciados na mem?ria no tempo de execu??o

> [!NOTE]
> As duas se??es a seguir discutem configura??es no contexto da API comum de JavaScript do Office. A API JavaScript do Excel com host espec?fico tamb?m fornece acesso ?s configura??es personalizadas. As APIs do Excel e os padr?es de programa??o s?o um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](https://dev.office.com/reference/add-ins/excel/settingcollection).

Internamente, os dados no conjunto de propriedades acessado com os objetos **Settings**, **CustomProperties** ou **RoamingSettings** s?o armazenados como um objeto JSON (JavaScript Object Notation) serializado que cont?m pares de nome/valor. O nome (chave) de cada valor deve ser uma **cadeia de caracteres**, e o valor armazenado pode ser uma **cadeia de caracteres**, um **n?mero**, uma **data** ou um **objeto** JavaScript, mas n?o uma **fun??o**.

Este exemplo da estrutura do conjunto de propriedades cont?m tr?s valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Depois que o conjunto de propriedades de configura??es ? salvo durante a sess?o anterior do suplemento, ele pode ser carregado quando o suplemento ? inicializado ou a qualquer momento depois disso durante a sess?o atual do suplemento. Durante a sess?o, as configura??es s?o gerenciadas inteiramente na mem?ria usando os m?todos **get**, **set** e **remove** do objeto que corresponde ?s configura??es de tipo que voc? est? criando (**Settings**, **CustomProperties** ou **RoamingSettings**). 


> [!IMPORTANT]
> Para persistir as adi??es, atualiza??es ou exclus?es feitas durante a sess?o atual do suplemento para o local de armazenamento, voc? deve chamar o m?todo **saveAsync** do objeto correspondente usado para trabalhar com esse tipo de configura??es. Os m?todos **get**, **set** e **remove** operam somente na c?pia na mem?ria do conjunto de propriedades de configura??es. Se o suplemento for fechado sem chamar **saveAsync**, as altera??es feitas nas configura??es durante a sess?o ser?o perdidas. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configura??es do suplemento por documento para suplementos de conte?do e de painel de tarefas


Para persistir as configura??es de estado ou personalizadas de um suplemento de conte?do ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings) e seus m?todos. O conjunto de propriedades criado com os m?todos do objeto **Settings** est? dispon?vel apenas para a inst?ncia do suplemento de conte?do ou de painel de tarefas que o criou e apenas por meio do documento no qual ? salvo.

O objeto **Settings** ? carregado automaticamente como parte do objeto [Document](https://dev.office.com/reference/add-ins/shared/document) e est? dispon?vel quando o suplemento de conte?do ou de painel de tarefas ? ativado. Depois que o objeto **Document** ? instanciado, voc? pode acessar o objeto **Settings** com a propriedade [settings](https://dev.office.com/reference/add-ins/shared/document.settings) do objeto **Document**. Durante o tempo de vida da sess?o, voc? pode simplesmente usar os m?todos **Settings.get**, **Settings.set** e **Settings.remove** para ler, gravar ou remover configura??es persistentes e o estado do suplementos da c?pia na mem?ria do conjunto de propriedades.

Como os m?todos set e remove operam apenas em rela??o ? c?pia na mem?ria do conjunto de propriedades de configura??es, para salvar configura??es novas ou alteradas no documento ao qual o suplemento est? associado, voc? deve chamar o m?todo [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync).


### <a name="creating-or-updating-a-setting-value"></a>Criar ou atualizar um valor de configura??o

O exemplo de c?digo a seguir mostra como usar o m?todo [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) para criar uma configura??o chamada `'themeColor'` com um valor `'green'`. O primeiro par?metro do m?todo set ? _name_ (Id) da configura??o a ser definida ou criada, que diferencia mai?sculas de min?sculas. O segundo par?metro ? o _value_ da configura??o.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 A configura??o com o nome especificado ? criada se ainda n?o existir, ou seu valor ? atualizado se j? existir. Use o m?todo **Settings.saveAsync** para persistir as configura??es novas ou atualizadas para o documento.


### <a name="getting-the-value-of-a-setting"></a>Obter o valor de uma configura??o

O exemplo a seguir mostra como usar o m?todo [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) para obter o valor de uma configura??o chamada "themeColor". O ?nico par?metro do m?todo **get** ? o _name_ da configura??o (que diferencia mai?sculas de min?sculas).


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 O m?todo **get** retorna o valor que foi salvo anteriormente para a configura??o _name_ que foi passada. Se a configura??o n?o existir, o m?todo retornar? **null**.


### <a name="removing-a-setting"></a>Remover uma configura??o

O exemplo a seguir mostra como usar o m?todo [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) para remover uma configura??o com o nome "themeColor". O ?nico par?metro do m?todo **remove** ? o _name_ da configura??o (que diferencia mai?sculas de min?sculas).


```js
Office.context.document.settings.remove('themeColor');
```

Nada acontecer? se a configura??o n?o existir. Use o m?todo **Settings.saveAsync** para persistir a remo??o da configura??o do documento.


### <a name="saving-your-settings"></a>Salvar suas configura??es

Para salvar adi??es, altera??es ou exclus?es que o suplemento fez na c?pia na mem?ria do conjunto de propriedades de configura??es durante a sess?o atual, voc? deve chamar o m?todo [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) para armazen?-lo no documento. O ?nico par?metro do m?todo **saveAsync** ? _callback_, que ? uma fun??o de retorno de chamada com um ?nico par?metro. 


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

A fun??o an?nima passada ao m?todo **saveAsync** como o par?metro _callback_ ? executada quando a opera??o ? conclu?da. O par?metro _asyncResult_ do retorno de chamada fornece acesso a um objeto **AsyncResult** que cont?m o status da opera??o. No exemplo, a fun??o verifica a propriedade **AsyncResult.status** para ver se a opera??o de salvamento teve ?xito ou falhou e exibe o resultado na p?gina do suplemento.

## <a name="how-to-save-custom-xml-to-the-document"></a>Como salvar XML personalizado no documento

> [!NOTE]
> Esta se??o discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. A API JavaScript do Excel com host espec?fico tamb?m fornece acesso a partes XML personalizado. As APIs do Excel e os padr?es de programa??o s?o um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart).

H? uma op??o de armazenamento adicional caso precise armazenar informa??es que excedem os limites de tamanho do documento Settings ou que tenham um caractere estruturado. Voc? pode manter a marca??o XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observa??o na parte superior desta se??o). No Word, use o objeto [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) e seus m?todos (novamente, consulte a observa??o acima para o Excel). O c?digo a seguir cria um componente XML personalizado e exibe sua ID e seu conte?do em divs na p?gina. Dever? haver um atributo `xmlns` na cadeia de caracteres de XML.

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

Para recuperar uma parte do XML personalizado, use o m?todo [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync), mas a ID ? um GUID gerado quando parte de XML ? criada, portanto, n?o ? poss?vel saber ao codificar qual ? a ID. Por esse motivo, ao criar uma parte de XML, ? uma pr?tica recomendada armazenar imediatamente a ID da parte de XML como uma configura??o e usar uma chave f?cil de lembrar. O m?todo a seguir mostra como fazer isso. (Mas confira as se??es anteriores deste artigo para obter detalhes e as pr?ticas recomendadas ao trabalhar com configura??es personalizadas).

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

O c?digo a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configura??o.

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Como salvar configura??es na caixa de correio do usu?rio para suplementos do Outlook como configura??es m?veis


Um suplemento do Outlook pode usar o objeto [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para salvar o estado do suplemento e os dados de configura??es espec?ficos da caixa de correio do usu?rio. Esses dados s?o acess?veis apenas por esse suplemento do Outlook em nome do usu?rio que est? executando o suplemento. Os dados s?o armazenados na caixa de correio do Exchange Server do usu?rio e podem ser acessados ??quando o usu?rio faz logon em sua conta e executa o suplemento do Outlook.


### <a name="loading-roaming-settings"></a>Carregar configura??es m?veis


Um suplemento do Outlook normalmente carrega configura??es m?veis no manipulador de eventos [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize). O exemplo de c?digo JavaScript a seguir mostra como carregar configura??es m?veis existentes.


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


### <a name="creating-or-assigning-a-roaming-setting"></a>Criar ou atribuir uma configura??o m?vel


Continuando com o exemplo anterior, a fun??o `setAppSetting` a seguir mostra como usar o m?todo [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para definir ou atualizar uma configura??o chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configura??es m?veis de volta no Exchange Server com o m?todo [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings).


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

O m?todo **saveAsync** salva as configura??es m?veis de forma ass?ncrona e utiliza uma fun??o de retorno de chamada opcional. Este exemplo de c?digo passa uma fun??o de retorno de chamada denominada `saveMyAppSettingsCallback` para o m?todo **saveAsync**. Quando a chamada ass?ncrona ? retornada, o par?metro _asyncResult_ da fun??o `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) que voc? pode usar para determinar o ?xito ou a falha da opera??o com a propriedade **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Remover uma configura??o m?vel


Tamb?m estendendo os exemplos anteriores, a fun??o `removeAppSetting` a seguir mostra como usar o m?todo [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para remover a configura??o `cookie` e salvar todas as configura??es m?veis de volta no Exchange Server.


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Como salvar configura??es por item para suplementos do Outlook como propriedades personalizadas


As propriedades personalizadas permitem que o suplemento do Outlook armazene informa??es sobre um item com o qual est? trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugest?o de reuni?o em uma mensagem, voc? pode usar propriedades personalizadas para armazenar o fato de que a reuni?o foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook n?o se ofere?a para criar novamente o compromisso.

Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicita??o de reuni?o espec?fico, voc? deve carregar as propriedades na mem?ria chamando o m?todo [loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) do objeto **Item**. Se propriedades personalizadas j? estiverem definidas para o item atual, elas ser?o carregadas do servidor Exchange nesse momento. Ap?s carregar as propriedades, voc? pode usar os m?todos [set](https://dev.office.com/reference/add-ins/outlook/CustomProperties) e [get](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na mem?ria. Para salvar as altera??es feitas nas propriedades personalizadas do item, voc? deve usar o m?todo [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) para persistir as altera??es no item no servidor Exchange.


### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas

O exemplo a seguir mostra um conjunto simplificado de fun??es para um suplemento do Outlook que usa propriedades personalizadas. Voc? pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas. 

Um suplemento do Outlook que usa essas fun??es recupera as propriedades personalizadas chamando o m?todo **get** na vari?vel `_customProps`, conforme mostrado no exemplo a seguir.




```js
var property = _customProps.get("propertyName");
```

Este exemplo inclui as seguintes fun??es:



|**Nome da fun??o**|**Descri??o**|
|:-----|:-----|
| `Office.initialize`|Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do servidor Exchange.|
| `customPropsCallback`|Obt?m as propriedades personalizadas que s?o retornadas do servidor Exchange e as salva para uso posterior.|
| `updateProperty`|Define ou atualiza uma propriedade espec?fica e salva a altera??o no servidor Exchange.|
| `removeProperty`|Remove uma propriedade espec?fica e persiste a remo??o no servidor Exchange.|
| `saveCallback`|Retorno de chamada para chamadas ao m?todo **saveAsync** nas fun??es `updateProperty` e `removeProperty`.|



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


## <a name="see-also"></a>Confira tamb?m

- [No??es b?sicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
- [Suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
