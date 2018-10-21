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
# <a name="persisting-add-in-state-and-settings"></a>Persistência do estado e das configurações do suplemento

Essencialmente, os suplementos do Office são aplicativos Web executados no ambiente sem estado do controle de um navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:

- Usar os membros da API JavaScript para Office que armazena dados como:
    -  Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.
    -  XML personalizado armazenado no documento.
    
- Usar técnicas fornecidas pelo controle do navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    
Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistir o estado e as configurações do suplemento com a API JavaScript para Office

A API JavaScript para Office fornece os objetos [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) e [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) para salvar o estado do suplemento entre sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configuração salvos são associados ao [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id?view=office-js) do suplemento que os criou.

|**Objeto**|**Tipo de suplemento compatível**|**Local de armazenamento**|**Host do Office compatível**|
|:-----|:-----|:-----|:-----|
|[Configurações](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js)|conteúdo e painel de tarefas|O documento, a planilha ou a apresentação com o qual o suplemento está trabalhando. As configurações dos suplementos de conteúdo e de painel de tarefas são disponibilizadas para o suplemento que os criou por meio do documento onde são salvos.<br/><br/>**Importante:** não armazene senhas e outras informações de identificação pessoal (PII) confidenciais junto com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer PII necessárias ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido de usuário.|Word, Excel ou PowerPoint<br/><br/> **Observação:** os suplementos de painel de tarefas para o Project 2013 não são compatíveis com a API **Settings** para armazenamento de estado e configurações do suplemento. No entanto, para suplementos executados no Project (bem como outros aplicativos host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js)|Outlook|A caixa de correio do Exchange Server do usuário onde o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem acompanhar o usuário e são disponibilizadas para o suplemento quando ele for executado no contexto de qualquer aplicativo host de cliente compatível ou navegador que acessa a caixa de correio do usuário.<br/><br/> As configurações móveis de um suplemento do Outlook são disponibilizadas apenas para o suplemento que as criou e somente por meio da caixa de correio onde o suplemento está instalado.|Outlook|
|[CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js)|Outlook|A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas dos itens de um suplemento do Outlook são disponibilizadas apenas para o suplemento que as criou e apenas por meio do item onde foram salvas.|Outlook|
|[CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)|painel de tarefas|O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas são disponibilizadas para o suplemento que as criou por meio do documento onde foram salvas.<br/><br/>**Importante:** não armazene senhas e outras informações de identificação pessoal (PII) confidenciais em uma parte XML personalizada. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer PII necessárias ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido de usuário.|Word (usando a API JavaScript Comum do Office), Excel (usando a API JavaScript do Excel específica do host)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configuração são gerenciados na memória em tempo de execução

> [!NOTE]
> As duas seções seguintes abordam as configurações no contexto da API JavaScript Comum do Office. A API JavaScript do Excel também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para obter mais informações, confira [SettingCollection do Excel](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).

Internamente, os dados no recipiente de propriedades acessado com as objetos **Settings**, **CustomProperties** ou **RoamingSettings** são armazenados como objetos JSON (JacaScript Object Notation) serializados contendo pares nome/valor. O nome (chave) de cada valor deve ser uma **sequência de caracteres** e o valor armazenado pode ser uma **sequência de caracteres**, um **número**, uma **data** ou um **objeto** Javascript, mas não uma **função**.

Este exemplo de estrutura do recipiente de propriedades contém três valores definidos de **sequência de caracteres** chamados `firstName`, `location` e `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Após o recipiente de propriedades de configuração ter sido salvo na sessão anterior do suplemento, ele pode ser carregado na inicialização do suplemento ou em qualquer outro momento durante a sessão atual do suplemento. Durante a sessão, as configurações são totalmente gerenciadas em memória usando os métodos **get**, **set** e **remove** do objeto que corresponde ao tipo de configuração que você criou ( **Settings**, **CustomProperties** ou **RoamingSettings**). 


> [!IMPORTANT]
> Para persistir quaisquer adições, atualizações ou exclusões realizadas durante a sessão atual do suplemento no local de armazenamento, você deve chamar o método **saveAsync** do objeto correspondente ao que foi usado para trabalhar com esse tipo de configuração. Os métodos **get**, **set** e **remove** operam somente na cópia em memória do recipiente de propriedades de configuração. Se o seu suplemento for fechado sem chamar **saveAsync**, quaisquer alterações realizadas nas configurações durante a sessão serão perdidas. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas


Para persistir o estado ou as configurações personalizadas de um suplemento de conteúdo ou de painel de tarefas do Word, Excel ou PowerPoint, use o objeto [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) e seus métodos. O recipiente de propriedades criado com os métodos do objeto **Settings** está disponível apenas para a instância do suplemento de conteúdo ou de painel de tarefas que o criou e apenas por meio do documento onde é salvo.

O objeto **Settings** é carregado automaticamente como parte do objeto [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) e fica disponível quando suplemento de painel de tarefas ou conteúdo é ativado. Depois que o objeto **Document** é instanciado, você pode acessar o objeto **Settings** com a propriedade [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) do objeto **Document** . Durante o ciclo de vida da sessão, você pode usar apenas os métodos **Settings.get**, **Settings.set** e **Settings.remove** para ler, gravar ou remover as configurações persistentes e o estado do suplemento na cópia em memória do recipiente de propriedades.

Como os métodos set e remove operam apenas em relação à cópia em memória do recipiente de propriedades de configuração, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-).


### <a name="creating-or-updating-a-setting-value"></a>Criação ou atualização de um valor de configuração

O exemplo de código a seguir mostra como usar o método [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro, _value_, é o valor da configuração.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 A configuração é criada com o nome especificado se ainda não existir, ou seu valor é atualizado se já existir. Use o método **Settings.saveAsync** para persistir as configurações novas ou atualizadas no documento.


### <a name="getting-the-value-of-a-setting"></a>Obtenção do valor de uma configuração

O exemplo a seguir mostra como usar o método [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do método **get** é _name_ , o nome da configuração (que diferencia maiúsculas de minúsculas).


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 O método **get** retorna o valor salvo anteriormente da configuração _name_ que foi passada para o método. Se a configuração não existir, o método retornará **null**.


### <a name="removing-a-setting"></a>Exclusão de uma configuração

O exemplo a seguir mostra como usar o método [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) para excluir uma configuração com o nome "themeColor". O único parâmetro do método **remove** é _name_, o nome da configuração (que diferencia maiúsculas de minúsculas).


```js
Office.context.document.settings.remove('themeColor');
```

Nada acontecerá se a configuração não existir. Use o método **Settings.saveAsync** para persistir a exclusão da configuração no documento.


### <a name="saving-your-settings"></a>Gravação das configurações

Para salvar adições, alterações ou exclusões que o suplemento fez na cópia em memória do recipiente de propriedades de configuração durante a sessão atual, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) para armazená-lo no documento. O único parâmetro do método **saveAsync** é _callback_, que é uma função de retorno de chamada com um único parâmetro. 


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

A função anônima passada para o método **saveAsync** como parâmetro de _retorno de chamada_ será executada quando a operação for concluída. O parâmetro do retorno de chamada  _asyncResult_ dá acesso a um objeto **AsyncResult** que contém o status da operação. No exemplo, a função verifica a propriedade **AsyncResult.status** para verificar se a operação de gravação foi bem-sucedida ou falhou e, em seguida, exibe o resultado na página do suplemento.

## <a name="how-to-save-custom-xml-to-the-document"></a>Como salvar XML personalizado no documento

> [!NOTE]
> Esta seção aborda partes XML personalizadas no contexto da API JavaScript comum do Office que é suportada no Word. A API JavaScript do Excel também fornece acesso às partes XML personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para obter mais informações, consulte [CustomXmlPart do Excel](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).

Há uma opção de armazenamento adicional caso você precise armazenar informações que excedam os limites de tamanho das configurações do documento ou que possuam um caráter estruturado. Você pode persistir a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação no início desta seção). No Word, você pode usar o objeto [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) e seus métodos (novamente, confira a observação acima sobre o Excel). O código a seguir cria uma parte XML personalizada, exibe seu ID e, em seguida, seu conteúdo em divs na página. Observe que deve exisitir um atributo `xmlns` na sequência de caracteres XML.

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

Para recuperar uma parte XML personalizada, você deve usar o método [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) , mas o ID é um GUID gerado durante a criação da parte XML , então você não pode referenciá-lo no código. Por esse motivo, uma boa prática é armazenar o ID da parte XML como uma configuração imediatamente após a sua criação e atribuir uma chave fácil de lembrar . O método a seguir mostra como fazer isso. (Mas consulte as seções anteriores deste artigo para obter detalhes e práticas recomendadas ao trabalhar com configurações personalizadas).

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

O código a seguir mostra como recuperar a parte do XML obtendo  o seu ID a partir de uma configuração.

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis


Um suplemento do Outlook pode usar o objeto [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) para salvar o estado do suplemento e dados de configurações específicas na caixa de correio do usuário. Esses dados ficam acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento. Os dados são armazenados na caixa de correio do usuário do Exchange Server e pode ser acessados quando o usuário faz logon em sua conta e executa o suplemento do Outlook.


### <a name="loading-roaming-settings"></a>Carga de configurações móveis


Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.


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


### <a name="creating-or-assigning-a-roaming-setting"></a>Criação ou atribuição de uma configuração móvel


Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-).


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

O método **saveAsync** salva as configurações móveis de forma assíncrona e recebe uma função de retorno de chamada opcional. Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**. Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Exclusão de uma configuração móvel


Ainda estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) para excluir a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Como salvar configurações por item como propriedades personalizadas em suplementos do Outlook


As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook cria um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.

Para poder usar propriedades personalizadas em um item específico como uma mensagem, um compromisso ou uma solicitação de reunião, você deve carregar as propriedades em memória chamando o método [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) e [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) do objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades em memória. Para salvar quaisquer alterações realizadas nas propriedades personalizadas do item, você deve usar o método [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) para persistir as alterações no item no servidor Exchange.


### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas

O exemplo a seguir mostra um conjunto simplificado de funções de um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o seu suplemento do Outlook que usa propriedades personalizadas. 

Um suplemento do Outlook que usa essas funções recupera quaisquer propriedades personalizadas chamando o método **get** na variável `_customProps`, conforme mostrado no exemplo a seguir.




```js
var property = _customProps.get("propertyName");
```

Este exemplo inclui as seguintes funções:



|**Nome da função**|**Descrição**|
|:-----|:-----|
| `Office.initialize`|Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do Exchange Server.|
| `customPropsCallback`|Obtém as propriedades personalizadas retornadas do Exchange Server e as salva para uso posterior.|
| `updateProperty`|Define ou atualiza uma propriedade específica e salva a alteração no Exchange Server.|
| `removeProperty`|Remove uma propriedade específica e persiste a remoção no servidor Exchange.|
| `saveCallback`|Retorno de chamada para chamadas ao método **saveAsync** nas funções `updateProperty` e `removeProperty`.|



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


## <a name="see-also"></a>Confira também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
