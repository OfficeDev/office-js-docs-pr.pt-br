---
title: Persistir o estado e as configurações do suplemento
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: ce2b9ffce97e6338d62cdf07d722ffa384283d28
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458066"
---
# <a name="persisting-add-in-state-and-settings"></a>Persistir o estado e as configurações do suplemento

Essencialmente, os suplementos do Office são aplicativos Web em execução no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário. Para fazer isso, você pode:

- Usar os membros da API JavaScript para Office que armazena dados como:
    -  Pares de nome/valor em um recipiente de propriedades armazenado em um local que depende do tipo de suplemento.
    -  XML personalizado armazenado no documento.
    
- Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    
Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistir o estado e as configurações do suplemento com a API JavaScript para Office

A API JavaScript para Office fornece os objetos [Settings](https://docs.microsoft.com/javascript/api/office/office.settings), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings) e [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties) para salvar o estado do suplemento entre sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id) do suplemento que os criou.

|**Object**|**Suporte a tipos de suplementos**|**Local de armazenamento**|**Suporte ao host do Office**|
|:-----|:-----|:-----|:-----|
|[Configurações](https://docs.microsoft.com/javascript/api/office/office.settings)|conteúdo e painel de tarefas|O documento, a planilha ou a apresentação com o qual o suplemento está trabalhando. Configurações de suplementos de conteúdo e de painel de tarefas estão disponíveis para o suplemento que os criou por meio do documento em que são salvos.<br/><br/>**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.|Word, Excel ou PowerPoint<br/><br/> **Observação:** os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para suplementos em execução no Project (bem como outros aplicativos de host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings)|Outlook|A caixa de correio do servidor Exchange do usuário em que o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem se "mover" com o usuário e estão disponíveis para o suplemento quando ele é executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessa a caixa de correio do usuário.<br/><br/> As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.|Outlook|
|[CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties)|Outlook|A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.|Outlook|
|[CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts)|painel de tarefas|O documento, planilha ou apresentação com o qual o suplemento está trabalhando. As configurações de suplementos do painel de tarefas estão disponíveis para o suplemento que as criou por meio do documento em que são salvos.<br/><br/>**Importante:** não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais em uma parte XML personalizada. objeto. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.|Word (usando a API comum de JavaScript do Office), Excel (usando a API do JavaScript do Excel com host específico)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configurações são gerenciados na memória no tempo de execução

> [!NOTE]
> As duas seções a seguir discutem configurações no contexto da API comum de JavaScript do Office. A API JavaScript do Excel com host específico também fornece acesso às configurações personalizadas. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [SettingCollection do Excel](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection).

Internamente, os dados no conjunto de propriedades acessados com os objetos **Configurações**, **CustomProperties** ou **RoamingSettings** são armazenados como um objeto JSON (JavaScript Object Notation) serializado que contém pares de nome/valor. O nome (chave) de cada valor deve ser uma **cadeia**, e o valor armazenado pode ser uma **cadeia**, **um número**, **uma data**, ou **objeto** JavaScript, mas não uma **função**.

Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento. Durante a sessão, as configurações são gerenciadas inteiramente na memória usando os métodos **obter**, **configurar** e **remover** do objeto que corresponde às configurações de tipo que você está criando (**Definições**, **CustomProperties** ou **RoamingSettings**). 


> [!IMPORTANT]
> Para persistir as adições, atualizações ou exclusões feitas durante a sessão atual do suplemento para o local de armazenamento, você deve chamar o método **saveAsync** do objeto correspondente usado para trabalhar com esse tipo de configurações. Os métodos **obter**, **definir**, e**remover** operam somente na cópia na memória do conjunto de propriedades de configurações. Se o suplemento for fechado sem chamar **saveAsync**, as alterações feitas nas configurações durante a sessão serão perdidas. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas


Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](https://docs.microsoft.com/javascript/api/office/office.settings) e seus métodos. O conjunto de propriedades criado com os métodos do objeto **Settings** está disponível apenas para a instância do suplemento de conteúdo ou de painel de tarefas que o criou e apenas por meio do documento no qual é salvo.

O objeto **Configurações** é carregado automaticamente como parte do objeto [Documento](https://docs.microsoft.com/javascript/api/office/office.document) e está disponível quando o suplemento de conteúdo ou de painel de tarefas é ativado. Depois que o objeto **Documento** é instanciado, você pode acessar o objeto **Configurações** com a propriedade [configurações](https://docs.microsoft.com/javascript/api/office/office.document#settings) do objeto **Documento**. Durante o tempo de vida da sessão, você pode simplesmente usar os métodos **Settings.get**, **Settings.set**, e **Settings.remove** para ler, gravar ou remover configurações persistentes e o estado do suplementos da cópia na memória do conjunto de propriedades.

Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings#saveasync-options--callback-).


### <a name="creating-or-updating-a-setting-value"></a>Criar ou atualizar um valor de configuração

O exemplo de código a seguir mostra como usar o método [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings#set-name--value-) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro é o _value_ da configuração.


```js
Office.context.document.settings.set('themeColor', 'green');
```

 A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir. Use o método **Settings.saveAsync** para persistir as configurações novas ou atualizadas para o documento.


### <a name="getting-the-value-of-a-setting"></a>Obter o valor de uma configuração

O exemplo a seguir mostra como usar o método [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings#get-name-) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do método **get** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 O método **get** retorna o valor que foi salvo anteriormente para a configuração _name_ que foi passada. Se a configuração não existir, o método retornará **null**.


### <a name="removing-a-setting"></a>Remover uma configuração

O exemplo a seguir mostra como usar o método [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings#remove-name-) para remover uma configuração com o nome "themeColor". O único parâmetro do método **remove** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).


```js
Office.context.document.settings.remove('themeColor');
```

Nada acontecerá se a configuração não existir. Use o método **Settings.saveAsync** para persistir a remoção da configuração do documento.


### <a name="saving-your-settings"></a>Salvar suas configurações

Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings#saveasync-options--callback-) para armazená-lo no documento. O único parâmetro do método **saveAsync** é _callback_, que é uma função de retorno de chamada com um único parâmetro. 


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

A função anônima passada ao método **saveAsync** como o parâmetro _callback_ é executada quando a operação é concluída. O parâmetro _asyncResult_ do retorno de chamada fornece acesso a um objeto **AsyncResult** que contém o status da operação. No exemplo, a função verifica a propriedade **AsyncResult.status** para ver se a operação de salvamento teve êxito ou falhou e exibe o resultado na página do suplemento.

## <a name="how-to-save-custom-xml-to-the-document"></a>Como salvar XML personalizado no documento

> [!NOTE]
> Esta seção discute as partes XML no contexto da API comum do JavaScript do Office com suporte no Word. A API JavaScript do Excel com host específico também fornece acesso a partes XML personalizado. As APIs do Excel e os padrões de programação são um pouco diferentes. Para saber mais, confira [Excel CustomXmlPart](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart).

Há uma opção de armazenamento adicional caso precise armazenar informações que excedem os limites de tamanho do documento Settings ou que tenham um caractere estruturado. Você pode manter a marcação XML personalizada em um suplemento do painel tarefas do Word (e do Excel, mas confira a observação na parte superior desta seção). No Word, use o objeto [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart) e seus métodos (novamente, consulte a observação acima do Excel). O código a seguir cria um componente XML personalizado e exibe sua ID e seu conteúdo no divs na página. Observe que deverá haver um atributo `xmlns` na cadeia de caracteres de XML.

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

Para recuperar uma parte do XML personalizado, use o método [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-), mas a ID é um GUID gerado quando parte de XML é criada, portanto, não é possível saber ao codificar qual é a ID. Por esse motivo, ao criar uma parte de XML, é uma prática recomendada armazenar imediatamente a ID da parte de XML como uma configuração e usar uma chave fácil de lembrar. O método a seguir mostra como fazer isso. (Mas confira as seções anteriores deste artigo para obter detalhes e as práticas recomendadas ao trabalhar com configurações personalizadas).

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

O código a seguir mostra como recuperar parte do XML obtendo primeiro a sua ID em uma configuração.

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


Um suplemento do Outlook pode usar o objeto [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings) para salvar o estado e os dados de configurações do suplemento específico da caixa de correio do usuário. Esses dados são acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento. Os dados são armazenados na caixa de correio do usuário do Exchange Server e ficam acessíveis quando esse usuário faz logon em sua conta e executa o suplemento do Outlook.


### <a name="loading-roaming-settings"></a>Carregar configurações de roaming


Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](https://docs.microsoft.com/javascript/api/office). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.


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


### <a name="creating-or-assigning-a-roaming-setting"></a>Criar ou atribuir uma configuração móvel


Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#set-name--value-) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#saveasync-callback-).


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

O método **saveAsync** salva as configurações móveis de forma assíncrona e utiliza uma função de retorno de chamada opcional. Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**. Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](https://docs.microsoft.com/javascript/api/outlook) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Remover uma configuração móvel


Também estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings#remove-name-) para remover a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Como salvar configurações por item para suplementos do Outlook como propriedades personalizadas


As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.

Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicitação de reunião específico, você deve carregar as propriedades na memória chamando o método [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties#set-name--value-) e [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na memória. Para salvar as alterações feitas nas propriedades personalizadas do item, você deve usar o método [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) para persistir as alterações no item no servidor Exchange.


### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas

O exemplo a seguir mostra um conjunto simplificado de funções para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas. 

Um suplemento do Outlook que usa essas funções recupera as propriedades personalizadas chamando o método **obter** na variável `_customProps`, conforme mostrado no exemplo a seguir.




```js
var property = _customProps.get("propertyName");
```

Este exemplo inclui as seguintes funções:



|**Nome da função**|**Descrição**|
|:-----|:-----|
| `Office.initialize`|Inicializa o suplemento e carrega as propriedades personalizadas para o item atual a partir do servidor Exchange.|
| `customPropsCallback`|Obtém as propriedades personalizadas que são retornadas do servidor Exchange e as salva para uso posterior.|
| `updateProperty`|Define ou atualiza uma propriedade específica e salva a alteração no servidor Exchange.|
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
    
