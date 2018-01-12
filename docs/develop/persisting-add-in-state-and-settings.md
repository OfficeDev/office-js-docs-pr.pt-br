
# <a name="persisting-add-in-state-and-settings"></a>Persistir o estado e as configurações do suplemento

Essencialmente, os suplementos do Office são aplicativos Web em execução no ambiente sem estado de um controle de navegador. Como resultado, talvez o suplemento precise persistir dados para manter a continuidade de determinados recursos ou operações entre sessões de uso do suplemento. Por exemplo, o suplemento pode ter configurações personalizadas ou outros valores que precisa salvar e recarregar na próxima vez em que for inicializado, como o modo de exibição preferido ou o local padrão de um usuário.

Para fazer isso, você pode:


- Usar membros da API JavaScript para Office que armazenam dados como pares de nome/valor em um conjunto de propriedades armazenado em um local que depende do tipo de suplemento.
    
- Usar técnicas fornecidas pelo controle de navegador subjacente: cookies de navegador ou armazenamento Web HTML5 ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) ou [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)).
    
Este artigo concentra-se em como usar a API JavaScript para Office para persistir o estado do suplemento. Para obter exemplos do uso de cookies de navegador e armazenamento na Web, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>Persistir o estado e as configurações do suplemento com a API JavaScript para Office


A API JavaScript para Office fornece os objetos [Settings](../../reference/shared/settings.md), [RoamingSettings](../../reference/outlook/RoamingSettings.md) e [CustomProperties](../../reference/outlook/CustomProperties.md) para salvar o estado do suplemento entre sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) do suplemento que os criou.



|**Object**|**Suporte a tipos de suplementos**|**Local de armazenamento**|**Suporte ao host do Office**|
|:-----|:-----|:-----|:-----|
|[Configurações](../../reference/shared/settings.md)|conteúdo e painel de tarefas|O documento, a planilha ou a apresentação com o qual o suplemento está trabalhando. Configurações de suplementos de conteúdo e de painel de tarefas estão disponíveis para o suplemento que os criou por meio do documento em que são salvos. **Importante:** Não armazene senhas e outras IIP (informações de identificação pessoal) confidenciais com o objeto **Settings**. Os dados salvos não ficam visíveis para os usuários finais, mas são armazenados como parte do documento, que pode ser acessado pela leitura direta do formato de arquivo do documento. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento somente no servidor que hospeda o suplemento como um recurso protegido pelo usuário.|Word, Excel ou PowerPoint **Observação:** Os suplementos de painel de tarefas para o Project 2013 não dão suporte à API **Settings** para o armazenamento do estado ou das configurações do suplemento. No entanto, para suplementos em execução no Project (bem como outros aplicativos de host do Office), você pode usar técnicas como cookies de navegador ou armazenamento na Web. Para saber mais sobre essas técnicas, confira [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|A caixa de correio do servidor Exchange do usuário em que o suplemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem se "mover" com o usuário e estão disponíveis para o suplemento quando ele é executado no contexto de qualquer aplicativo de host de cliente com suporte ou navegador que acessa a caixa de correio do usuário. As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.|Outlook|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Os dados de configurações são gerenciados na memória no tempo de execução


Internamente, os dados no conjunto de propriedades acessado com os objetos **Settings**, **CustomProperties** ou **RoamingSettings** são armazenados como um objeto JSON (JavaScript Object Notation) serializado que contém pares de nome/valor. O nome (chave) de cada valor deve ser uma **cadeia de caracteres**, e o valor armazenado pode ser uma **cadeia de caracteres**, um **número**, uma **data** ou um **objeto** JavaScript, mas não uma **função**.

Este exemplo da estrutura do conjunto de propriedades contém três valores de **cadeia de caracteres** definidos nomeados como `firstName`, `location` e `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

Depois que o conjunto de propriedades de configurações é salvo durante a sessão anterior do suplemento, ele pode ser carregado quando o suplemento é inicializado ou a qualquer momento depois disso durante a sessão atual do suplemento. Durante a sessão, as configurações são gerenciadas inteiramente na memória usando os métodos **get**, **set** e **remove** do objeto que corresponde às configurações de tipo que você está criando (**Settings**, **CustomProperties** ou **RoamingSettings**). 


 >**Importante**  Para persistir as adições, atualizações ou exclusões feitas durante a sessão atual do suplemento para o local de armazenamento, você deve chamar o método **saveAsync** do objeto correspondente usado para trabalhar com esse tipo de configurações. Os métodos **get**, **set** e **remove** operam somente na cópia na memória do conjunto de propriedades de configurações. Se o suplemento for fechado sem chamar **saveAsync**, as alterações feitas nas configurações durante a sessão serão perdidas. 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas


Para persistir as configurações de estado ou personalizadas de um suplemento de conteúdo ou de painel de tarefas para Word, Excel ou PowerPoint, use o objeto [Settings](../../reference/shared/settings.md) e seus métodos. O conjunto de propriedades criado com os métodos do objeto **Settings** está disponível apenas para a instância do suplemento de conteúdo ou de painel de tarefas que o criou e apenas por meio do documento no qual é salvo.

O objeto **Settings** é carregado automaticamente como parte do objeto [Document](../../reference/shared/document.md) e está disponível quando o suplemento de conteúdo ou de painel de tarefas é ativado. Depois que o objeto **Document** é instanciado, você pode acessar o objeto **Settings** com a propriedade [settings](../../reference/shared/document.settings.md) do objeto **Document**. Durante o tempo de vida da sessão, você pode simplesmente usar os métodos **Settings.get**, **Settings.set** e **Settings.remove** para ler, gravar ou remover configurações persistentes e o estado do suplementos da cópia na memória do conjunto de propriedades.

Como os métodos set e remove operam apenas em relação à cópia na memória do conjunto de propriedades de configurações, para salvar configurações novas ou alteradas no documento ao qual o suplemento está associado, você deve chamar o método [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


### <a name="creating-or-updating-a-setting-value"></a>Criar ou atualizar um valor de configuração

O exemplo de código a seguir mostra como usar o método [Settings.set](../../reference/shared/settings.set.md) para criar uma configuração chamada `'themeColor'` com um valor `'green'`. O primeiro parâmetro do método set é _name_ (Id) da configuração a ser definida ou criada, que diferencia maiúsculas de minúsculas. O segundo parâmetro é o _value_ da configuração.


```
Office.context.document.settings.set('themeColor', 'green');
```

 A configuração com o nome especificado é criada se ainda não existir, ou seu valor é atualizado se já existir. Use o método **Settings.saveAsync** para persistir as configurações novas ou atualizadas para o documento.


### <a name="getting-the-value-of-a-setting"></a>Obter o valor de uma configuração

O exemplo a seguir mostra como usar o método [Settings.get](../../reference/shared/settings.get.md) para obter o valor de uma configuração chamada "themeColor". O único parâmetro do método **get** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 O método **get** retorna o valor que foi salvo anteriormente para a configuração _name_ que foi passada. Se a configuração não existir, o método retornará **null**.


### <a name="removing-a-setting"></a>Remover uma configuração

O exemplo a seguir mostra como usar o método [Settings.remove](../../reference/shared/settings.removehandlerasync.md) para remover uma configuração com o nome "themeColor". O único parâmetro do método **remove** é o _name_ da configuração (que diferencia maiúsculas de minúsculas).


```
Office.context.document.settings.remove('themeColor');
```

Nada acontecerá se a configuração não existir. Use o método **Settings.saveAsync** para persistir a remoção da configuração do documento.


### <a name="saving-your-settings"></a>Salvar suas configurações

Para salvar adições, alterações ou exclusões que o suplemento fez na cópia na memória do conjunto de propriedades de configurações durante a sessão atual, você deve chamar o método [Settings.saveAsync](../../reference/shared/settings.saveasync.md) para armazená-lo no documento. O único parâmetro do método **saveAsync** é _callback_, que é uma função de retorno de chamada com um único parâmetro. 


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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis


Um suplemento do Outlook pode usar o objeto [RoamingSettings](../../reference/outlook/RoamingSettings.md) para salvar o estado e os dados de configurações do suplemento específico da caixa de correio do usuário. Esses dados são acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento. Os dados são armazenados na caixa de correio do usuário do Exchange Server e ficam acessíveis quando esse usuário faz logon em sua conta e executa o suplemento do Outlook.


### <a name="loading-roaming-settings"></a>Carregar configurações móveis


Um suplemento do Outlook normalmente carrega configurações móveis no manipulador de eventos [Office.initialize](../../reference/shared/office.initialize.md). O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.


```
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


Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md).


```
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

O método **saveAsync** salva as configurações móveis de forma assíncrona e utiliza uma função de retorno de chamada opcional. Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**. Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](../../reference/outlook/simple-types.md) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.


### <a name="removing-a-roaming-setting"></a>Remover uma configuração móvel


Também estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) para remover a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Como salvar configurações por item para suplementos do Outlook como propriedades personalizadas


As propriedades personalizadas permitem que o suplemento do Outlook armazene informações sobre um item com o qual está trabalhando. Por exemplo, se o suplemento do Outlook criar um compromisso com base em uma sugestão de reunião em uma mensagem, você pode usar propriedades personalizadas para armazenar o fato de que a reunião foi criada. Isso garante que, se a mensagem for aberta novamente, o suplemento do Outlook não se ofereça para criar novamente o compromisso.

Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicitação de reunião específico, você deve carregar as propriedades na memória chamando o método [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](../../reference/outlook/CustomProperties.md) e [get](../../reference/outlook/RoamingSettings.md) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na memória. Para salvar as alterações feitas nas propriedades personalizadas do item, você deve usar o método [saveAsync](../../reference/outlook/CustomProperties.md) para persistir as alterações no item no servidor Exchange.


### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas

O exemplo a seguir mostra um conjunto simplificado de funções para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas. 

Um suplemento do Outlook que usa essas funções recupera as propriedades personalizadas chamando o método **get** na variável `_customProps`, conforme mostrado no exemplo a seguir.




```
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



```
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


## <a name="additional-resources"></a>Recursos adicionais



- [Noções básicas da API JavaScript para Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Suplementos do Outlook](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
