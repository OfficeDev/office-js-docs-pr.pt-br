---
title: Gerenciar estado e configurações para um Outlook de um Outlook de dados
description: Saiba como persistir o estado e as configurações do Outlook de um Outlook.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 896c473baad95515b199d8934c81745c619374a0
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484677"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Gerenciar estado e configurações para um Outlook de um Outlook de dados

> [!NOTE]
> Revise o estado e as configurações [persistentes do add-in](../develop/persisting-add-in-state-and-settings.md) na seção **Conceitos principais** desta documentação antes de ler este artigo.

Para Outlook de Outlook, a API javaScript Office fornece objetos [RoamingSettings](/javascript/api/outlook/office.roamingsettings) e [CustomProperties](/javascript/api/outlook/office.customproperties) para salvar o estado do add-in em sessões, conforme descrito na tabela a seguir. Em todos os casos, os valores de configurações salvos são associados à [Id](/javascript/api/manifest/id) do suplemento que os criou.

|**Objeto**|**Local de armazenamento**|
|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|A caixa de correio Exchange servidor do usuário onde o complemento está instalado. Como essas configurações são armazenadas na caixa de correio do servidor do usuário, elas podem "circular" com o usuário e estão disponíveis para o add-in quando ele está em execução no contexto de qualquer cliente com suporte acessando a caixa de correio desse usuário.<br/><br/> As configurações móveis de suplementos do Outlook estão disponíveis apenas para o suplemento que os criou e somente por meio da caixa de correio em que o suplemento está instalado.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|A mensagem, o compromisso ou o item de solicitação de reunião com o qual o suplemento está trabalhando. As propriedades personalizadas de itens de suplementos do Outlook estão disponíveis apenas para o suplemento que as criou e apenas por meio do item em que estão salvas.|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Como salvar configurações na caixa de correio do usuário para suplementos do Outlook como configurações móveis

Um suplemento do Outlook pode usar o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) para salvar o estado e os dados de configurações do suplemento específico da caixa de correio do usuário. Esses dados são acessíveis somente para esse suplemento do Outlook em nome do usuário que executa o suplemento. Os dados são armazenados na caixa de correio do usuário do Exchange Server e ficam acessíveis quando esse usuário faz logon em sua conta e executa o suplemento do Outlook.

### <a name="loading-roaming-settings"></a>Carregar configurações de roaming

O exemplo de código JavaScript a seguir mostra como carregar configurações móveis existentes.

```js
var _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a>Criar ou atribuir uma configuração móvel

Continuando com o exemplo anterior, a função `setAppSetting` a seguir mostra como usar o método [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-set-member(1)) para definir ou atualizar uma configuração chamada `cookie` com a data de hoje. Em seguida, ele salva todas as configurações móveis de volta no Exchange Server com o método [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)).

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

O método **saveAsync** salva as configurações móveis de forma assíncrona e utiliza uma função de retorno de chamada opcional. Este exemplo de código passa uma função de retorno de chamada denominada `saveMyAppSettingsCallback` para o método **saveAsync**. Quando a chamada assíncrona é retornada, o parâmetro _asyncResult_ da função `saveMyAppSettingsCallback` fornece acesso a um objeto [AsyncResult](/javascript/api/office/office.asyncresult) que você pode usar para determinar o êxito ou a falha da operação com a propriedade **AsyncResult.status**.

### <a name="removing-a-roaming-setting"></a>Remover uma configuração móvel

Também estendendo os exemplos anteriores, a função `removeAppSetting` a seguir mostra como usar o método [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) para remover a configuração `cookie` e salvar todas as configurações móveis de volta no Exchange Server.

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

Para poder usar propriedades personalizadas para uma mensagem, um compromisso ou um item de solicitação de reunião específico, você deve carregar as propriedades na memória chamando o método [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) do objeto **Item**. Se propriedades personalizadas já estiverem definidas para o item atual, elas serão carregadas do servidor Exchange nesse momento. Após carregar as propriedades, você pode usar os métodos [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) e [get](/javascript/api/outlook/office.roamingsettings) para o objeto **CustomProperties** para adicionar, atualizar e recuperar propriedades na memória. Para salvar as alterações feitas nas propriedades personalizadas do item, você deve usar o método [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) para persistir as alterações no item no servidor Exchange.

### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas

O exemplo a seguir mostra um conjunto simplificado de funções para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar esse exemplo como ponto de partida para o suplemento do Outlook que usa propriedades personalizadas.

Um suplemento do Outlook que usa essas funções recupera as propriedades personalizadas chamando o método **obter** na variável `_customProps`, conforme mostrado no exemplo a seguir.

```js
var property = _customProps.get("propertyName");
```

Este exemplo inclui as seguintes funções.

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

### <a name="platform-behavior-in-emails"></a>Comportamento da plataforma em emails

A tabela a seguir resume o comportamento de propriedades personalizadas salvas em emails para vários Outlook clientes.

|Cenário|Windows|Web|Mac|
|---|---|---|---|
|Nova composição|null|null|null|
|Responder, responder a todos|null|null|null|
|Encaminhar|Carrega as propriedades do pai|null|null|
|Item enviado de nova composição|null|null|null|
|Item enviado de resposta ou resposta a todos|null|null|null|
|Item enviado do forward|Remove as propriedades do pai se não for salva|null|null|

Para lidar com a situação em Windows:

1. Verifique se há propriedades existentes na inicialização do seu add-in e mantenha-os ou desmarcar-os conforme necessário.
1. Ao definir propriedades personalizadas, inclua uma propriedade adicional para indicar se as propriedades personalizadas foram adicionadas durante a leitura da mensagem ou pelo modo de leitura do complemento. Isso ajudará você a diferenciar se a propriedade foi criada durante a composição ou herdada do pai.
1. Para verificar se o usuário está encaminhando um email ou respondendo, você pode usar [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)) (disponível no conjunto de requisitos 1.10).

## <a name="see-also"></a>Confira também

- [Persistir o estado e as configurações do suplemento](../develop/persisting-add-in-state-and-settings.md)
- [Inicialize seu suplemento do Office](../develop/initialize-add-in.md)
