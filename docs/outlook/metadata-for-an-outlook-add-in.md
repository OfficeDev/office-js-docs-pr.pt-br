---
title: Obter e definir metadados em um suplemento do Outlook
description: Gerencie dados personalizados no suplemento do Outlook usando configurações de roaming ou propriedades personalizadas.
ms.date: 10/31/2019
ms.localizationpriority: medium
ms.openlocfilehash: fcff058fe05229d13a378fcba9c1b165e84fdd51
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148586"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>Obter e definir metadados de suplemento para um suplemento do Outlook

Você pode gerenciar dados personalizados em seu suplemento do Outlook usando um destes procedimentos:

- As configurações de roaming, que gerenciam dados personalizados para uma caixa de correio de usuário.
- Propriedades personalizadas, que gerenciam dados personalizados para um item na caixa de correio do usuário.

Ambos dão acesso a dados personalizados que só podem ser acessados por seu suplemento do Outlook, mas cada método armazena os dados separadamente. Ou seja, os dados armazenados por meio de configurações de roaming não podem ser acessados por propriedades personalizadas e vice-versa. Os dados são armazenados no servidor dessa caixa de correio e podem ser acessados nas sessões subsequentes do Outlook em todos os fatores forma a que o suplemento dê suporte.

## <a name="custom-data-per-mailbox-roaming-settings"></a>Dados personalizados por caixa de correio: configurações de roaming

Você pode especificar dados específicos para uma caixa de correio do Exchange de um usuário usando o objeto [RoamingSettings](/javascript/api/outlook/office.RoamingSettings). Exemplos desses dados incluem os dados pessoais e as preferências do usuário. O suplemento de email pode acessar as configurações de roaming quando faz roaming em qualquer dispositivo no qual deva ser executado (área de trabalho, tablet ou smartphone).

As mudanças nesses dados são armazenadas em uma cópia na memória dessas configurações para a sessão atual do Outlook. Você deve salvar explicitamente todas as configurações de roaming após a atualização para que elas fiquem disponíveis na próxima vez em que o usuário abrir o suplemento, no mesmo ou em qualquer outro dispositivo com suporte.


### <a name="roaming-settings-format"></a>Formato das configurações de roaming

Os dados de um objeto **RoamingSettings** são armazenados como uma cadeia de caracteres serializada JavaScript Object Notation (JSON). 

Abaixo temos um exemplo da estrutura, supondo que existam três configurações de roaming definidas chamadas `add-in_setting_name_0`, `add-in_setting_name_1` e `add-in_setting_name_2`.


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>Carregar configurações de roaming

Um suplemento de email normalmente carrega configurações de roaming no manipulador de eventos [Office.initialize](/javascript/api/office#Office_initialize_reason_). O exemplo de código JavaScript a seguir mostra como carregar configurações de roaming existentes e obter os valores de 2 configurações, **customerName** e **customerBalance**.


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


### <a name="creating-or-assigning-a-roaming-setting"></a>Criar ou atribuir uma configuração de roaming

Continuando com o exemplo anterior, a função JavaScript a seguir, `setAddInSetting`, mostra como usar o método [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) para definir uma configuração denominada `cookie` com a data de hoje e manter os dados usando o método [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveAsync_callback_) para salvar todas as configurações de roaming de volta no servidor.

O método cria a configuração se a configuração ainda não existir e atribui a `set` configuração ao valor especificado. O `saveAsync` método salva as configurações de roaming de forma assíncrona. Este exemplo de código passa um método de retorno de chamada, , para Quando a chamada assíncrona terminar, é chamado usando um `saveMyAddInSettingsCallback` `saveAsync`  `saveMyAddInSettingsCallback` parâmetro, _asyncResult_. Esse parâmetro é um objeto [AsyncResult](/javascript/api/office/office.asyncresult) que contém o resultado e detalhes sobre a chamada assíncrona. Você pode usar o parâmetro opcional _userContext_ para passar as informações de estado de chamada assíncrona à função de retorno de chamada.

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


### <a name="removing-a-roaming-setting"></a>Remover uma configuração móvel

Estendendo também os exemplos anteriores, a função JavaScript a seguir, `removeAddInSetting`, mostra como usar o método [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove_name_) para remover a definição `cookie` e salvar todas as configurações de roaming de volta no Exchange Server.


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


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a>Dados personalizados por item em uma caixa de correio: propriedades personalizadas

Você pode especificar dados específicos de um item na caixa de correio do usuário usando o objeto [CustomProperties](/javascript/api/outlook/office.CustomProperties). Por exemplo, seu suplemento de e-mail poderia categorizar determinadas mensagens e anotar a categoria usando uma propriedade personalizada `messageCategory`. Ou, se seu suplemento de e-mail cria compromissos de sugestões de reunião em uma mensagem, você pode usar uma propriedade personalizada para controlar cada um desses compromissos. Isso garante que se o usuário abrir a mensagem novamente, o suplemento de e-mail não se oferecerá para criar o compromisso uma segunda vez.

Semelhante às configurações de roaming, as mudanças nas propriedades personalizadas são armazenadas em cópias na memória das propriedades para a sessão atual do Outlook. Para garantir que essas propriedades personalizadas estarão disponíveis na próxima sessão, use[CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#saveAsync_callback__asyncContext_).

Essas propriedades personalizadas específicas do item e específicas do complemento só podem ser acessadas usando o `CustomProperties` objeto. Essas propriedades são diferentes das [UserProperties](/office/vba/api/Outlook.UserProperties) personalizadas baseadas em MAPI no modelo de objeto Outlook e propriedades estendidas no Exchange Web Services (EWS). Você não pode acessar `CustomProperties` diretamente usando o modelo de objeto Outlook, EWS ou REST. Para saber como acessar usando EWS ou REST, consulte a seção Obter propriedades `CustomProperties` [personalizadas usando EWS ou REST](#get-custom-properties-using-ews-or-rest).

### <a name="using-custom-properties"></a>Usar propriedades personalizadas

Antes de poder usar propriedades personalizadas, você precisa carregá-las chamando o método [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods). Após ter criado o conjunto de propriedades, você poderá usar os métodos [set](/javascript/api/outlook/office.customproperties#set_name__value_) e [get](/javascript/api/outlook/office.customproperties) para adicionar e recuperar propriedades personalizadas. Você deve usar o [saveAsync](/javascript/api/outlook/office.customproperties#saveAsync_callback__asyncContext_) método para salvar as alterações feitas no conjunto de propriedades.


 > [!NOTE]
 > Como o Outlook no Mac não armazena propriedades personalizadas em cache, se a rede do usuário é desativada, os suplementos de e-mail no Outlook no Mac não conseguem acessar suas propriedades personalizadas.


### <a name="custom-properties-example"></a>Exemplo de propriedades personalizadas


O exemplo a seguir mostra um conjunto de métodos simplificado para um suplemento do Outlook que usa propriedades personalizadas. Você pode usar este exemplo como ponto de partida para o seu suplemento que usa propriedades personalizadas.

Este exemplo inclui os métodos a seguir.


- [Office.initialize](/javascript/api/office#Office_initialize_reason_): inicializa o suplemento e carrega o conjunto de propriedades personalizadas do Exchange Server.

- **customPropsCallback**: obtém o recipiente de propriedades personalizadas que é retornado do servidor e o salva para uso posterior.

- **updateProperty**: define ou atualiza uma propriedade específica e salva a alteração no servidor.

- **removeProperty**: remove uma propriedade específica do recipiente de propriedades e salva a remoção no servidor.


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

### <a name="get-custom-properties-using-ews-or-rest"></a>Obtenha propriedades personalizadas usando EWS ou REST

Para obter **CustomProperties** usando EWS ou restante, você deverá primeiro determinar o nome do seu MAPI baseado propriedade estendida. Você pode obter propriedade da mesma forma que você teria qualquer propriedade com base MAPI estendida.

#### <a name="how-custom-properties-are-stored-on-an-item"></a>Como as propriedades personalizadas são armazenadas em um item

Propriedades personalizadas definidas por um suplemento não são equivalentes normal MAPI com base em Propriedades. APIs de complemento serializam todos os seus complementos como uma carga JSON e salvam-os em uma única propriedade estendida baseada em MAPI cujo nome é ( é a ID do seu complemento) e o GUID do conjunto de propriedades `CustomProperties` `cecp-<app-guid>` é `<app-guid>` `{00020329-0000-0000-C000-000000000046}` . (Para saber mais sobre esse objeto, confira [MS-OXCEXT 2.2.5 propriedades personalizadas do aplicativo de e-mail](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481).) Em seguida, você pode usar EWS ou REST para obter essa propriedade com base MAPI.

#### <a name="get-custom-properties-using-ews"></a>Obtenha propriedades personalizadas usando EWS

Seu complemento de email pode obter a propriedade estendida baseada em MAPI usando a `CustomProperties` operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) do EWS. Acesse no lado do servidor usando um token de retorno de chamada ou no lado do cliente usando o método `GetItem` [mailbox.makeEwsRequestAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) Na `GetItem` solicitação, especifique a propriedade baseada em MAPI em seu conjunto de propriedades usando os detalhes fornecidos na seção anterior Como as propriedades personalizadas são armazenadas `CustomProperties` [em um item](#how-custom-properties-are-stored-on-an-item).

O exemplo a seguir mostra como acessar um item e suas propriedades personalizadas.

> [!IMPORTANT]
> No exemplo a seguir, substituir `<app-guid>` com ID do suplemento.

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

Você também pode obter mais propriedades personalizadas se especificar na cadeia de caracteres solicitação como outros [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elementos.

#### <a name="get-custom-properties-using-rest"></a>Obtenha propriedades personalizadas usando REST

No seu suplemento, você pode criar sua consulta REST para mensagens e eventos para obter as que já têm propriedades personalizadas. Em sua consulta, você deve incluir o **CustomProperties** propriedades de MAPI baseados na sua propriedade definida utilizando os detalhes fornecidos na seção [como as propriedades personalizadas são armazenadas em um item](#how-custom-properties-are-stored-on-an-item).

O exemplo a seguir mostra como obter todos os eventos com as propriedades personalizadas definidos pelo seu suplemento e certifique-se que a resposta inclui o valor da propriedade para que você possa aplicar mais filtragem lógica.

> [!IMPORTANT]
> No exemplo a seguir substituir `<app-guid>` com ID do suplemento.

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

Outros exemplos que usam o REST para obter um único valor com base MAPI estendida, confira [Obter singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).

O exemplo a seguir mostra como acessar um item e suas propriedades personalizadas. Na função retorno de chamada para o `done` método `item.SingleValueExtendedProperties` contém uma lista das propriedades personalizadas solicitadas.

> [!IMPORTANT]
> No exemplo a seguir, substituir `<app-guid>` com ID do suplemento.

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

## <a name="see-also"></a>Confira também

- [Visão geral da propriedade MAPI](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Visão geral das propriedades do Outlook](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [Chamar APIs REST do Outlook de um suplemento do Outlook](use-rest-api.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)
- [Propriedades e propriedades estendidas no EWS no Exchange](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [Conjuntos de propriedades e formas de resposta no EWS no Exchange](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)