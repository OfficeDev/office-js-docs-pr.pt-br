---
title: Obter ou definir o local de um compromisso em um suplemento.
description: Saiba como obter ou definir o local de um compromisso em um suplemento do Outlook.
ms.date: 10/31/2019
ms.localizationpriority: medium
ms.openlocfilehash: 02a6360d43b91cde773d767d9a9838c015d9ecd7
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151855"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Obter ou definir o local ao compor um compromisso no Outlook

A Office API JavaScript fornece propriedades e métodos para gerenciar o local de um compromisso que o usuário está compondo. Atualmente, há duas propriedades que fornecem o local de um compromisso:

- [item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API básica que permite obter e definir o local.
- [item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API aprimorada que permite obter e definir o local e inclui a especificação do tipo [de local](/javascript/api/outlook/office.mailboxenums.locationtype). O tipo é `LocationType.Custom` se você definir o local usando `item.location` .

A tabela a seguir lista as APIs de local e os modos (ou seja, Redação ou Leitura) onde eles estão disponíveis.

| API | Modos de compromisso aplicáveis |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#location) | Participante/Leitura |
| [item.location.getAsync](/javascript/api/outlook/office.location#getAsync_options__callback_) | Organizer/Compose |
| [item.location.setAsync](/javascript/api/outlook/office.location#setAsync_location__options__callback_) | Organizer/Compose |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_) | Organizer/Compose,<br>Participante/Leitura |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_) | Organizer/Compose |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_) | Organizer/Compose |

Para usar os métodos disponíveis apenas para compor os complementos, configure o manifesto do add-in para ativar o add-in no modo Organizador/Redação. Consulte [Criar Outlook para obter formulários de composição](compose-scenario.md) para obter mais detalhes.

## <a name="use-the-enhancedlocation-api"></a>Usar a `enhancedLocation` API

Você pode usar a `enhancedLocation` API para obter e definir o local de um compromisso. O campo de localização dá suporte a vários locais e, para cada local, você pode definir o nome de exibição, o tipo e o endereço de email da sala de conferência (se aplicável). Consulte [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) para tipos de local com suporte.

### <a name="add-location"></a>Adicionar local

O exemplo a seguir mostra como adicionar um local chamando [addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_) em [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation).

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>Obter localização

O exemplo a seguir mostra como obter o local chamando [getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_) em [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedLocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a>Remover local

O exemplo a seguir mostra como remover o local chamando [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_) em [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>Usar a `location` API

Você pode usar a `location` API para obter e definir o local de um compromisso.

### <a name="get-the-location"></a>Obter o local

Esta seção mostra um exemplo de código que obtém o local do compromisso que o usuário está compondo e o exibe.

Para usar `item.location.getAsync`, forneça um método de retorno de chamada que verifica o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional `asyncContext`. Você pode obter status, resultados e qualquer erro usando o parâmetro de `asyncResult` saída do retorno de chamada. Se a chamada assíncrona for bem-sucedida, você poderá obter o local como uma cadeia de caracteres usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a>Definir o local

Esta seção mostra um exemplo de código que define a localização do compromisso que o usuário está redigindo.

Para usar `item.location.setAsync`, especifique uma cadeia de até 255 caracteres no parâmetro de dados. Opcionalmente, você pode fornecer um método de retorno de chamada e os argumentos para o método de retorno de chamada no parâmetro `asyncContext`. Você deve verificar o status, o resultado e qualquer mensagem de erro no parâmetro `asyncResult` de saída do retorno de chamada. Se a chamada assíncrona for bem-sucedida, `setAsync` inserirá a cadeia de caracteres de local especificada como texto sem formatação, substituindo o local existente pelo item.

> [!NOTE]
> Você pode definir vários locais usando um ponto e vírgula como separador (por exemplo, "Sala de conferência A; Sala de conferência B').

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a>Confira também

- [Criar seu primeiro Outlook de usuário](../quickstarts/outlook-quickstart.md)
- [Programação assíncrona nos Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md)
