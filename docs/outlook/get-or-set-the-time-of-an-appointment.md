---
title: Obter ou definir a hora do compromisso em um suplemento do Outlook
description: Saiba como obter ou definir a hora de início e término de um compromisso em um suplemento do Outlook.
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7aa40fda15c613aca869af8b277d4deb6fbf833
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541230"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Obter ou definir a hora ao compor um compromisso no Outlook

A API JavaScript do Office fornece métodos assíncronos ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) e [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) para obter e definir a hora de início ou término de um compromisso que o usuário está redigindo. Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto XML do suplemento adequadamente para o Outlook ativar o suplemento em formulários de composição, conforme descrito em Criar [suplementos do Outlook](compose-scenario.md) para formulários de redação. Não há suporte para regras de ativação em suplementos que usam um manifesto do [Teams para suplementos do Office (versão prévia).](../develop/json-manifest-overview.md)

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

```js
item.start
```

e em:

```js
item.end
```

Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar a hora ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter a hora de início ou de término, conforme mostrado abaixo:

```js
item.start.getAsync
```

e:

```js
item.end.getAsync
```

Assim como na maioria dos métodos assíncronos na API JavaScript do Office, **getAsync** e **setAsync** usam parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira [Passar parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-start-or-end-time"></a>Obter a hora de início ou de término

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

Para usar **item.start.getAsync** ou **item.end.getAsync**, forneça uma função de retorno de chamada que verifica o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para a função de retorno de chamada por meio do  _parâmetro opcional asyncContext_ . Você pode obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, você pode obter a hora de início como um objeto **Date** no formato UTC usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-the-start-or-end-time"></a>Definir a hora de início ou de término

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

Semelhante ao exemplo anterior, o código a seguir considera uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso.

Para usar **item.start.setAsync** ou **item.end.setAsync**, especifique um valor **date** em UTC no _parâmetro dateTime_ . Se você obtiver uma data com base em uma entrada do usuário no cliente, poderá usar [mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para converter o valor para um objeto **Date** em UTC. Você pode fornecer uma função de retorno de chamada opcional e quaisquer argumentos para a função de retorno de chamada no _parâmetro asyncContext_ . Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de hora de início ou de término especificada como texto sem formatação, substituindo a hora de início ou de término existente para o item.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
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

- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Obter e definir dados de item do Outlook em formulários de leitura ou composição](item-data.md)
- [Criar suplementos do Outlook para formulários de composição](compose-scenario.md)
- [Programação assíncrona em Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook](get-set-or-add-recipients.md)  
- [Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook](get-or-set-the-subject.md)
- [Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook](insert-data-in-the-body.md)
- [Obter ou definir o local ao criar um compromisso no Outlook](get-or-set-the-location-of-an-appointment.md)
