---
title: Obter ou definir a hora do compromisso em um suplemento do Outlook
description: Saiba como obter ou definir a hora de início e término de um compromisso em um suplemento do Outlook.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: eadce9b540a9b3b8a03186340fff4511d42dd35a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937850"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Obter ou definir a hora ao compor um compromisso no Outlook

A API JavaScript Office fornece métodos assíncronos ([Time.getAsync](/javascript/api/outlook/office.time#getAsync_options__callback_) e [Time.setAsync](/javascript/api/outlook/office.time#setAsync_dateTime__options__callback_)) para obter e definir a hora de início ou término de um compromisso que o usuário está compondo. Esses métodos assíncronos estão disponíveis apenas para compor os complementos. Para usar esses métodos, certifique-se de configurar o manifesto do complemento adequadamente para Outlook ativar os formulários de redação do complemento, conforme descrito em [Create Outlook add-ins for compose forms](compose-scenario.md).

As propriedades [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) estão disponíveis para compromissos tanto em formulários de composição quanto de leitura. No formulário de leitura, você pode acessar as propriedades diretamente do objeto pai, como em:

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

Como na maioria dos métodos assíncronos na API javaScript Office, **getAsync** e **setAsync** levam parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira [Passar parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-start-or-end-time"></a>Obter a hora de início ou de término

Esta seção mostra um exemplo de código que obtém a hora de início do compromisso que o usuário está compondo e a exibe. Você pode usar o mesmo código e substituir a propriedade **start** pela propriedade **end** para obter a hora de término. Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação de um compromisso, conforme mostrado abaixo.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Para usar **item.stsart.getAsync** ou **item.end.getAsync**, forneça um método de retorno de chamada que verifique o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional _asyncContext_. É possível obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, pode-se obter a hora de início como um objeto **Date** no formato UTC usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

Esta seção mostra um exemplo de código que define a hora de início do compromisso ou da mensagem que o usuário está redigindo. Você pode usar o mesmo código e substituir a propriedade **start** pela propriedade **end** para definir a hora de término. Observe que se o formulário de redação de compromisso já tiver uma hora de início, definir a hora de início ajustará a hora de término para manter a duração anterior do compromisso. Se o formulário de redação de compromisso já tiver uma hora de término, definir a hora de término ajustará a hora de término e a duração. Se o compromisso tiver sido definido como um evento de dia inteiro, definir a hora de início ajustará a hora de término para 24 horas depois e desmarcará a interface do usuário do evento de dia inteiro no formulário de redação.

Semelhante ao exemplo anterior, o código a seguir considera uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso.

Para usar **item.start.setAsync** ou **item.end.setAsync**, especifique um valor **Date** em UTC no parâmetro _dateTime_. Se você obtiver uma data com base em uma entrada do usuário no cliente, poderá usar [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para converter o valor para um objeto **Date** em UTC. É possível fornecer um método de retorno opcional e os argumentos para o método de retorno de chamada no parâmetro _asyncContext_. Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de hora de início ou de término especificada como texto sem formatação, substituindo a hora de início ou de término existente para o item.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
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
    
