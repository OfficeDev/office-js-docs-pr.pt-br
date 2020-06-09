---
title: Obter ou definir a hora do compromisso em um suplemento do Outlook
description: Saiba como obter ou definir a hora de início e término de um compromisso em um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5e02523852584d4b5f1ede9bcd191b9ee16d4c24
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609130"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="d09c1-103">Obter ou definir a hora ao compor um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-103">Get or set the time when composing an appointment in Outlook</span></span>

<span data-ttu-id="d09c1-104">A API JavaScript do Office fornece métodos assíncronos ([time. getasync](/javascript/api/outlook/office.Time#getasync-options--callback-) e [time. setasync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) para obter e definir a hora de início ou de término de um compromisso que o usuário está redigindo.</span><span class="sxs-lookup"><span data-stu-id="d09c1-104">The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) and [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) to get and set the start or end time of an appointment that the user is composing.</span></span> <span data-ttu-id="d09c1-105">Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto do suplemento apropriadamente para que o Outlook ative o suplemento nos formulários de redação, conforme descrito em [criar suplementos do Outlook para formulários de composição](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="d09c1-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="d09c1-p102">As propriedades [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) estão disponíveis para compromissos tanto em formulários de composição quanto de leitura. No formulário de leitura, você pode acessar as propriedades diretamente do objeto pai, como em:</span><span class="sxs-lookup"><span data-stu-id="d09c1-p102">The [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:</span></span>

```js
item.start
```

<span data-ttu-id="d09c1-108">e em:</span><span class="sxs-lookup"><span data-stu-id="d09c1-108">and in:</span></span>

```js
item.end
```

<span data-ttu-id="d09c1-109">Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar a hora ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter a hora de início ou de término, conforme mostrado abaixo:</span><span class="sxs-lookup"><span data-stu-id="d09c1-109">But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:</span></span>

```js
item.start.getAsync
```

<span data-ttu-id="d09c1-110">e:</span><span class="sxs-lookup"><span data-stu-id="d09c1-110">and:</span></span>

```js
item.end.getAsync
```

<span data-ttu-id="d09c1-111">Assim como a maioria dos métodos assíncronos na API JavaScript do Office, **getasync** e **setasync** aceita parâmetros de entrada opcionais.</span><span class="sxs-lookup"><span data-stu-id="d09c1-111">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="d09c1-112">Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira [Passar parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="d09c1-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-start-or-end-time"></a><span data-ttu-id="d09c1-113">Obter a hora de início ou de término</span><span class="sxs-lookup"><span data-stu-id="d09c1-113">Get the start or end time</span></span>

<span data-ttu-id="d09c1-p104">Esta seção mostra um exemplo de código que obtém a hora de início do compromisso que o usuário está compondo e a exibe. Você pode usar o mesmo código e substituir a propriedade **start** pela propriedade **end** para obter a hora de término. Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação de um compromisso, conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="d09c1-p104">This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.</span></span>


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

<span data-ttu-id="d09c1-p105">Para usar **item.stsart.getAsync** ou **item.end.getAsync**, forneça um método de retorno de chamada que verifique o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional _asyncContext_. É possível obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, pode-se obter a hora de início como um objeto **Date** no formato UTC usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="d09c1-p105">To use **item.start.getAsync** or **item.end.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


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


## <a name="set-the-start-or-end-time"></a><span data-ttu-id="d09c1-121">Definir a hora de início ou de término</span><span class="sxs-lookup"><span data-stu-id="d09c1-121">Set the start or end time</span></span>

<span data-ttu-id="d09c1-p106">Esta seção mostra um exemplo de código que define a hora de início do compromisso ou da mensagem que o usuário está redigindo. Você pode usar o mesmo código e substituir a propriedade **start** pela propriedade **end** para definir a hora de término. Observe que se o formulário de redação de compromisso já tiver uma hora de início, definir a hora de início ajustará a hora de término para manter a duração anterior do compromisso. Se o formulário de redação de compromisso já tiver uma hora de término, definir a hora de término ajustará a hora de término e a duração. Se o compromisso tiver sido definido como um evento de dia inteiro, definir a hora de início ajustará a hora de término para 24 horas depois e desmarcará a interface do usuário do evento de dia inteiro no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="d09c1-p106">This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.</span></span>

<span data-ttu-id="d09c1-127">Semelhante ao exemplo anterior, o código a seguir considera uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso.</span><span class="sxs-lookup"><span data-stu-id="d09c1-127">Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.</span></span>

<span data-ttu-id="d09c1-p107">Para usar **item.start.setAsync** ou **item.end.setAsync**, especifique um valor **Date** em UTC no parâmetro _dateTime_. Se você obtiver uma data com base em uma entrada do usuário no cliente, poderá usar [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para converter o valor para um objeto **Date** em UTC. É possível fornecer um método de retorno opcional e os argumentos para o método de retorno de chamada no parâmetro _asyncContext_. Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de hora de início ou de término especificada como texto sem formatação, substituindo a hora de início ou de término existente para o item.</span><span class="sxs-lookup"><span data-stu-id="d09c1-p107">To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback method and any arguments for the callback method in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.</span></span>




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


## <a name="see-also"></a><span data-ttu-id="d09c1-133">Confira também</span><span class="sxs-lookup"><span data-stu-id="d09c1-133">See also</span></span>

- [<span data-ttu-id="d09c1-134">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-134">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="d09c1-135">Obter e definir dados de item do Outlook em formulários de leitura ou composição</span><span class="sxs-lookup"><span data-stu-id="d09c1-135">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="d09c1-136">Criar suplementos do Outlook para formulários de composição</span><span class="sxs-lookup"><span data-stu-id="d09c1-136">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="d09c1-137">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d09c1-137">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="d09c1-138">Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-138">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="d09c1-139">Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-139">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)   
- [<span data-ttu-id="d09c1-140">Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-140">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="d09c1-141">Obter ou definir o local ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="d09c1-141">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
    
