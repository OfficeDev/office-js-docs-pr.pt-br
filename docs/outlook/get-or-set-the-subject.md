---
title: Obter ou definir o assunto em um suplemento do Outlook
description: Saiba como obter ou definir o assunto de uma mensagem ou compromisso em um suplemento do Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 93864aee005af61d9648c39402a843d9105bb021
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325437"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="bce02-103">Obter ou definir o assunto ao compor um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-103">Get or set the subject when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="bce02-104">A API JavaScript do Office fornece métodos assíncronos ([Subject. getasync](/javascript/api/outlook/office.Subject#getasync-options--callback-) e [Subject. setasync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) para obter e definir o assunto de um compromisso ou uma mensagem que o usuário está redigindo.</span><span class="sxs-lookup"><span data-stu-id="bce02-104">The Office JavaScript API provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) and [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing.</span></span> <span data-ttu-id="bce02-105">Esses métodos assíncronos estão disponíveis somente para os suplementos de composição. Para usar esses métodos, verifique se você configurou o manifesto do suplemento apropriadamente para que o Outlook ative o suplemento nos formulários de composição.</span><span class="sxs-lookup"><span data-stu-id="bce02-105">These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.</span></span>

<span data-ttu-id="bce02-p102">A propriedade **subject** está disponível para acesso de leitura nos formulários de leitura e de redação de compromissos e de mensagens. Em um formulário de leitura, é possível acessar a propriedade diretamente do objeto pai, como em:</span><span class="sxs-lookup"><span data-stu-id="bce02-p102">The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:</span></span>

```js
item.subject
```

<span data-ttu-id="bce02-108">Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar o assunto ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter o assunto, conforme mostrado abaixo:</span><span class="sxs-lookup"><span data-stu-id="bce02-108">But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:</span></span>

```js
item.subject.getAsync
```

<span data-ttu-id="bce02-109">A propriedade **subject** está disponível para acesso de gravação somente nos formulários de redação, e não nos de leitura.</span><span class="sxs-lookup"><span data-stu-id="bce02-109">The **subject** property is available for write access in only compose forms and not in read forms.</span></span>

<span data-ttu-id="bce02-110">Assim como a maioria dos métodos assíncronos na API JavaScript do Office, **getasync** e **setasync** aceita parâmetros de entrada opcionais.</span><span class="sxs-lookup"><span data-stu-id="bce02-110">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="bce02-111">Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira "Passar parâmetros opcionais para métodos assíncronos" em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bce02-111">For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-subject"></a><span data-ttu-id="bce02-112">Obter o assunto</span><span class="sxs-lookup"><span data-stu-id="bce02-112">Get the subject</span></span>

<span data-ttu-id="bce02-p104">Esta seção mostra um exemplo de código que obtém o assunto do compromisso ou da mensagem que o usuário está compondo e o exibe. Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="bce02-p104">This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

<span data-ttu-id="bce02-p105">Para usar **item.subject.getAsync**, forneça um método de retorno de chamada que verifique o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional _asyncContext_. É possível obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, você poderá obter o assunto como texto sem formatação usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="bce02-p105">To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a><span data-ttu-id="bce02-119">Definir o assunto</span><span class="sxs-lookup"><span data-stu-id="bce02-119">Set the subject</span></span>


<span data-ttu-id="bce02-p106">Esta seção mostra um exemplo de código que define o assunto do compromisso ou da mensagem que o usuário está compondo. Semelhante ao exemplo anterior, o código a seguir considera uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bce02-p106">This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.</span></span>

<span data-ttu-id="bce02-p107">Para usar **item.subject.setAsync**, especifique uma cadeia de até 255 caracteres no parâmetro de dados. Outra opção é fornecer um método de retorno de chamada e os argumentos para o método de retorno de chamada no parâmetro _asyncContext_. Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de assunto especificada como texto sem formatação, substituindo o assunto existente pelo item.</span><span class="sxs-lookup"><span data-stu-id="bce02-p107">To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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


## <a name="see-also"></a><span data-ttu-id="bce02-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="bce02-126">See also</span></span>

- [<span data-ttu-id="bce02-127">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-127">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)   
- [<span data-ttu-id="bce02-128">Obter e definir dados de item do Outlook em formulários de leitura ou composição</span><span class="sxs-lookup"><span data-stu-id="bce02-128">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="bce02-129">Criar suplementos do Outlook para formulários de composição</span><span class="sxs-lookup"><span data-stu-id="bce02-129">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="bce02-130">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bce02-130">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="bce02-131">Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-131">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="bce02-132">Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-132">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="bce02-133">Obter ou definir o local ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-133">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="bce02-134">Obter ou definir a hora ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="bce02-134">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
