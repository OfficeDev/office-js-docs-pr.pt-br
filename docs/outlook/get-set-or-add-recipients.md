---
title: Obter ou modificar destinatários em um suplemento do Outlook
description: Saiba como obter, definir ou adicionar destinatários de uma mensagem ou compromisso em um suplemento do Outlook.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 36849b0ebb7e1dff34d59305d265294452bf395d
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165820"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="89c8e-103">Obter, configurar ou adicionar destinatários ao compor um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-103">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>


<span data-ttu-id="89c8e-104">A API JavaScript para Office fornece métodos assíncronos ([Recipients. getasync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. Setasync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)ou [Recipients. addasync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) para obter, definir ou Adicionar destinatários em um formulário de composição de um compromisso ou uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="89c8e-104">The JavaScript API for Office provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-), or [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) to respectively get, set, or add recipients in a compose form of an appointment or message.</span></span> <span data-ttu-id="89c8e-105">Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto do suplemento apropriadamente para que o Outlook ative o suplemento nos formulários de redação, conforme descrito em [criar suplementos do Outlook para formulários de composição](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="89c8e-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="89c8e-p102">Algumas das propriedades que representam destinatários em um compromisso ou uma mensagem estão disponíveis para acesso de leitura em formulários de redação e de leitura. Essas propriedades incluem [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para compromissos, e [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para mensagens.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p102">Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for appointments, and [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and  [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for messages.</span></span>

<span data-ttu-id="89c8e-108">No formulário de leitura, você pode acessar a propriedade diretamente do objeto pai, como:</span><span class="sxs-lookup"><span data-stu-id="89c8e-108">In a read form, you can access the property directly from the parent object, such as:</span></span>

```js
item.cc
```

<span data-ttu-id="89c8e-109">Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar um destinatário ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter essas propriedades, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="89c8e-109">But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method **getAsync** to get these properties, as in the following example:</span></span>


```js
item.cc.getAsync
```

<span data-ttu-id="89c8e-110">Essas propriedades estão disponíveis para acesso de gravação somente nos formulários de redação, e não nos de leitura.</span><span class="sxs-lookup"><span data-stu-id="89c8e-110">These properties are available for write access in only compose forms and not read forms.</span></span>

<span data-ttu-id="89c8e-p103">Como ocorre com a maioria dos métodos assíncronos da API JavaScript para Office, **getAsync**, **setAsync** e **addAsync** usam parâmetros de entrada opcionais. Saiba mais sobre como especificar esses parâmetros de entrada opcionais em [Transmitir parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="89c8e-p103">As with most asynchronous methods in the JavaScript API for Office, **getAsync**, **setAsync**, and **addAsync** take optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-recipients"></a><span data-ttu-id="89c8e-113">Obter os destinatários</span><span class="sxs-lookup"><span data-stu-id="89c8e-113">Get recipients</span></span>


<span data-ttu-id="89c8e-p104">Esta seção mostra um exemplo de código que obtém os destinatários do compromisso ou da mensagem que está sendo composta e exibe os endereços de email dos destinatários. O exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p104">This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

<span data-ttu-id="89c8e-p105">Na API JavaScript para Office, como as propriedades que representam os destinatários de um compromisso (**optionalAttendees** e **requiredAttendees**) são diferentes das de uma mensagem ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc** e **to**), use primeiro a propriedade [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para identificar se o item que está sendo redigido é um compromisso ou uma mensagem. No modo de redação, todas essas propriedades de compromissos e mensagens são objetos [Recipients](/javascript/api/outlook/office.Recipients) e, portanto, você pode aplicar o método assíncrono **Recipients.getAsync** para obter os destinatários correspondentes.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p105">In the JavaScript API for Office, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc**, and **to**), you should first use the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to identify whether the item being composed is an appointment or message. In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.Recipients) objects, so you can then apply the asynchronous method, **Recipients.getAsync**, to get the corresponding recipients.</span></span>

<span data-ttu-id="89c8e-p106">Para usar **getAsync**, forneça um método de retorno de chamada para verificar o status, os resultados e eventuais erros retornados pela chamada assíncrona **getAsync**. Você pode fornecer argumentos para o método de retorno de chamada usando o parâmetro opcional _asyncContext_. O método de retorno de chamada retorna um parâmetro de saída _asyncResult_. É possível usar as propriedades **status** e **error** do objeto de parâmetro [AsyncResult](/javascript/api/office/office.asyncresult) para verificar o status e as mensagens de erro da chamada assíncrona e a propriedade **value** para obter os destinatários reais. Os destinatários são representados como uma matriz de objetos [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).</span><span class="sxs-lookup"><span data-stu-id="89c8e-p106">To use **getAsync**, provide a callback method to check for the status, results, and any error returned by the asynchronous **getAsync** call. You can provide any arguments to the callback method using the optional _asyncContext_ parameter. The callback method returns an _asyncResult_ output parameter. You can use the **status** and **error** properties of the [AsyncResult](/javascript/api/office/office.asyncresult) parameter object to check for status and any error messages of the asynchronous call, and the **value** property to get the actual recipients. Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects.</span></span>

<span data-ttu-id="89c8e-123">Observe que, como o método **getAsync** é assíncrono, se houver ações subsequentes que dependam da obtenção bem-sucedida dos destinatários, você deverá organizar seu código para iniciar essas ações somente no método de retorno de chamada correspondente quando a chamada assíncrona for concluída com êxito.</span><span class="sxs-lookup"><span data-stu-id="89c8e-123">Note that because the **getAsync** method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-recipients"></a><span data-ttu-id="89c8e-124">Definir os destinatários</span><span class="sxs-lookup"><span data-stu-id="89c8e-124">Set recipients</span></span>


<span data-ttu-id="89c8e-p107">Esta seção mostra um exemplo de código que define os destinatários do compromisso ou da mensagem que o usuário está redigindo. A definição de destinatários substitui os destinatários existentes. De maneira semelhante ao exemplo anterior que obtém os destinatários em um formulário de redação, este exemplo pressupõe que o suplemento é ativado nos formulários de redação para compromissos e mensagens. Este exemplo primeiro verifica se o item composto é um compromisso ou uma mensagem para aplicar o método assíncrono, **Recipients.setAsync**, nas propriedades apropriadas que representam os destinatários do compromisso ou da mensagem.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p107">This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user. Setting recipients overwrites any existing recipients. Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages. This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method, **Recipients.setAsync**, on the appropriate properties that represent recipients of the appointment or message.</span></span>

<span data-ttu-id="89c8e-129">Ao chamar **setAsync**, forneça uma matriz como argumento de entrada para o parâmetro _recipients_ em um dos seguintes formatos:</span><span class="sxs-lookup"><span data-stu-id="89c8e-129">When calling **setAsync**, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:</span></span>


- <span data-ttu-id="89c8e-130">Uma matriz de cadeias de caracteres que são endereços SMTP.</span><span class="sxs-lookup"><span data-stu-id="89c8e-130">An array of strings that are SMTP addresses.</span></span>
    
- <span data-ttu-id="89c8e-131">Uma matriz de dicionários, cada um contendo um nome para exibição e um endereço de email, conforme mostrado no exemplo de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="89c8e-131">An array of dictionaries, each containing a display name and email address, as shown in the following code sample.</span></span>
    
- <span data-ttu-id="89c8e-132">Uma matriz de objetos **EmailAddressDetails**, semelhante à retornada pelo método **getAsync**.</span><span class="sxs-lookup"><span data-stu-id="89c8e-132">An array of **EmailAddressDetails** objects, similar to the one returned by the **getAsync** method.</span></span>
    
<span data-ttu-id="89c8e-p108">Outra opção é fornecer um método de retorno de chamada como um argumento de entrada para o método **setAsync** para garantir que os códigos que dependam da definição bem-sucedida dos destinatários só sejam executados quando isso ocorrer. Você também pode fornecer argumentos para o método de retorno de chamada usando o parâmetro opcional _asyncContext_. Se você usar um método de retorno de chamada, será possível acessar um parâmetro de saída _asyncResult_ e usar as propriedades **status** e **error** do objeto de parâmetro **AsyncResult** para verificar o status e as mensagens de erro da chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p108">You can optionally provide a callback method as an input argument to the **setAsync** method, to make sure any code that depends on successfully setting the recipients would execute only when that happens. You can also provide any arguments for the callback method using the optional _asyncContext_ parameter. If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the **AsyncResult** parameter object to check for status and any error messages of the asynchronous call.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="add-recipients"></a><span data-ttu-id="89c8e-136">Adicionar destinatários</span><span class="sxs-lookup"><span data-stu-id="89c8e-136">Add recipients</span></span>


<span data-ttu-id="89c8e-p109">Se não quiser substituir os destinatários existentes em um compromisso ou uma mensagem, em vez de usar **Recipients.setAsync**, use o método assíncrono **Recipients.addAsync** para acrescentar os destinatários. **addAsync** funciona da mesma forma que **setAsync** no sentido de exigir um argumento de entrada _recipients_. Outra opção é fornecer um método de retorno de chamada e os argumentos para o retorno de chamada usando o parâmetro asyncContext. Em seguida, é possível verificar o status, o resultado e eventuais erros da chamada assíncrona **addAsync** usando o parâmetro de saída _asyncResult_ do método de retorno de chamada. O exemplo a seguir verifica se o item que está sendo composto é um compromisso, e acrescenta dois destinatários obrigatórios a ele.</span><span class="sxs-lookup"><span data-stu-id="89c8e-p109">If you do not want to overwrite any existing recipients in an appointment or message, instead of using **Recipients.setAsync**, you can use the **Recipients.addAsync** asynchronous method to append recipients. **addAsync** works similarly as **setAsync** in that it requires a _recipients_ input argument. You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter. You can then check the status, result, and any error of the asynchronous **addAsync** call by using the _asyncResult_ output parameter of the callback method. The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.</span></span>


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="see-also"></a><span data-ttu-id="89c8e-142">Confira também</span><span class="sxs-lookup"><span data-stu-id="89c8e-142">See also</span></span>

- [<span data-ttu-id="89c8e-143">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-143">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="89c8e-144">Obter e definir dados de item do Outlook em formulários de leitura ou composição</span><span class="sxs-lookup"><span data-stu-id="89c8e-144">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="89c8e-145">Criar suplementos do Outlook para formulários de composição</span><span class="sxs-lookup"><span data-stu-id="89c8e-145">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="89c8e-146">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="89c8e-146">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="89c8e-147">Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-147">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)    
- [<span data-ttu-id="89c8e-148">Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-148">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)    
- [<span data-ttu-id="89c8e-149">Obter ou definir o local ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-149">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="89c8e-150">Obter ou definir a hora ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="89c8e-150">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
