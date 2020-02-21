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
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Obter, configurar ou adicionar destinatários ao compor um compromisso ou uma mensagem no Outlook


A API JavaScript para Office fornece métodos assíncronos ([Recipients. getasync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. Setasync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)ou [Recipients. addasync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) para obter, definir ou Adicionar destinatários em um formulário de composição de um compromisso ou uma mensagem. Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto do suplemento apropriadamente para que o Outlook ative o suplemento nos formulários de redação, conforme descrito em [criar suplementos do Outlook para formulários de composição](compose-scenario.md).

Algumas das propriedades que representam destinatários em um compromisso ou uma mensagem estão disponíveis para acesso de leitura em formulários de redação e de leitura. Essas propriedades incluem [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para compromissos, e [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para mensagens.

No formulário de leitura, você pode acessar a propriedade diretamente do objeto pai, como:

```js
item.cc
```

Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar um destinatário ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter essas propriedades, como no exemplo a seguir:


```js
item.cc.getAsync
```

Essas propriedades estão disponíveis para acesso de gravação somente nos formulários de redação, e não nos de leitura.

Como ocorre com a maioria dos métodos assíncronos da API JavaScript para Office, **getAsync**, **setAsync** e **addAsync** usam parâmetros de entrada opcionais. Saiba mais sobre como especificar esses parâmetros de entrada opcionais em [Transmitir parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-recipients"></a>Obter os destinatários


Esta seção mostra um exemplo de código que obtém os destinatários do compromisso ou da mensagem que está sendo composta e exibe os endereços de email dos destinatários. O exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Na API JavaScript para Office, como as propriedades que representam os destinatários de um compromisso (**optionalAttendees** e **requiredAttendees**) são diferentes das de uma mensagem ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc** e **to**), use primeiro a propriedade [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para identificar se o item que está sendo redigido é um compromisso ou uma mensagem. No modo de redação, todas essas propriedades de compromissos e mensagens são objetos [Recipients](/javascript/api/outlook/office.Recipients) e, portanto, você pode aplicar o método assíncrono **Recipients.getAsync** para obter os destinatários correspondentes.

Para usar **getAsync**, forneça um método de retorno de chamada para verificar o status, os resultados e eventuais erros retornados pela chamada assíncrona **getAsync**. Você pode fornecer argumentos para o método de retorno de chamada usando o parâmetro opcional _asyncContext_. O método de retorno de chamada retorna um parâmetro de saída _asyncResult_. É possível usar as propriedades **status** e **error** do objeto de parâmetro [AsyncResult](/javascript/api/office/office.asyncresult) para verificar o status e as mensagens de erro da chamada assíncrona e a propriedade **value** para obter os destinatários reais. Os destinatários são representados como uma matriz de objetos [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

Observe que, como o método **getAsync** é assíncrono, se houver ações subsequentes que dependam da obtenção bem-sucedida dos destinatários, você deverá organizar seu código para iniciar essas ações somente no método de retorno de chamada correspondente quando a chamada assíncrona for concluída com êxito.




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


## <a name="set-recipients"></a>Definir os destinatários


Esta seção mostra um exemplo de código que define os destinatários do compromisso ou da mensagem que o usuário está redigindo. A definição de destinatários substitui os destinatários existentes. De maneira semelhante ao exemplo anterior que obtém os destinatários em um formulário de redação, este exemplo pressupõe que o suplemento é ativado nos formulários de redação para compromissos e mensagens. Este exemplo primeiro verifica se o item composto é um compromisso ou uma mensagem para aplicar o método assíncrono, **Recipients.setAsync**, nas propriedades apropriadas que representam os destinatários do compromisso ou da mensagem.

Ao chamar **setAsync**, forneça uma matriz como argumento de entrada para o parâmetro _recipients_ em um dos seguintes formatos:


- Uma matriz de cadeias de caracteres que são endereços SMTP.
    
- Uma matriz de dicionários, cada um contendo um nome para exibição e um endereço de email, conforme mostrado no exemplo de código a seguir.
    
- Uma matriz de objetos **EmailAddressDetails**, semelhante à retornada pelo método **getAsync**.
    
Outra opção é fornecer um método de retorno de chamada como um argumento de entrada para o método **setAsync** para garantir que os códigos que dependam da definição bem-sucedida dos destinatários só sejam executados quando isso ocorrer. Você também pode fornecer argumentos para o método de retorno de chamada usando o parâmetro opcional _asyncContext_. Se você usar um método de retorno de chamada, será possível acessar um parâmetro de saída _asyncResult_ e usar as propriedades **status** e **error** do objeto de parâmetro **AsyncResult** para verificar o status e as mensagens de erro da chamada assíncrona.




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


## <a name="add-recipients"></a>Adicionar destinatários


Se não quiser substituir os destinatários existentes em um compromisso ou uma mensagem, em vez de usar **Recipients.setAsync**, use o método assíncrono **Recipients.addAsync** para acrescentar os destinatários. **addAsync** funciona da mesma forma que **setAsync** no sentido de exigir um argumento de entrada _recipients_. Outra opção é fornecer um método de retorno de chamada e os argumentos para o retorno de chamada usando o parâmetro asyncContext. Em seguida, é possível verificar o status, o resultado e eventuais erros da chamada assíncrona **addAsync** usando o parâmetro de saída _asyncResult_ do método de retorno de chamada. O exemplo a seguir verifica se o item que está sendo composto é um compromisso, e acrescenta dois destinatários obrigatórios a ele.


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


## <a name="see-also"></a>Confira também

- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)    
- [Obter e definir dados de item do Outlook em formulários de leitura ou composição](item-data.md)   
- [Criar suplementos do Outlook para formulários de composição](compose-scenario.md)    
- [Programação assíncrona em Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md)    
- [Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook](get-or-set-the-subject.md)    
- [Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook](insert-data-in-the-body.md)    
- [Obter ou definir o local ao criar um compromisso no Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Obter ou definir a hora ao criar um compromisso no Outlook](get-or-set-the-time-of-an-appointment.md)
    
