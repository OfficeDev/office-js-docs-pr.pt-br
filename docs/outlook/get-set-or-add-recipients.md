---
title: Obter ou modificar destinatários em um suplemento do Outlook
description: Saiba como obter, definir ou adicionar destinatários de uma mensagem ou compromisso em um suplemento do Outlook.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: bcc4a76ef89e3bfaf7e884ad2fa4e1595782c62f
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958317"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Obter, configurar ou adicionar destinatários ao compor um compromisso ou uma mensagem no Outlook

A API JavaScript do Office fornece métodos assíncronos ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) ou [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) para obter, definir ou adicionar destinatários respectivamente em uma forma de composição de um compromisso ou mensagem. Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto do suplemento adequadamente para o Outlook ativar o suplemento em formulários de composição, conforme descrito em Criar [suplementos do Outlook](compose-scenario.md) para formulários de redação.

Algumas das propriedades que representam destinatários em um compromisso ou uma mensagem estão disponíveis para acesso de leitura em formulários de redação e de leitura. Essas propriedades incluem [optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) para compromissos, e [cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) para mensagens.

No formulário de leitura, você pode acessar a propriedade diretamente do objeto pai, como:

```js
item.cc
```

Mas em um formulário de composição, como o usuário e o suplemento podem inserir ou alterar um destinatário ao mesmo tempo, você deve usar o método assíncrono `getAsync` para obter essas propriedades, como no exemplo a seguir.

```js
item.cc.getAsync
```

Essas propriedades estão disponíveis para acesso de gravação somente nos formulários de composição, e não nos de leitura.

Assim como na maioria dos métodos assíncronos na API JavaScript para Office, `getAsync`e `setAsync``addAsync` adote parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira [Passar parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-recipients"></a>Obter os destinatários

Esta seção mostra um exemplo de código que obtém os destinatários do compromisso ou da mensagem que está sendo composta e exibe os endereços de email dos destinatários. O exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Na API JavaScript do Office, como as propriedades que representam os destinatários de um compromisso ( **optionalAttendees** e **requiredAttendees**) são diferentes daquelas de uma mensagem ([cco](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), **cc** e **para**), primeiro você deve usar a propriedade [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) para identificar se o item que está sendo composto é um compromisso ou uma mensagem. No modo de composição, todas essas propriedades de compromissos e mensagens são objetos [Recipients](/javascript/api/outlook/office.recipients) , portanto, você pode aplicar o método assíncrono para `Recipients.getAsync`obter os destinatários correspondentes.

Para usar `getAsync`, forneça uma função de retorno de chamada para verificar o status, os resultados e qualquer erro retornado pela chamada assíncrona `getAsync` . Você pode fornecer argumentos para a função de retorno de chamada usando o _parâmetro asyncContext_ opcional. A função de retorno de chamada retorna um _parâmetro de saída asyncResult_ . Você pode `status` `error` usar as propriedades do objeto de parâmetro [AsyncResult](/javascript/api/office/office.asyncresult) para verificar o status e as mensagens de erro da chamada assíncrona `value` e a propriedade para obter os destinatários reais. Os destinatários são representados como uma matriz de objetos [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

`getAsync` Observe que, como o método é assíncrono, se houver ações subsequentes que dependam de obter os destinatários com êxito, você deverá organizar seu código para iniciar essas ações somente na função de retorno de chamada correspondente quando a chamada assíncrona for concluída com êxito.

> [!IMPORTANT]
> O `getAsync` método retorna apenas destinatários resolvidos pelo cliente do Outlook. Um destinatário resolvido tem as seguintes características.
>
> - Se o destinatário tiver uma entrada salva no catálogo de endereços do remetente, o Outlook resolverá o endereço de email para o nome de exibição salvo do destinatário.
> - Um ícone de status de reunião do Teams é exibido antes do nome ou endereço de email do destinatário.
> - Um ponto e vírgula aparece após o nome ou endereço de email do destinatário.
> - O nome ou endereço de email do destinatário está sublinhado ou incluído em uma caixa.
>
> Para resolver um endereço de email depois que ele é adicionado a um item de email, o remetente deve usar a tecla **Tab** ou selecionar um contato ou endereço de email sugerido na lista de preenchimento automático.

> [!NOTE]
> No Outlook na Web e no Windows, se um usuário criar uma nova mensagem ativando o link de endereço de email de um contato do cartão de contato ou perfil, `Recipients.getAsync` a chamada do suplemento retornará o endereço de email `displayName` `EmailAddressDetails` do contato na propriedade do objeto associado em vez do nome salvo do contato.
>
> Para obter mais detalhes, consulte o [problema relacionado do GitHub](https://github.com/OfficeDev/office-js/issues/2201).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;
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
    for (let i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-recipients"></a>Definir os destinatários

Esta seção mostra um exemplo de código que define os destinatários do compromisso ou da mensagem que o usuário está compondo. A definição de destinatários substitui os destinatários existentes. Semelhante ao exemplo anterior que obtém os destinatários em um formulário de composição, este exemplo pressupõe que o suplemento é ativado nos formulários de composição para compromissos e mensagens. Este exemplo primeiro verifica se o item composto é um compromisso ou uma mensagem, portanto, para aplicar o método assíncrono, `Recipients.setAsync`nas propriedades apropriadas que representam os destinatários do compromisso ou da mensagem.

Ao chamar `setAsync`, forneça uma matriz como argumento de entrada para o parâmetro  _de destinatários_ , em um dos formatos a seguir.

- Uma matriz de cadeias de caracteres que são endereços SMTP.
- Uma matriz de dicionários, cada um contendo um nome para exibição e um endereço de email, conforme mostrado no exemplo de código a seguir.
- Uma matriz de `EmailAddressDetails` objetos, semelhante à retornada pelo `getAsync` método.
  
Opcionalmente, `setAsync` você pode fornecer uma função de retorno de chamada como um argumento de entrada para o método, para garantir que qualquer código que dependa da configuração bem-sucedida dos destinatários seja executado somente quando isso acontecer. Você também pode fornecer argumentos para a função de retorno de chamada usando o _parâmetro asyncContext_ opcional. Se você usar uma função de retorno de chamada, poderá acessar um parâmetro de saída _assíncrono_ e usar as propriedades  de **status** `AsyncResult` e erro do objeto de parâmetro para verificar o status e as mensagens de erro da chamada assíncrona.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;

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

Se você não quiser substituir nenhum destinatário existente em um compromisso ou mensagem, `Recipients.setAsync`em vez de usar, `Recipients.addAsync` poderá usar o método assíncrono para acrescentar destinatários. `addAsync` funciona da mesma forma que `setAsync` requer um argumento _de entrada_ de destinatários. Opcionalmente, você pode fornecer uma função de retorno de chamada e quaisquer argumentos para o retorno de chamada usando o parâmetro asyncContext. Em seguida, você pode verificar o status, o resultado e qualquer erro da chamada assíncrona `addAsync` usando o _parâmetro de saída asyncResult_ da função de retorno de chamada. O exemplo a seguir verifica se o item que está sendo composto é um compromisso, e anexa dois destinatários obrigatórios a ele.

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
