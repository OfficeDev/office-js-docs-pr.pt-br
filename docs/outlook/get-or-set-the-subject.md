---
title: Obter ou definir o assunto em um suplemento do Outlook
description: Saiba como obter ou definir o assunto de uma mensagem ou compromisso em um suplemento do Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: cfe2ad8010090e21606d2b9ec95ab2bed79d5ccf956a27e4cee303e09ca68349
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098320"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Obter ou definir o assunto ao compor um compromisso ou uma mensagem no Outlook

A API javaScript Office fornece métodos assíncronos ([subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_) e [subject.setAsync](/javascript/api/outlook/office.subject#setAsync_subject__options__callback_)) para obter e definir o assunto de um compromisso ou mensagem que o usuário está compondo. Esses métodos assíncronos estão disponíveis apenas para compor os complementos. Para usar esses métodos, certifique-se de configurar o manifesto do complemento adequadamente para Outlook ativar o complemento em formulários de redação.

A propriedade **subject** está disponível para acesso de leitura nos formulários de leitura e de redação de compromissos e de mensagens. Em um formulário de leitura, é possível acessar a propriedade diretamente do objeto pai, como em:

```js
item.subject
```

Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar o assunto ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter o assunto, conforme mostrado abaixo:

```js
item.subject.getAsync
```

A propriedade **subject** está disponível para acesso de gravação somente nos formulários de redação, e não nos de leitura.

Como na maioria dos métodos assíncronos na API javaScript Office, **getAsync** e **setAsync** levam parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira "Passar parâmetros opcionais para métodos assíncronos" em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-subject"></a>Obter o assunto

Esta seção mostra um exemplo de código que obtém o assunto do compromisso ou da mensagem que o usuário está compondo e o exibe. Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Para usar **item.subject.getAsync**, forneça um método de retorno de chamada que verifique o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional _asyncContext_. É possível obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, você poderá obter o assunto como texto sem formatação usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


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


## <a name="set-the-subject"></a>Definir o assunto


Esta seção mostra um exemplo de código que define o assunto do compromisso ou da mensagem que o usuário está compondo. Semelhante ao exemplo anterior, o código a seguir considera uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem.

Para usar **item.subject.setAsync**, especifique uma cadeia de até 255 caracteres no parâmetro de dados. Outra opção é fornecer um método de retorno de chamada e os argumentos para o método de retorno de chamada no parâmetro _asyncContext_. Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de assunto especificada como texto sem formatação, substituindo o assunto existente pelo item.

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


## <a name="see-also"></a>Confira também

- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)   
- [Obter e definir dados de item do Outlook em formulários de leitura ou composição](item-data.md)    
- [Criar suplementos do Outlook para formulários de composição](compose-scenario.md)    
- [Programação assíncrona em Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook](get-set-or-add-recipients.md)  
- [Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook](insert-data-in-the-body.md)   
- [Obter ou definir o local ao criar um compromisso no Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Obter ou definir a hora ao criar um compromisso no Outlook](get-or-set-the-time-of-an-appointment.md)
    
