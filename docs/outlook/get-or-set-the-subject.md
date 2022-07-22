---
title: Obter ou definir o assunto em um suplemento do Outlook
description: Saiba como obter ou definir o assunto de uma mensagem ou compromisso em um suplemento do Outlook.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf221b03753fd76966eb5c6270da68e94abfe0f9
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959067"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Obter ou definir o assunto ao compor um compromisso ou uma mensagem no Outlook

A API JavaScript do Office fornece métodos assíncronos ([subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) e [subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))) para obter e definir o assunto de um compromisso ou mensagem que o usuário está redigindo. Esses métodos assíncronos estão disponíveis apenas para compor suplementos. Para usar esses métodos, verifique se você configurou o manifesto do suplemento adequadamente para o Outlook ativar o suplemento em formulários de composição.

A propriedade **subject** está disponível para acesso de leitura nos formulários de leitura e de redação de compromissos e de mensagens. Em um formulário de leitura, é possível acessar a propriedade diretamente do objeto pai, como em:

```js
item.subject
```

Mas em um formulário de redação, como o usuário e o suplemento podem inserir ou mudar o assunto ao mesmo tempo, você deve usar o método assíncrono **getAsync** para obter o assunto, conforme mostrado abaixo:

```js
item.subject.getAsync
```

A propriedade **subject** está disponível para acesso de gravação somente nos formulários de redação, e não nos de leitura.

Assim como na maioria dos métodos assíncronos na API JavaScript do Office, **getAsync** e **setAsync** usam parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira "Passar parâmetros opcionais para métodos assíncronos" em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-subject"></a>Obter o assunto

Esta seção mostra um exemplo de código que obtém o assunto do compromisso ou da mensagem que o usuário está compondo e o exibe. Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Para usar **item.subject.getAsync**, forneça uma função de retorno de chamada que verifica o status e o resultado da chamada assíncrona. Você pode fornecer os argumentos necessários para a função de retorno de chamada por meio do  _parâmetro opcional asyncContext_ . Você pode obter o status, os resultados e eventuais erros usando o parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, você poderá obter o assunto como texto sem formatação usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

Para usar **item.subject.setAsync**, especifique uma cadeia de caracteres de até 255 caracteres no parâmetro de dados. Opcionalmente, você pode fornecer uma função de retorno de chamada e quaisquer argumentos para a função de retorno de chamada no  _parâmetro asyncContext_ . Você deve verificar o status, o resultado e eventuais mensagens de erro no parâmetro de saída _asyncResult_ do retorno de chamada. Se a chamada assíncrona for bem-sucedida, **setAsync** inserirá a cadeia de caracteres de assunto especificada como texto sem formatação, substituindo o assunto existente pelo item.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    const today = new Date();
    let subject;

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
