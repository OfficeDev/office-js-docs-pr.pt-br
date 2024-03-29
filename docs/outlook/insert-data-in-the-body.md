---
title: Inserir dados no corpo de um suplemento do Outlook
description: Saiba como inserir dados no corpo de um compromisso ou mensagem em um suplemento do Outlook.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7319a3bb41d857fcae32ea118a3f3e60197bf751
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958324"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>Inserir dados no corpo ao compor um compromisso ou uma mensagem no Outlook

Você pode usar os métodos assíncronos ([Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1)), [Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1)), [Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)) e [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))) para obter o tipo de corpo e inserir dados no corpo de um item de compromisso ou de uma mensagem que o usuário está compondo. Esses métodos assíncronos estão disponíveis somente para suplementos de composição. Para usar esses métodos, verifique se você configurou o manifesto do suplemento adequadamente para o Outlook ativar o suplemento nos formulários de composição, conforme descrito em [Criar suplementos do Outlook para formulários de composição](compose-scenario.md).

No Outlook, um usuário pode criar uma mensagem em texto, HTML ou RTF (Rich Text Format) e pode criar um compromisso no formato HTML. Antes de inserir, você sempre deve verificar primeiro o formato de item com suporte chamando **getTypeAsync**, pois talvez seja necessário executar etapas adicionais. O valor **retornado por getTypeAsync** depende do formato de item original, bem como do suporte do sistema operacional do dispositivo e do aplicativo para edição no formato HTML (1). Em seguida, defina _coercionType_ parâmetro da **prependAsync** ou **setSelectedDataAsync** adequadamente (2) para inserir os dados, conforme mostrado na tabela a seguir. Se você não especifica um argumento, **prependAsync** e **setSelectedDataAsync** pressupõem que os dados a inserir estão no formato de texto.

|Dados a inserir|Formato de item retornado por getTypeAsync|Usar este coercionType|
|:-----|:-----|:-----|
|Texto|Texto (1)|Texto|
|HTML|Texto (1)|Texto (2)|
|Texto|HTML|Texto/HTML|
|HTML|HTML |HTML|

1. Em tablets e smartphones, **getTypeAsync** retornará **Office.MailboxEnums.BodyType.Text** se o sistema operacional ou o aplicativo não for compatível com a edição de um item, que foi originalmente criado em HTML, no formato HTML.

1. Se os dados a serem inseridos for HTML e **getTypeAsync** retornar um tipo de texto para esse item, reorganize seus dados como texto e insira-os com **Office.MailboxEnums.BodyType.Text** como _coercionType_. Se você simplesmente inserir os dados HTML com um tipo de coerção de texto, o aplicativo exibirá as marcas HTML como texto. Se você tentar inserir os dados HTML com **Office.MailboxEnums.BodyType.Html** como _coercionType_, receberá um erro.

Além de  _coercionType_, assim como na maioria dos métodos assíncronos na API JavaScript do Office, **getTypeAsync**, **prependAsync** e **setSelectedDataAsync** usam outros parâmetros de entrada opcionais. Para saber mais sobre como especificar esses parâmetros de entrada opcionais, confira [Passar parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="insert-data-at-the-current-cursor-position"></a>Inserir dados na posição atual do cursor

Esta seção mostra um exemplo de código que usa **getTypeAsync** para verificar o tipo de corpo do item que está sendo redigido e usa **setSelectedDataAsync** para inserir dados no local atual do cursor.

Você pode passar uma função de retorno de chamada e parâmetros de entrada opcionais para **getTypeAsync** e obter qualquer status e resultados no  _parâmetro de saída asyncResult_ . Se o método for bem-sucedido, você poderá obter o tipo do corpo do item na propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member), que é "texto" ou "HTML".

Você deve passar uma cadeia de caracteres de dados como um parâmetro de entrada para **setSelectedDataAsync**. Dependendo do tipo do corpo do item, você pode especificar essa cadeia de caracteres de dados no formato HTML ou de texto adequadamente. Conforme mencionado acima, opcionalmente, você pode especificar o tipo de dados a ser inserido no parâmetro _coercionType_. Além disso, você pode fornecer uma função de retorno de chamada e qualquer um de seus parâmetros como parâmetros de entrada opcionais.

Se o usuário não tiver colocado o cursor no corpo do item, **setSelectedDataAsync** inserirá os dados na parte superior do corpo. Se o usuário tiver selecionado texto no corpo do item, **setSelectedDataAsync** substituirá o texto selecionado pelos dados que você especificar. Observe que **setSelectedDataAsync** pode dar erro se o usuário estiver mudando a posição do cursor ao escrever o item simultaneamente. A quantidade máxima de caracteres que é possível inserir de cada vez é de um milhão.

Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}

// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="insert-data-at-the-beginning-of-the-item-body"></a>Inserir dados no início do corpo do item

Como alternativa, você pode usar **prependAsync** para inserir dados no início do corpo do item e desconsiderar o local atual do cursor. Não sendo o ponto de inserção, **prependAsync** e **setSelectedDataAsync** se comportam de maneiras semelhantes:

- Se você estiver pré-anexando dados HTML em um corpo de mensagem, primeiro verifique o tipo do corpo da mensagem para evitar o acréscimo de dados HTML a uma mensagem no formato de texto.

- Forneça o seguinte como parâmetros de entrada para **prependAsync**: uma cadeia de caracteres de dados no formato TEXTO ou HTML e, opcionalmente, o formato dos dados a serem inseridos, uma função de retorno de chamada e qualquer um de seus parâmetros.

- O número máximo de caracteres que você pode anexar no início de cada vez é um milhão.

O código JavaScript a seguir faz parte de um suplemento de exemplo que é ativado nos formulários de redação de compromissos e mensagens. O exemplo chama **getTypeAsync** para verificar o tipo do corpo do item, insere dados HTML na parte superior do corpo do item se este for um compromisso ou uma mensagem em HTML. Caso contrário, ele insere os dados no formato de texto.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });
}

// Writes to a div with id='message' on the page.
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
- [Obter ou definir o local ao criar um compromisso no Outlook](get-or-set-the-location-of-an-appointment.md)
- [Obter ou definir a hora ao criar um compromisso no Outlook](get-or-set-the-time-of-an-appointment.md)
