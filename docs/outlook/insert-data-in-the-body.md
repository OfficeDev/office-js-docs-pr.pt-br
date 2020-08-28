---
title: Inserir dados no corpo de um suplemento do Outlook
description: Saiba como inserir dados no corpo de um compromisso ou mensagem em um suplemento do Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 0e875619520ee309dec97b2db60ed49c29b2a463
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293867"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="47d63-103">Inserir dados no corpo ao compor um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-103">Insert data in the body when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="47d63-p101">Você pode usar os métodos assíncronos ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) e [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) para obter o tipo de corpo e inserir dados no corpo de um item de compromisso ou de uma mensagem que o usuário está compondo. Esses métodos assíncronos estão disponíveis somente para suplementos de composição. Para usar esses métodos, verifique se você configurou o manifesto do suplemento adequadamente para o Outlook ativar o suplemento nos formulários de composição, conforme descrito em [Criar suplementos do Outlook para formulários de composição](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="47d63-p101">You can use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="47d63-p102">No Outlook, um usuário pode criar uma mensagem em texto, HTML ou Rich Text Format (RTF) e pode criar um compromisso no formato HTML. Antes de inserir, você deve sempre verificar primeiro o formato de item suportado chamando **getTypeAsync**, já que talvez seja necessário executar etapas adicionais. O valor que **getTypeAsync** retorna depende do formato de item original, bem como o suporte do sistema operacional do dispositivo e do aplicativo a ser editado no formato HTML (1). Em seguida, defina o parâmetro  _coercionType_ de **prependAsync** ou **setSelectedDataAsync** de acordo (2) para inserir os dados, conforme mostrado na tabela a seguir. Se você não especificar um argumento, **prependAsync** e **setSelectedDataAsync** assumem que os dados inseridos estão no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="47d63-p102">In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling **getTypeAsync**, as you may need to take additional steps. The value that **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and application to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.</span></span>

<br/>

|<span data-ttu-id="47d63-111">**Dados a inserir**</span><span class="sxs-lookup"><span data-stu-id="47d63-111">**Data to insert**</span></span>|<span data-ttu-id="47d63-112">**Formato de item retornado por getTypeAsync**</span><span class="sxs-lookup"><span data-stu-id="47d63-112">**Item format returned by getTypeAsync**</span></span>|<span data-ttu-id="47d63-113">**Usar este coercionType**</span><span class="sxs-lookup"><span data-stu-id="47d63-113">**Use this coercionType**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="47d63-114">Texto</span><span class="sxs-lookup"><span data-stu-id="47d63-114">Text</span></span>|<span data-ttu-id="47d63-115">Texto (1)</span><span class="sxs-lookup"><span data-stu-id="47d63-115">Text (1)</span></span>|<span data-ttu-id="47d63-116">Texto</span><span class="sxs-lookup"><span data-stu-id="47d63-116">Text</span></span>|
|<span data-ttu-id="47d63-117">HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-117">HTML</span></span>|<span data-ttu-id="47d63-118">Texto (1)</span><span class="sxs-lookup"><span data-stu-id="47d63-118">Text (1)</span></span>|<span data-ttu-id="47d63-119">Texto (2)</span><span class="sxs-lookup"><span data-stu-id="47d63-119">Text (2)</span></span>|
|<span data-ttu-id="47d63-120">Texto</span><span class="sxs-lookup"><span data-stu-id="47d63-120">Text</span></span>|<span data-ttu-id="47d63-121">HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-121">HTML</span></span>|<span data-ttu-id="47d63-122">Texto/HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-122">Text/HTML</span></span>|
|<span data-ttu-id="47d63-123">HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-123">HTML</span></span>|<span data-ttu-id="47d63-124">HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-124">HTML</span></span> |<span data-ttu-id="47d63-125">HTML</span><span class="sxs-lookup"><span data-stu-id="47d63-125">HTML</span></span>|

1.  <span data-ttu-id="47d63-126">Em tablets e smartphones, **getTypeAsync** retorna **Office. MailboxEnums. BodyType. Text** se o sistema operacional ou o aplicativo não oferecer suporte à edição de um item, que foi originalmente criado no HTML, no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="47d63-126">On tablets and smartphones, **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or application does not support editing an item, which was originally created in HTML, in HTML format.</span></span>

2.  <span data-ttu-id="47d63-p103">Se os dados a serem inseridos forem HTML e **getTypeAsync** retornar um tipo de texto para esse item, reorganize seus dados como texto e insira-o com **Office. MailboxEnums. BodyType. Text** como _coercionType_. Se você simplesmente inserir os dados HTML com um tipo de coerção de texto, o aplicativo exibirá as marcas HTML como texto. Se você tentar inserir os dados HTML com **Office.MailboxEnums.BodyType.Html** como _coercionType_, receberá um erro.</span><span class="sxs-lookup"><span data-stu-id="47d63-p103">If your data to insert is HTML and **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the application would display the HTML tags as text. If you attempt to insert the HTML data with **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.</span></span>

<span data-ttu-id="47d63-p104">Além de  _coercionType_, assim como a maioria dos métodos assíncronos na API JavaScript do Office, o **getTypeAsync**, o **prependAsync** e o **setSelectedDataAsync** usam outros parâmetros de entrada opcionais. Para obter mais informações sobre como especificar esses parâmetros de entrada opcionais, consulte [passando parâmetros opcionais para métodos assíncronos](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) em [programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="47d63-p104">In addition to  _coercionType_, as with most asynchronous methods in the Office JavaScript API, **getTypeAsync**, **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="insert-data-at-the-current-cursor-position"></a><span data-ttu-id="47d63-132">Inserir dados na posição atual do cursor</span><span class="sxs-lookup"><span data-stu-id="47d63-132">Insert data at the current cursor position</span></span>


<span data-ttu-id="47d63-133">Esta seção mostra um exemplo de código que usa **getTypeAsync** para verificar o tipo de corpo do item que está sendo redigido e usa **setSelectedDataAsync** para inserir dados no local atual do cursor.</span><span class="sxs-lookup"><span data-stu-id="47d63-133">This section shows a code sample that uses **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.</span></span>

<span data-ttu-id="47d63-p105">Você pode transmitir um método de retorno e parâmetros de entrada opcionais para **getTypeAsync** e obter status e resultados no parâmetro de saída _asyncResult_. Se o método for bem-sucedido, você poderá obter o tipo do corpo do item na propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value), que é “texto” ou “html”.</span><span class="sxs-lookup"><span data-stu-id="47d63-p105">You can pass a callback method and optional input parameters to **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property, which is either "text" or "html".</span></span>

<span data-ttu-id="47d63-p106">Você deve transmitir uma cadeia de caracteres de dados como um parâmetro de entrada para **setSelectedDataAsync**. Dependendo do tipo do corpo do item, é possível especificar essa cadeia de caracteres de dados no formato HTML ou de texto adequadamente. Conforme mencionado acima, outra opção é especificar o tipo de dados a ser inserido no parâmetro _coercionType_. Além disso, é possível fornecer um método de retorno de chamada e seus parâmetros como parâmetros de entrada opcionais.</span><span class="sxs-lookup"><span data-stu-id="47d63-p106">You must pass a data string as an input parameter to **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.</span></span>

<span data-ttu-id="47d63-p107">Se o usuário não tiver colocado o cursor no corpo do item, **setSelectedDataAsync** inserirá os dados na parte superior do corpo. Se o usuário tiver selecionado texto no corpo do item, **setSelectedDataAsync** substituirá o texto selecionado pelos dados que você especificar. Observe que **setSelectedDataAsync** pode dar erro se o usuário estiver mudando a posição do cursor ao escrever o item simultaneamente. A quantidade máxima de caracteres que é possível inserir de cada vez é de um milhão.</span><span class="sxs-lookup"><span data-stu-id="47d63-p107">If the user hasn't placed the cursor in the item body, **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.</span></span>

<span data-ttu-id="47d63-144">Este exemplo de código assume uma regra no manifesto do suplemento que ativa o suplemento em um formulário de redação para um compromisso ou uma mensagem, conforme mostrado abaixo.</span><span class="sxs-lookup"><span data-stu-id="47d63-144">This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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


## <a name="insert-data-at-the-beginning-of-the-item-body"></a><span data-ttu-id="47d63-145">Inserir dados no início do corpo do item</span><span class="sxs-lookup"><span data-stu-id="47d63-145">Insert data at the beginning of the item body</span></span>


<span data-ttu-id="47d63-p108">Como alternativa, você pode usar **prependAsync** para inserir dados no início do corpo do item e desconsiderar o local atual do cursor. Não sendo o ponto de inserção, **prependAsync** e **setSelectedDataAsync** se comportam de maneiras semelhantes:</span><span class="sxs-lookup"><span data-stu-id="47d63-p108">Alternatively, you can use **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:</span></span>


- <span data-ttu-id="47d63-148">Se você estiver anexando dados HTML ao corpo da mensagem, primeiro deverá verificar o tipo do corpo da mensagem para evitar anexar dados HTML a uma mensagem no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="47d63-148">If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.</span></span>
    
- <span data-ttu-id="47d63-149">Forneça os itens a seguir como parâmetros de entrada para **prependAsync**: uma cadeia de caracteres de dados em formato de texto ou HTML e, opcionalmente, o formato dos dados a ser inserido, um método de retorno de chamada e seus parâmetros.</span><span class="sxs-lookup"><span data-stu-id="47d63-149">Provide the following as input parameters to **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.</span></span>
    
- <span data-ttu-id="47d63-150">O número máximo de caracteres que você pode anexar no início de cada vez é um milhão.</span><span class="sxs-lookup"><span data-stu-id="47d63-150">The maximum number of characters you can prepend at one time is 1,000,000 characters.</span></span>
    
<span data-ttu-id="47d63-p109">O código JavaScript a seguir faz parte de um suplemento de exemplo que é ativado nos formulários de redação de compromissos e mensagens. O exemplo chama **getTypeAsync** para verificar o tipo do corpo do item, insere dados HTML na parte superior do corpo do item se este for um compromisso ou uma mensagem em HTML. Caso contrário, ele insere os dados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="47d63-p109">The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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


## <a name="see-also"></a><span data-ttu-id="47d63-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="47d63-153">See also</span></span>

- [<span data-ttu-id="47d63-154">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-154">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="47d63-155">Obter e definir dados de item do Outlook em formulários de leitura ou composição</span><span class="sxs-lookup"><span data-stu-id="47d63-155">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="47d63-156">Criar suplementos do Outlook para formulários de composição</span><span class="sxs-lookup"><span data-stu-id="47d63-156">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="47d63-157">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="47d63-157">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="47d63-158">Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-158">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="47d63-159">Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-159">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)  
- [<span data-ttu-id="47d63-160">Obter ou definir o local ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-160">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="47d63-161">Obter ou definir a hora ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="47d63-161">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
