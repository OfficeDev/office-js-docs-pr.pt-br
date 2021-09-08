---
title: Extrair cadeias de caracteres de entidade de um item do Outlook
description: Saiba como extrair cadeias de caracteres de entidade de um item do Outlook em um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d266795e3794cfa293d073dafc1ca714644faa5b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937504"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a>Extrair cadeias de caracteres de entidade de um item do Outlook

Este artigo descreve como criar o suplemento do Outlook **Exibir entidades**, que extrai instâncias de cadeia de caracteres de entidades conhecidas compatíveis no assunto e no corpo do item do Outlook escolhido. Esse item pode ser um compromisso, uma mensagem de email ou solicitação, resposta ou cancelamento de reunião.

As entidades compatíveis incluem:

- **Endereço**: um endereço postal brasileiro, com pelo menos um subconjunto dos elementos de um número de rua, nome de rua, cidade, estado e CEP.
    
- **Contato**: informações de contato de uma pessoa, no contexto das outras entidades, como endereço ou nome comercial.
    
- **Endereço de email**: um endereço de email SMTP.
    
- **Sugestão de reunião**: uma sugestão de reunião, como uma referência a um evento. Observe que somente as mensagens, e não compromissos, dão suporte à extração de sugestões de reunião.
    
- **Número do telefone**: um número de telefone brasileiro.
    
- **Sugestão de tarefa**: uma sugestão de tarefa, normalmente expressa em uma frase acionável.
    
- **URL**
    
A maioria dessas entidades depende de reconhecimento de linguagem natural, que é baseado no aprendizado da máquina de grandes quantidades de dados. Esse reconhecimento não é determinístico e às vezes depende do contexto do item do Outlook.

O Outlook ativa o suplemento de entidades sempre que o usuário seleciona um compromisso, uma mensagem de email ou uma solicitação, resposta ou cancelamento de reunião para visualização. Durante a inicialização, o suplemento de entidades de exemplo lê todas as instâncias das entidades compatíveis do item atual. 

O suplemento fornece botões para o usuário escolher um tipo de entidade. Quando o usuário seleciona uma entidade, o suplemento exibe instâncias da entidade selecionada no painel do suplemento. As seções a seguir listam o manifesto XML e arquivos HTML e JavaScript do suplemento de entidades, e realçam o código que dá suporte à extração de entidade respectiva.

## <a name="xml-manifest"></a>Manifesto XML

O suplemento de entidade tem duas regras de ativação unidas por um operador lógico OU. 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

Essas regras especificam que o Outlook deve ativar esse suplemento quando o item selecionado no momento no Painel de Leitura ou no inspetor de leitura é um compromisso ou uma mensagem (incluindo mensagem de email ou solicitação, resposta ou cancelamento de reunião).

A seguir temos o manifesto do suplemento das entidades. Ele usa a versão 1.1 do esquema de manifestos de suplementos do Office.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## <a name="html-implementation"></a>Implementação HTML

O arquivo HTML do suplemento de entidades especifica botões para o usuário selecionar cada tipo de entidade e outro botão para limpar instâncias exibidas de uma entidade. Ele inclui um arquivo JavaScript, default_entities.js, que é descrito na próxima seção em [Implementação de JavaScript](#javascript-implementation). O arquivo JavaScript inclui os manipuladores de evento para cada um dos botões.

Observe que todos os suplementos do Outlook devem incluir o office.js. O arquivo HTML a seguir inclui a versão 1.1 do office.js na CDN. 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## <a name="style-sheet"></a>Folha de estilos


O suplemento de entidades usa um arquivo CSS opcional, default_entities.css, para especificar o layout da saída. A seguir temos uma listagem do arquivo CSS.


```CSS
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## <a name="javascript-implementation"></a>Implementação de JavaScript

As seções restantes descrevem como essa amostra (arquivo default_entities.js) extrai entidades conhecidas do assunto e do corpo da mensagem ou do compromisso que o usuário está exibindo.

## <a name="extracting-entities-upon-initialization"></a>Extrair entidades na inicialização

Após o evento [Office.initialize](/javascript/api/office#Office_initialize_reason_), o suplemento de entidades chama o método [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) do item atual. O `getEntities` método retorna a variável global uma matriz de `_MyEntities` instâncias de entidades com suporte. A seguir apresentamos o código JavaScript relacionado.


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a>Extrair endereços


Quando o usuário clica no botão **Obter Endereços**, o manipulador de eventos `myGetAddresses` obtém uma matriz dos endereços da propriedade [addresses](/javascript/api/outlook/office.entities#addresses) do objeto `_MyEntities`, caso algum endereço seja extraído. Cada endereço extraído é armazenado como uma cadeia de caracteres da matriz. `myGetAddresses` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de endereços extraídos. A seguir, apresentamos o código JavaScript relacionado.


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a>Extrair informações de contato


Quando o usuário clica no botão **Obter** Informações de Contato, o manipulador de eventos obtém uma matriz de contatos juntamente com suas informações da propriedade contacts do objeto, se algum tiver sido `myGetContacts` [](/javascript/api/outlook/office.entities#contacts) `_MyEntities` extraído. Cada contato extraído é armazenado como o objeto [Contact](/javascript/api/outlook/office.contact) na matriz. `myGetContacts` obtém mais dados sobre cada contato. Observe que o contexto determina se Outlook pode extrair um contato de um item uma assinatura no final de uma mensagem de email ou pelo menos algumas das informações a seguir teriam que existir nas proximidades do &mdash; contato.


- A cadeia de caracteres que representa o nome do contato da propriedade [Contact.personName](/javascript/api/outlook/office.contact#personName).

- A cadeia de caracteres que representa o nome comercial associado ao contato na propriedade [Contact.businessName](/javascript/api/outlook/office.contact#businessName).

- A matriz de números de telefone associada ao contato na propriedade [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phoneNumbers). Cada número de telefone é representado por um objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber).

- Para cada membro **PhoneNumber** na matriz de números de telefone, a cadeia de caracteres que representa o número de telefone da propriedade [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString).

- A matriz de URLs associada ao contato na propriedade [Contact.urls](/javascript/api/outlook/office.contact#urls). Cada URL é representada como uma cadeia de caracteres em um membro da matriz.

- A matriz de endereços de email associada ao contato na propriedade [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailAddresses). Cada endereço de email é representado como uma cadeia de caracteres em um membro da matriz.

- A matriz de endereços postais associada ao contato na propriedade [Contact.addresses](/javascript/api/outlook/office.contact#addresses). Cada endereço postal é representado como uma cadeia de caracteres em um membro da matriz.

`myGetContacts` forma uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada contato. A seguir apresentamos o código JavaScript relacionado.




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a>Extrair endereços de email


Quando o usuário clica no botão **Obter Endereços de Email**, o manipulador de eventos `myGetEmailAddresses` obtém uma matriz de endereços de email SMTP na propriedade [emailAddresses](/javascript/api/outlook/office.entities#emailAddresses) do objeto `_MyEntities`, caso algum seja extraído. Cada endereço de email extraído é armazenado como uma cadeia de caracteres na matriz. `myGetEmailAddresses` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de endereços de email extraídos. A seguir apresentamos o código JavaScript relacionado.


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a>Extrair sugestões de reunião


Quando o usuário clica no botão **Obter Sugestões de Reunião**, o manipulador de eventos `myGetMeetingSuggestions` obtém uma matriz de sugestões de reunião da propriedade [meetingSuggestions](/javascript/api/outlook/office.entities#meetingSuggestions) do objeto `_MyEntities`, caso algum seja extraído.


 > [!NOTE]
 > Somente mensagens, mas não compromissos, suportam o `MeetingSuggestion` tipo de entidade.

Cada sugestão de reunião extraída é armazenada como um objeto [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) na matriz. `myGetMeetingSuggestions` obtém dados adicionais sobre cada sugestão de reunião:


- A cadeia de caracteres que foi identificada como uma sugestão de reunião na propriedade [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingString).

- A matriz de participantes da reunião na propriedade [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees). Cada participante é representado por um objeto [EmailUser](/javascript/api/outlook/office.emailuser).

- Para cada participante, o nome na propriedade [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName).

- Para cada participante, o endereço SMTP na propriedade [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress).

- A cadeia de caracteres que representa a localização de sugestão de reunião na propriedade [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location).

- A cadeia de caracteres que representa o assunto da sugestão de reunião na propriedade [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject).

- A cadeia de caracteres que representa a hora de início da sugestão de reunião na propriedade [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start).

- A cadeia de caracteres que representa a hora de término da sugestão de reunião na propriedade [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end).

`myGetMeetingSuggestions` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada uma das sugestões de reunião. A seguir apresentamos o código JavaScript relacionado.




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-phone-numbers"></a>Extrair números de telefone


Quando o usuário clica no botão **Obter Números de Telefone**, o manipulador de eventos `myGetPhoneNumbers` obtém uma matriz de números de telefone na propriedade [phoneNumbers](/javascript/api/outlook/office.entities#phoneNumbers) do objeto `_MyEntities`, caso algum seja extraído. Cada número de telefone extraído é armazenado como objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber) na matriz. `myGetPhoneNumbers` obtém mais dados sobre cada número de telefone:


- A cadeia de caracteres que representa o tipo de número de telefone, por exemplo, número de telefone residencial, na propriedade [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type).

- A cadeia de caracteres que representa o número de telefone real na propriedade [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString).

- A cadeia de caracteres que foi originalmente identificada como o número de telefone na propriedade [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalPhoneString).

`myGetPhoneNumbers` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada um dos números de telefone. A seguir apresentamos o código JavaScript relacionado.




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-task-suggestions"></a>Extrair sugestões de tarefa


Quando o usuário clica no botão **Obter Sugestões de Tarefa**, o manipulador de eventos `myGetTaskSuggestions` obtém uma matriz de sugestões de tarefa na propriedade [taskSuggestions](/javascript/api/outlook/office.entities#taskSuggestions) do objeto `_MyEntities`, caso algum seja extraído. Cada sugestão de tarefa extraída é armazenada como um objeto [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) da matriz. `myGetTaskSuggestions` obtém dados adicionais sobre cada sugestão de tarefa:


- A cadeia de caracteres que foi originalmente identificada como uma sugestão de tarefa na propriedade [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskString).

- A matriz dos destinatários da tarefa na propriedade [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees). Cada destinatário é representado por um objeto [EmailUser](/javascript/api/outlook/office.emailuser).

- Para cada destinatário, o nome na propriedade [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName).

- Para cada destinatário, o endereço SMTP da propriedade [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress).

`myGetTaskSuggestions` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada sugestão de tarefa. A seguir apresentamos o código JavaScript relacionado.




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-urls"></a>Extrair URLs


Quando o usuário clica no botão **Obter URLs**, o manipulador de eventos `myGetUrls` obtém uma matriz de URLs na propriedade [urls](/javascript/api/outlook/office.entities#urls) do objeto `_MyEntities`, caso algum seja extraído. Cada URL extraída é armazenada como uma cadeia de caracteres na matriz. `myGetUrls` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de URLs extraídas.


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a>Limpar cadeias de caracteres de entidade exibidas


Por fim, o suplemento de entidades especifica um manipulador de eventos `myClearEntitiesBox` que limpa as cadeias de caracteres exibidas. A seguir apresentamos o código relacionado.


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a>Listagem de JavaScript


A seguir apresentamos uma listagem completa da implementação do JavaScript.


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de leitura](read-scenario.md)
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)
- [Método item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
