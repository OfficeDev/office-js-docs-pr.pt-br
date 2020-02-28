---
title: Extrair cadeias de caracteres de entidade de um item do Outlook
description: Saiba como extrair cadeias de caracteres de entidade de um item do Outlook em um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 0a9a41d0b479420c0754c0e0d283982082a1452f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325451"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a><span data-ttu-id="17038-103">Extrair cadeias de caracteres de entidade de um item do Outlook</span><span class="sxs-lookup"><span data-stu-id="17038-103">Extract entity strings from an Outlook item</span></span>

<span data-ttu-id="17038-p101">Este artigo descreve como criar o suplemento do Outlook **Exibir entidades**, que extrai instâncias de cadeia de caracteres de entidades conhecidas compatíveis no assunto e no corpo do item do Outlook escolhido. Esse item pode ser um compromisso, uma mensagem de email ou solicitação, resposta ou cancelamento de reunião.</span><span class="sxs-lookup"><span data-stu-id="17038-p101">This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.</span></span>

<span data-ttu-id="17038-106">As entidades compatíveis incluem:</span><span class="sxs-lookup"><span data-stu-id="17038-106">The supported entities include:</span></span>

- <span data-ttu-id="17038-107">**Endereço**: um endereço postal brasileiro, com pelo menos um subconjunto dos elementos de um número de rua, nome de rua, cidade, estado e CEP.</span><span class="sxs-lookup"><span data-stu-id="17038-107">**Address**: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.</span></span>
    
- <span data-ttu-id="17038-108">**Contato**: informações de contato de uma pessoa, no contexto das outras entidades, como endereço ou nome comercial.</span><span class="sxs-lookup"><span data-stu-id="17038-108">**Contact**: A person's contact information, in the context of other entities such as an address or business name.</span></span>
    
- <span data-ttu-id="17038-109">**Endereço de email**: um endereço de email SMTP.</span><span class="sxs-lookup"><span data-stu-id="17038-109">**Email address**: An SMTP email address.</span></span>
    
- <span data-ttu-id="17038-p102">**Sugestão de reunião**: uma sugestão de reunião, como uma referência a um evento. Observe que somente as mensagens, e não compromissos, dão suporte à extração de sugestões de reunião.</span><span class="sxs-lookup"><span data-stu-id="17038-p102">**Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.</span></span>
    
- <span data-ttu-id="17038-112">**Número do telefone**: um número de telefone brasileiro.</span><span class="sxs-lookup"><span data-stu-id="17038-112">**Phone number**: A North American phone number.</span></span>
    
- <span data-ttu-id="17038-113">**Sugestão de tarefa**: uma sugestão de tarefa, normalmente expressa em uma frase acionável.</span><span class="sxs-lookup"><span data-stu-id="17038-113">**Task suggestion**: A task suggestion, typically expressed in an actionable phrase.</span></span>
    
- <span data-ttu-id="17038-114">**URL**</span><span class="sxs-lookup"><span data-stu-id="17038-114">**URL**</span></span>
    
<span data-ttu-id="17038-p103">A maioria dessas entidades depende de reconhecimento de linguagem natural, que é baseado no aprendizado da máquina de grandes quantidades de dados. Esse reconhecimento não é determinístico e às vezes depende do contexto do item do Outlook.</span><span class="sxs-lookup"><span data-stu-id="17038-p103">Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.</span></span>

<span data-ttu-id="17038-p104">O Outlook ativa o suplemento de entidades sempre que o usuário seleciona um compromisso, uma mensagem de email ou uma solicitação, resposta ou cancelamento de reunião para visualização. Durante a inicialização, o suplemento de entidades de exemplo lê todas as instâncias das entidades compatíveis do item atual.</span><span class="sxs-lookup"><span data-stu-id="17038-p104">Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.</span></span> 

<span data-ttu-id="17038-p105">O suplemento fornece botões para o usuário escolher um tipo de entidade. Quando o usuário seleciona uma entidade, o suplemento exibe instâncias da entidade selecionada no painel do suplemento. As seções a seguir listam o manifesto XML e arquivos HTML e JavaScript do suplemento de entidades, e realçam o código que dá suporte à extração de entidade respectiva.</span><span class="sxs-lookup"><span data-stu-id="17038-p105">The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.</span></span>

## <a name="xml-manifest"></a><span data-ttu-id="17038-122">Manifesto XML</span><span class="sxs-lookup"><span data-stu-id="17038-122">XML manifest</span></span>

<span data-ttu-id="17038-123">O suplemento de entidade tem duas regras de ativação unidas por um operador lógico OU.</span><span class="sxs-lookup"><span data-stu-id="17038-123">The entities add-in has two activation rules joined by a logical OR operation.</span></span> 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

<span data-ttu-id="17038-124">Essas regras especificam que o Outlook deve ativar esse suplemento quando o item selecionado no momento no Painel de Leitura ou no inspetor de leitura é um compromisso ou uma mensagem (incluindo mensagem de email ou solicitação, resposta ou cancelamento de reunião).</span><span class="sxs-lookup"><span data-stu-id="17038-124">These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).</span></span>

<span data-ttu-id="17038-p106">A seguir temos o manifesto do suplemento das entidades. Ele usa a versão 1.1 do esquema de manifestos de suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="17038-p106">The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.</span></span>

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


## <a name="html-implementation"></a><span data-ttu-id="17038-127">Implementação HTML</span><span class="sxs-lookup"><span data-stu-id="17038-127">HTML implementation</span></span>

<span data-ttu-id="17038-p107">O arquivo HTML do suplemento de entidades especifica botões para o usuário selecionar cada tipo de entidade e outro botão para limpar instâncias exibidas de uma entidade. Ele inclui um arquivo JavaScript, default_entities.js, que é descrito na próxima seção em [Implementação de JavaScript](#javascript-implementation). O arquivo JavaScript inclui os manipuladores de evento para cada um dos botões.</span><span class="sxs-lookup"><span data-stu-id="17038-p107">The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.</span></span>

<span data-ttu-id="17038-p108">Observe que todos os suplementos do Outlook devem incluir o office.js. O arquivo HTML a seguir inclui a versão 1.1 do office.js na CDN.</span><span class="sxs-lookup"><span data-stu-id="17038-p108">Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN.</span></span> 

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


## <a name="style-sheet"></a><span data-ttu-id="17038-133">Folha de estilos</span><span class="sxs-lookup"><span data-stu-id="17038-133">Style sheet</span></span>


<span data-ttu-id="17038-p109">O suplemento de entidades usa um arquivo CSS opcional, default_entities.css, para especificar o layout da saída. A seguir temos uma listagem do arquivo CSS.</span><span class="sxs-lookup"><span data-stu-id="17038-p109">The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.</span></span>


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


## <a name="javascript-implementation"></a><span data-ttu-id="17038-136">Implementação de JavaScript</span><span class="sxs-lookup"><span data-stu-id="17038-136">JavaScript implementation</span></span>

<span data-ttu-id="17038-137">As seções restantes descrevem como essa amostra (arquivo default_entities.js) extrai entidades conhecidas do assunto e do corpo da mensagem ou do compromisso que o usuário está exibindo.</span><span class="sxs-lookup"><span data-stu-id="17038-137">The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing.</span></span>

## <a name="extracting-entities-upon-initialization"></a><span data-ttu-id="17038-138">Extrair entidades na inicialização</span><span class="sxs-lookup"><span data-stu-id="17038-138">Extracting entities upon initialization</span></span>

<span data-ttu-id="17038-139">Após o evento [Office.initialize](/javascript/api/office#office-initialize-reason-), o suplemento de entidades chama o método [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) do item atual.</span><span class="sxs-lookup"><span data-stu-id="17038-139">Upon the [Office.initialize](/javascript/api/office#office-initialize-reason-) event, the entities add-in calls the [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method of the current item.</span></span> <span data-ttu-id="17038-140">O `getEntities` método retorna a variável `_MyEntities` global uma matriz de instâncias de entidades com suporte.</span><span class="sxs-lookup"><span data-stu-id="17038-140">The `getEntities` method returns the global variable `_MyEntities` an array of instances of supported entities.</span></span> <span data-ttu-id="17038-141">A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-141">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-addresses"></a><span data-ttu-id="17038-142">Extrair endereços</span><span class="sxs-lookup"><span data-stu-id="17038-142">Extracting addresses</span></span>


<span data-ttu-id="17038-143">Quando o usuário clica no botão **Obter Endereços**, o manipulador de eventos `myGetAddresses` obtém uma matriz dos endereços da propriedade [addresses](/javascript/api/outlook/office.entities#addresses) do objeto `_MyEntities`, caso algum endereço seja extraído.</span><span class="sxs-lookup"><span data-stu-id="17038-143">When the user clicks the **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](/javascript/api/outlook/office.entities#addresses) property of the `_MyEntities` object, if any address was extracted.</span></span> <span data-ttu-id="17038-144">Cada endereço extraído é armazenado como uma cadeia de caracteres da matriz.</span><span class="sxs-lookup"><span data-stu-id="17038-144">Each extracted address is stored as a string in the array.</span></span> <span data-ttu-id="17038-145">`myGetAddresses` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de endereços extraídos.</span><span class="sxs-lookup"><span data-stu-id="17038-145">`myGetAddresses` forms a local HTML string in `htmlText` to display the list of extracted addresses.</span></span> <span data-ttu-id="17038-146">A seguir, apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-146">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-contact-information"></a><span data-ttu-id="17038-147">Extrair informações de contato</span><span class="sxs-lookup"><span data-stu-id="17038-147">Extracting contact information</span></span>


<span data-ttu-id="17038-p112">Quando o usuário clica no botão **Obter Informações de Contato**, o manipulador de eventos `myGetContacts` obtém uma matriz de contatos em conjunto com suas informações da propriedade [contacts](/javascript/api/outlook/office.entities#contacts) do objeto `_MyEntities`, caso algum seja extraído. Cada contato extraído é armazenado como o objeto [Contact](/javascript/api/outlook/office.contact) na matriz. `myGetContacts` obtém mais dados sobre cada contato. Observe que o contexto determina se o Outlook pode extrair um contato de um item&mdash; uma assinatura no final de uma mensagem de email ou ao menos algumas das informações seguintes teriam que existir perto do contato:</span><span class="sxs-lookup"><span data-stu-id="17038-p112">When the user clicks the **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](/javascript/api/outlook/office.entities#contacts) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](/javascript/api/outlook/office.contact) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item&mdash;a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:</span></span>


- <span data-ttu-id="17038-152">A cadeia de caracteres que representa o nome do contato da propriedade [Contact.personName](/javascript/api/outlook/office.contact#personname).</span><span class="sxs-lookup"><span data-stu-id="17038-152">The string representing the contact's name from the [Contact.personName](/javascript/api/outlook/office.contact#personname) property.</span></span>

- <span data-ttu-id="17038-153">A cadeia de caracteres que representa o nome comercial associado ao contato na propriedade [Contact.businessName](/javascript/api/outlook/office.contact#businessname).</span><span class="sxs-lookup"><span data-stu-id="17038-153">The string representing the company name associated with the contact from the [Contact.businessName](/javascript/api/outlook/office.contact#businessname) property.</span></span>

- <span data-ttu-id="17038-p113">A matriz de números de telefone associada ao contato na propriedade [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers). Cada número de telefone é representado por um objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber).</span><span class="sxs-lookup"><span data-stu-id="17038-p113">The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.</span></span>

- <span data-ttu-id="17038-156">Para cada membro **PhoneNumber** na matriz de números de telefone, a cadeia de caracteres que representa o número de telefone da propriedade [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring).</span><span class="sxs-lookup"><span data-stu-id="17038-156">For each **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="17038-p114">A matriz de URLs associada ao contato na propriedade [Contact.urls](/javascript/api/outlook/office.contact#urls). Cada URL é representada como uma cadeia de caracteres em um membro da matriz.</span><span class="sxs-lookup"><span data-stu-id="17038-p114">The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#urls) property. Each URL is represented as a string in an array member.</span></span>

- <span data-ttu-id="17038-p115">A matriz de endereços de email associada ao contato na propriedade [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses). Cada endereço de email é representado como uma cadeia de caracteres em um membro da matriz.</span><span class="sxs-lookup"><span data-stu-id="17038-p115">The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) property. Each email address is represented as a string in an array member.</span></span>

- <span data-ttu-id="17038-p116">A matriz de endereços postais associada ao contato na propriedade [Contact.addresses](/javascript/api/outlook/office.contact#addresses). Cada endereço postal é representado como uma cadeia de caracteres em um membro da matriz.</span><span class="sxs-lookup"><span data-stu-id="17038-p116">The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#addresses) property. Each postal address is represented as a string in an array member.</span></span>

<span data-ttu-id="17038-p117">`myGetContacts` forma uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada contato. A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p117">`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-email-addresses"></a><span data-ttu-id="17038-165">Extrair endereços de email</span><span class="sxs-lookup"><span data-stu-id="17038-165">Extracting email addresses</span></span>


<span data-ttu-id="17038-p118">Quando o usuário clica no botão **Obter Endereços de Email**, o manipulador de eventos `myGetEmailAddresses` obtém uma matriz de endereços de email SMTP na propriedade [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) do objeto `_MyEntities`, caso algum seja extraído. Cada endereço de email extraído é armazenado como uma cadeia de caracteres na matriz. `myGetEmailAddresses` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de endereços de email extraídos. A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p118">When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.</span></span>


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


## <a name="extracting-meeting-suggestions"></a><span data-ttu-id="17038-170">Extrair sugestões de reunião</span><span class="sxs-lookup"><span data-stu-id="17038-170">Extracting meeting suggestions</span></span>


<span data-ttu-id="17038-171">Quando o usuário clica no botão **Obter Sugestões de Reunião**, o manipulador de eventos `myGetMeetingSuggestions` obtém uma matriz de sugestões de reunião da propriedade [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) do objeto `_MyEntities`, caso algum seja extraído.</span><span class="sxs-lookup"><span data-stu-id="17038-171">When the user clicks the **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) property of the `_MyEntities` object, if any was extracted.</span></span>


 > [!NOTE]
 > <span data-ttu-id="17038-172">Somente as mensagens, mas não os compromissos `MeetingSuggestion` , dão suporte ao tipo de entidade.</span><span class="sxs-lookup"><span data-stu-id="17038-172">Only messages but not appointments support the `MeetingSuggestion` entity type.</span></span>

<span data-ttu-id="17038-p119">Cada sugestão de reunião extraída é armazenada como um objeto [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) na matriz. `myGetMeetingSuggestions` obtém dados adicionais sobre cada sugestão de reunião:</span><span class="sxs-lookup"><span data-stu-id="17038-p119">Each extracted meeting suggestion is stored as a [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:</span></span>


- <span data-ttu-id="17038-175">A cadeia de caracteres que foi identificada como uma sugestão de reunião na propriedade [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring).</span><span class="sxs-lookup"><span data-stu-id="17038-175">The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) property.</span></span>

- <span data-ttu-id="17038-p120">A matriz de participantes da reunião na propriedade [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees). Cada participante é representado por um objeto [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="17038-p120">The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="17038-178">Para cada participante, o nome na propriedade [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname).</span><span class="sxs-lookup"><span data-stu-id="17038-178">For each attendee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="17038-179">Para cada participante, o endereço SMTP na propriedade [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress).</span><span class="sxs-lookup"><span data-stu-id="17038-179">For each attendee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

- <span data-ttu-id="17038-180">A cadeia de caracteres que representa a localização de sugestão de reunião na propriedade [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location).</span><span class="sxs-lookup"><span data-stu-id="17038-180">The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) property.</span></span>

- <span data-ttu-id="17038-181">A cadeia de caracteres que representa o assunto da sugestão de reunião na propriedade [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject).</span><span class="sxs-lookup"><span data-stu-id="17038-181">The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) property.</span></span>

- <span data-ttu-id="17038-182">A cadeia de caracteres que representa a hora de início da sugestão de reunião na propriedade [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start).</span><span class="sxs-lookup"><span data-stu-id="17038-182">The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) property.</span></span>

- <span data-ttu-id="17038-183">A cadeia de caracteres que representa a hora de término da sugestão de reunião na propriedade [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end).</span><span class="sxs-lookup"><span data-stu-id="17038-183">The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) property.</span></span>

<span data-ttu-id="17038-p121">`myGetMeetingSuggestions` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada uma das sugestões de reunião. A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p121">`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-phone-numbers"></a><span data-ttu-id="17038-186">Extrair números de telefone</span><span class="sxs-lookup"><span data-stu-id="17038-186">Extracting phone numbers</span></span>


<span data-ttu-id="17038-p122">Quando o usuário clica no botão **Obter Números de Telefone**, o manipulador de eventos `myGetPhoneNumbers` obtém uma matriz de números de telefone na propriedade [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) do objeto `_MyEntities`, caso algum seja extraído. Cada número de telefone extraído é armazenado como objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber) na matriz. `myGetPhoneNumbers` obtém mais dados sobre cada número de telefone:</span><span class="sxs-lookup"><span data-stu-id="17038-p122">When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:</span></span>


- <span data-ttu-id="17038-190">A cadeia de caracteres que representa o tipo de número de telefone, por exemplo, número de telefone residencial, na propriedade [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type).</span><span class="sxs-lookup"><span data-stu-id="17038-190">The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) property.</span></span>

- <span data-ttu-id="17038-191">A cadeia de caracteres que representa o número de telefone real na propriedade [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring).</span><span class="sxs-lookup"><span data-stu-id="17038-191">The string representing the actual phone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="17038-192">A cadeia de caracteres que foi originalmente identificada como o número de telefone na propriedade [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring).</span><span class="sxs-lookup"><span data-stu-id="17038-192">The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) property.</span></span>

<span data-ttu-id="17038-p123">`myGetPhoneNumbers` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada um dos números de telefone. A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p123">`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-task-suggestions"></a><span data-ttu-id="17038-195">Extrair sugestões de tarefa</span><span class="sxs-lookup"><span data-stu-id="17038-195">Extracting task suggestions</span></span>


<span data-ttu-id="17038-p124">Quando o usuário clica no botão **Obter Sugestões de Tarefa**, o manipulador de eventos `myGetTaskSuggestions` obtém uma matriz de sugestões de tarefa na propriedade [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) do objeto `_MyEntities`, caso algum seja extraído. Cada sugestão de tarefa extraída é armazenada como um objeto [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) da matriz. `myGetTaskSuggestions` obtém dados adicionais sobre cada sugestão de tarefa:</span><span class="sxs-lookup"><span data-stu-id="17038-p124">When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:</span></span>


- <span data-ttu-id="17038-199">A cadeia de caracteres que foi originalmente identificada como uma sugestão de tarefa na propriedade [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring).</span><span class="sxs-lookup"><span data-stu-id="17038-199">The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) property.</span></span>

- <span data-ttu-id="17038-p125">A matriz dos destinatários da tarefa na propriedade [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees). Cada destinatário é representado por um objeto [EmailUser](/javascript/api/outlook/office.emailuser).</span><span class="sxs-lookup"><span data-stu-id="17038-p125">The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="17038-202">Para cada destinatário, o nome na propriedade [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname).</span><span class="sxs-lookup"><span data-stu-id="17038-202">For each assignee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="17038-203">Para cada destinatário, o endereço SMTP da propriedade [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress).</span><span class="sxs-lookup"><span data-stu-id="17038-203">For each assignee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

<span data-ttu-id="17038-p126">`myGetTaskSuggestions` forma de uma cadeia de caracteres HTML local em `htmlText` para exibir os dados de cada sugestão de tarefa. A seguir apresentamos o código JavaScript relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p126">`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-urls"></a><span data-ttu-id="17038-206">Extrair URLs</span><span class="sxs-lookup"><span data-stu-id="17038-206">Extracting URLs</span></span>


<span data-ttu-id="17038-p127">Quando o usuário clica no botão **Obter URLs**, o manipulador de eventos `myGetUrls` obtém uma matriz de URLs na propriedade [urls](/javascript/api/outlook/office.entities#urls) do objeto `_MyEntities`, caso algum seja extraído. Cada URL extraída é armazenada como uma cadeia de caracteres na matriz. `myGetUrls` forma uma cadeia de caracteres HTML local em `htmlText` para exibir a lista de URLs extraídas.</span><span class="sxs-lookup"><span data-stu-id="17038-p127">When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#urls) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.</span></span>


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


## <a name="clearing-displayed-entity-strings"></a><span data-ttu-id="17038-210">Limpar cadeias de caracteres de entidade exibidas</span><span class="sxs-lookup"><span data-stu-id="17038-210">Clearing displayed entity strings</span></span>


<span data-ttu-id="17038-p128">Por fim, o suplemento de entidades especifica um manipulador de eventos `myClearEntitiesBox` que limpa as cadeias de caracteres exibidas. A seguir apresentamos o código relacionado.</span><span class="sxs-lookup"><span data-stu-id="17038-p128">Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.</span></span>


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a><span data-ttu-id="17038-213">Listagem de JavaScript</span><span class="sxs-lookup"><span data-stu-id="17038-213">JavaScript listing</span></span>


<span data-ttu-id="17038-214">A seguir apresentamos uma listagem completa da implementação do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="17038-214">The following is the complete listing of the JavaScript implementation.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="17038-215">Confira também</span><span class="sxs-lookup"><span data-stu-id="17038-215">See also</span></span>

- [<span data-ttu-id="17038-216">Criar suplementos do Outlook para formulários de leitura</span><span class="sxs-lookup"><span data-stu-id="17038-216">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="17038-217">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="17038-217">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="17038-218">Método item.getEntities</span><span class="sxs-lookup"><span data-stu-id="17038-218">item.getEntities method</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
