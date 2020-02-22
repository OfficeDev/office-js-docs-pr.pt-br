---
title: Usar os Serviços Web do Exchange a partir de um suplemento do Outlook
description: Fornece um exemplo que mostra como um suplemento do Outlook pode solicitar informações dos Serviços Web do Exchange.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4c0c97a9a796dc1f257b1a0b0ec880b3ca3d8e74
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165686"
---
# <a name="call-web-services-from-an-outlook-add-in"></a><span data-ttu-id="8dd65-103">Chamar serviços Web de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="8dd65-103">Call web services from an Outlook add-in</span></span>

<span data-ttu-id="8dd65-p101">O suplemento pode usar os EWS (Serviços Web do Exchange) de um computador que esteja executando o Exchange Server 2013, um serviço Web que está disponível no servidor que fornece o local de origem para interface do usuário do suplemento ou um serviço Web que está disponível na Internet. Este artigo fornece um exemplo que mostra como um suplemento do Outlook pode solicitar informações dos EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p101">Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.</span></span>

<span data-ttu-id="8dd65-p102">A maneira usada para chamar um serviço Web varia com base em onde o serviço Web está localizado. A Tabela 1 lista as diferentes maneiras que podem ser usadas para chamar um serviço Web baseado no local.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p102">The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.</span></span>


<span data-ttu-id="8dd65-108">**Tabela 1. Maneiras de chamar serviços Web de um suplemento do Outlook**</span><span class="sxs-lookup"><span data-stu-id="8dd65-108">**Table 1. Ways to call web services from an Outlook add-in**</span></span>

<br/>

|<span data-ttu-id="8dd65-109">**Local do serviço Web**</span><span class="sxs-lookup"><span data-stu-id="8dd65-109">**Web service location**</span></span>|<span data-ttu-id="8dd65-110">**Maneira de chamar o serviço Web**</span><span class="sxs-lookup"><span data-stu-id="8dd65-110">**Way to call the web service**</span></span>|
|:-----|:-----|
|<span data-ttu-id="8dd65-111">O servidor Exchange que hospeda a caixa de correio do cliente</span><span class="sxs-lookup"><span data-stu-id="8dd65-111">The Exchange server that hosts the client mailbox</span></span>|<span data-ttu-id="8dd65-p103">Use o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para chamar operações EWS com suporte dos suplementos. O servidor Exchange que hospeda a caixa de correio também expõe os EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p103">Use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.</span></span>|
|<span data-ttu-id="8dd65-114">O servidor Web que fornece o local de origem para a interface do usuário</span><span class="sxs-lookup"><span data-stu-id="8dd65-114">The web server that provides the source location for the add-in UI</span></span>|<span data-ttu-id="8dd65-p104">Chame o serviço Web usando técnicas JavaScript padrão. O código JavaScript no quadro da interface do usuário é executado no contexto do servidor Web que fornece a interface do usuário. Portanto, ele pode chamar serviços Web nesse servidor sem causar um erro de script entre sites.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p104">Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.</span></span>|
|<span data-ttu-id="8dd65-118">Todos os outros locais</span><span class="sxs-lookup"><span data-stu-id="8dd65-118">All other locations</span></span>|<span data-ttu-id="8dd65-p105">Crie um proxy para o serviço Web no servidor Web que fornece o local de origem para a interface do usuário. Se você não fornecer um proxy, erros de script entre sites impedirão a execução do suplemento. Uma maneira de fornecer um proxy é usar JSON/P. Para saber mais, confira [Privacidade e segurança para suplementos do Office](../develop/privacy-and-security.md).</span><span class="sxs-lookup"><span data-stu-id="8dd65-p105">Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../develop/privacy-and-security.md).</span></span>|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a><span data-ttu-id="8dd65-123">Usar o método makeEwsRequestAsync para acessar operações dos EWS</span><span class="sxs-lookup"><span data-stu-id="8dd65-123">Using the makeEwsRequestAsync method to access EWS operations</span></span>

<span data-ttu-id="8dd65-124">Você pode usar o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para fazer uma solicitação dos EWS ao servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8dd65-124">You can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to make an EWS request to the Exchange server that hosts the user's mailbox.</span></span>

<span data-ttu-id="8dd65-p106">Os EWS oferecem suporte a diferentes operações em um servidor Exchange, por exemplo, operações no nível do item para copiar, localizar, atualizar ou enviar um item e operações no nível da pasta para criar, acessar ou atualizar uma pasta. Para executar uma operação dos EWS, crie uma solicitação SOAP XML para a operação. Quando a operação termina, você recebe uma resposta SOAP XML que contém dados que são relevantes para a operação. As solicitações e respostas SOAP dos EWS seguem o esquema definido no arquivo Messages.xsd. Como outros arquivos de esquema dos EWS, o arquivo Message.xsd está localizado no diretório virtual do IIS que hospeda os EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p106">EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.</span></span>

<span data-ttu-id="8dd65-130">Para usar o método **makeEwsRequestAsync** para iniciar uma operação dos EWS, forneça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8dd65-130">To use the **makeEwsRequestAsync** method to initiate an EWS operation, provide the following:</span></span>

- <span data-ttu-id="8dd65-131">O XML para a solicitação SOAP dessa operação dos EWS, como um argumento para o parâmetro _data_</span><span class="sxs-lookup"><span data-stu-id="8dd65-131">The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter</span></span>

- <span data-ttu-id="8dd65-132">Um método de retorno (como o argumento _callback_)</span><span class="sxs-lookup"><span data-stu-id="8dd65-132">A callback method (as the  _callback_ argument)</span></span>

- <span data-ttu-id="8dd65-133">Outros dados de entrada opcionais para esse método de retorno de chamada (como o argumento _userContext_)</span><span class="sxs-lookup"><span data-stu-id="8dd65-133">Any optional input data for that callback method (as the  _userContext_ argument)</span></span>

<span data-ttu-id="8dd65-p107">Quando a solicitação SOAP dos EWS é concluída, o Outlook chama o método de retorno de chamada com um argumento, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult). O método de retorno de chamada pode acessar duas propriedades do objeto **AsyncResult**: a propriedade **value**, que contém a resposta SOAP XML da operação dos EWS e, opcionalmente, a propriedade **asyncContext**, que contém todos os dados passados como o parâmetro **userContext**. Normalmente, o método de retorno de chamada analisa o XML na resposta SOAP para obter todas as informações relevantes e processa essas informações da maneira adequada.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p107">When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object. The callback method can access two properties of the **AsyncResult** object: the **value** property, which contains the XML SOAP response of the EWS operation, and optionally, the **asyncContext** property, which contains any data passed as the **userContext** parameter. Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.</span></span>


## <a name="tips-for-parsing-ews-responses"></a><span data-ttu-id="8dd65-137">Dicas para analisar respostas dos EWS</span><span class="sxs-lookup"><span data-stu-id="8dd65-137">Tips for parsing EWS responses</span></span>

<span data-ttu-id="8dd65-138">Ao analisar uma resposta SOAP de uma operação dos EWS, observe os seguintes problemas que dependem do navegador:</span><span class="sxs-lookup"><span data-stu-id="8dd65-138">When parsing a SOAP response from an EWS operation, note the following browser-dependent issues:</span></span>


- <span data-ttu-id="8dd65-139">Especifique o prefixo para um nome de marca ao usar o método DOM **getElementsByTagName**, para incluir suporte para o Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="8dd65-139">Specify the prefix for a tag name when using the DOM method **getElementsByTagName**, to include support for Internet Explorer.</span></span>

  <span data-ttu-id="8dd65-p108">**getElementsByTagName** se comporta de forma diferente dependendo do tipo de navegador. Por exemplo, uma resposta EWS pode conter o seguinte XML (formatado e abreviado para fins de exibição):</span><span class="sxs-lookup"><span data-stu-id="8dd65-p108">**getElementsByTagName** behaves differently depending on browser type. For example, an EWS response can contain the following XML (formatted and abbreviated for display purposes):</span></span>

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   <span data-ttu-id="8dd65-142">Um código como o que é mostrado abaixo funcionaria em um navegador como o Chrome para obter o XML delimitado pelas marcas **ExtendedProperty**:</span><span class="sxs-lookup"><span data-stu-id="8dd65-142">Code, as in the following, would work on a browser like Chrome to get the XML enclosed by the **ExtendedProperty** tags:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   <span data-ttu-id="8dd65-143">No Internet Explorer, você precisa incluir o prefixo `t:` do nome da marca, conforme mostrado abaixo:</span><span class="sxs-lookup"><span data-stu-id="8dd65-143">On Internet Explorer, you must include the `t:` prefix of the tag name, as shown below:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- <span data-ttu-id="8dd65-144">Use a propriedade DOM **textContent** para acessar o conteúdo de uma marca em uma resposta dos EWS, conforme mostrado abaixo:</span><span class="sxs-lookup"><span data-stu-id="8dd65-144">Use the DOM property **textContent** to get the contents of a tag in an EWS response, as shown below:</span></span>
    
   ```js
      content = $.parseJSON(value.textContent);
   ```

   <span data-ttu-id="8dd65-145">Outras propriedades, como **innerHTML** podem não funcionar no Internet Explorer para algumas marcas em uma resposta dos EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-145">Other properties such as **innerHTML** may not work on Internet Explorer for some tags in an EWS response.</span></span>
    

## <a name="example"></a><span data-ttu-id="8dd65-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8dd65-146">Example</span></span>

<span data-ttu-id="8dd65-p109">O exemplo a seguir chama **makeEwsRequestAsync** para usar a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) a fim de obter o assunto de um item. Esse exemplo inclui as três funções a seguir:</span><span class="sxs-lookup"><span data-stu-id="8dd65-p109">The following example calls **makeEwsRequestAsync** to use the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to get the subject of an item. This example includes the following three functions:</span></span>

-  <span data-ttu-id="8dd65-149">`getSubjectRequest` &ndash; considera uma ID do item como entrada e retorna o XML a fim de que a solicitação SOAP possa chamar **GetItem** para o item especificado.</span><span class="sxs-lookup"><span data-stu-id="8dd65-149">`getSubjectRequest` &ndash; Takes an item ID as input, and returns the XML for the SOAP request to call **GetItem** for the specified item.</span></span>
    
-  <span data-ttu-id="8dd65-150">`sendRequest` &ndash; chama `getSubjectRequest` para obter a solicitação SOAP para o item selecionado, passa a solicitação SOAP e o método de retorno de chamada `callback` para **makeEwsRequestAsync** a fim de obter o assunto do item especificado.</span><span class="sxs-lookup"><span data-stu-id="8dd65-150">`sendRequest` &ndash; Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to **makeEwsRequestAsync** to get the subject of the specified item.</span></span>
    
-  <span data-ttu-id="8dd65-151">`callback` &ndash; processa a resposta SOAP que inclui o assunto e outras informações sobre o item especificado.</span><span class="sxs-lookup"><span data-stu-id="8dd65-151">`callback` &ndash; Processes the SOAP response which includes any subject and other information about the specified item.</span></span>
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}
```


## <a name="ews-operations-that-add-ins-support"></a><span data-ttu-id="8dd65-152">Operações dos EWS compatíveis com suplementos</span><span class="sxs-lookup"><span data-stu-id="8dd65-152">EWS operations that add-ins support</span></span>

<span data-ttu-id="8dd65-p110">Os suplementos do Outlook podem acessar um subconjunto de operações que estão disponíveis nos EWS pelo método **makeEwsRequestAsync**. Se você não estiver familiarizado com operações dos EWS e com o uso do método **makeEwsRequestAsync** para acessar uma operação, comece com um exemplo de solicitação SOAP para personalizar seu argumento _data_.</span><span class="sxs-lookup"><span data-stu-id="8dd65-p110">Outlook add-ins can access a subset of operations that are available in EWS via the **makeEwsRequestAsync** method. If you are unfamiliar with EWS operations and how to use the **makeEwsRequestAsync** method to access an operation, start with a SOAP request example to customize your _data_ argument.</span></span> 

<span data-ttu-id="8dd65-155">Este procedimento descreve como é possível usar o método **makeEwsRequestAsync**:</span><span class="sxs-lookup"><span data-stu-id="8dd65-155">The following describes how you can use the **makeEwsRequestAsync** method:</span></span>

1. <span data-ttu-id="8dd65-156">No XML, substitua as IDs de item e atributos relevantes da operação dos EWS por valores apropriados.</span><span class="sxs-lookup"><span data-stu-id="8dd65-156">In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.</span></span>
    
2. <span data-ttu-id="8dd65-157">Inclua a solicitação SOAP como um argumento para o parâmetro _data_ de **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="8dd65-157">Include the SOAP request as an argument for the  _data_ parameter of **makeEwsRequestAsync**.</span></span>
    
3. <span data-ttu-id="8dd65-158">Especifique um método de retorno de chamada e chame **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="8dd65-158">Specify a callback method and call **makeEwsRequestAsync**.</span></span>
    
4. <span data-ttu-id="8dd65-159">No método de retorno de chamada, verifique os resultados da operação na resposta SOAP.</span><span class="sxs-lookup"><span data-stu-id="8dd65-159">In the callback method, verify the results of the operation in the SOAP response.</span></span>
    
5. <span data-ttu-id="8dd65-160">Use os resultados da operação dos EWS de acordo com as suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="8dd65-160">Use the results of the EWS operation according to your needs.</span></span>
    
<span data-ttu-id="8dd65-p111">A tabela a seguir lista as operações dos EWS compatíveis com suplementos. Para ver exemplos de solicitações e respostas SOAP, escolha o link para cada operação. Para saber mais sobre operações dos EWS, confira [Operações dos EWS no Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="8dd65-p111">The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span></span>

<span data-ttu-id="8dd65-164">**Tabela 2. Operações compatíveis do EWS**</span><span class="sxs-lookup"><span data-stu-id="8dd65-164">**Table 2. Supported EWS operations**</span></span>

<br/>

|<span data-ttu-id="8dd65-165">**Operação do EWS**</span><span class="sxs-lookup"><span data-stu-id="8dd65-165">**EWS operation**</span></span>|<span data-ttu-id="8dd65-166">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="8dd65-166">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="8dd65-167">Operação CopyItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-167">CopyItem operation</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)|<span data-ttu-id="8dd65-168">Copia os itens especificados e coloca os novos itens em uma pasta designada no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-168">Copies the specified items and puts the new items in a designated folder in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-169">Operação CreateFolder</span><span class="sxs-lookup"><span data-stu-id="8dd65-169">CreateFolder operation</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)|<span data-ttu-id="8dd65-170">Cria pastas no local especificado no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-170">Creates folders in the specified location in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-171">Operação CreateItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-171">CreateItem operation</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)|<span data-ttu-id="8dd65-172">Cria os itens especificados no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-172">Creates the specified items in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-173">Operação ExpandDL</span><span class="sxs-lookup"><span data-stu-id="8dd65-173">ExpandDL operation</span></span>](/exchange/client-developer/web-service-reference/expanddl-operation)|<span data-ttu-id="8dd65-174">Exibe a associação completa das listas de distribuição.</span><span class="sxs-lookup"><span data-stu-id="8dd65-174">Displays the full membership of distribution lists.</span></span>|
|[<span data-ttu-id="8dd65-175">Operação FindConversation</span><span class="sxs-lookup"><span data-stu-id="8dd65-175">FindConversation operation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)|<span data-ttu-id="8dd65-176">Enumera uma lista de conversas na pasta especificada no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-176">Enumerates a list of conversations in the specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-177">Operação FindFolder</span><span class="sxs-lookup"><span data-stu-id="8dd65-177">FindFolder operation</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)|<span data-ttu-id="8dd65-178">Localiza subpastas de uma pasta identificada e retorna um conjunto de propriedades que descreve o conjunto de subpastas.</span><span class="sxs-lookup"><span data-stu-id="8dd65-178">Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.</span></span>|
|[<span data-ttu-id="8dd65-179">Operação FindItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-179">FindItem operation</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)|<span data-ttu-id="8dd65-180">Identifica os itens que estão localizados em uma pasta especificada no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-180">Identifies items that are located in a specified folder in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-181">Operação GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="8dd65-181">GetConversationItems operation</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)|<span data-ttu-id="8dd65-182">Obtém um ou mais conjuntos de itens que estão organizados em nós em uma conversa.</span><span class="sxs-lookup"><span data-stu-id="8dd65-182">Gets one or more sets of items that are organized in nodes in a conversation.</span></span>|
|[<span data-ttu-id="8dd65-183">Operação GetFolder</span><span class="sxs-lookup"><span data-stu-id="8dd65-183">GetFolder operation</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)|<span data-ttu-id="8dd65-184">Obtém as propriedades especificadas e o conteúdo de pastas do repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-184">Gets the specified properties and contents of folders from the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-185">Operação GetItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-185">GetItem operation</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)|<span data-ttu-id="8dd65-186">Obtém as propriedades especificadas e o conteúdo de itens do repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-186">Gets the specified properties and contents of items from the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-187">Operação GetUserAvailability</span><span class="sxs-lookup"><span data-stu-id="8dd65-187">GetUserAvailability operation</span></span>](/exchange/client-developer/web-service-reference/getuseravailability-operation)|<span data-ttu-id="8dd65-188">Fornece informações detalhadas sobre a disponibilidade de um conjunto de usuários, salas e recursos em um período especificado.</span><span class="sxs-lookup"><span data-stu-id="8dd65-188">Provides detailed information about the availability of a set of users, rooms, and resources within a specified time period.</span></span>|
|[<span data-ttu-id="8dd65-189">Operação MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="8dd65-189">MarkAsJunk operation</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)|<span data-ttu-id="8dd65-190">Move mensagens de email para a pasta Lixo Eletrônico e adiciona ou remove, adequadamente, remetentes das mensagens na lista de remetentes bloqueados.</span><span class="sxs-lookup"><span data-stu-id="8dd65-190">Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.</span></span>|
|[<span data-ttu-id="8dd65-191">Operação MoveItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-191">MoveItem operation</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)|<span data-ttu-id="8dd65-192">Move itens para uma única pasta de destino no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-192">Moves items to a single destination folder in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-193">Operação ResolveNames</span><span class="sxs-lookup"><span data-stu-id="8dd65-193">ResolveNames operation</span></span>](/exchange/client-developer/web-service-reference/resolvenames-operation)|<span data-ttu-id="8dd65-194">Resolve endereços de email e nomes de exibição ambíguos.</span><span class="sxs-lookup"><span data-stu-id="8dd65-194">Resolves ambiguous email addresses and display names.</span></span>|
|[<span data-ttu-id="8dd65-195">Operação SendItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-195">SendItem operation</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)|<span data-ttu-id="8dd65-196">Envia mensagens de email que estão localizadas no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-196">Sends email messages that are located in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-197">Operação UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="8dd65-197">UpdateFolder operation</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)|<span data-ttu-id="8dd65-198">Modifica as propriedades de pastas existentes no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-198">Modifies the properties of existing folders in the Exchange store.</span></span>|
|[<span data-ttu-id="8dd65-199">Operação UpdateItem</span><span class="sxs-lookup"><span data-stu-id="8dd65-199">UpdateItem operation</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)|<span data-ttu-id="8dd65-200">Modifica as propriedades de itens existentes no repositório do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8dd65-200">Modifies the properties of existing items in the Exchange store.</span></span>|

 > [!NOTE]
 > <span data-ttu-id="8dd65-201">Não é possível atualizar (ou criar) itens FAI (Informações Associadas da Pasta) usando um suplemento.</span><span class="sxs-lookup"><span data-stu-id="8dd65-201">FAI (Folder Associated Information) items cannot be updated (or created) from an add-in.</span></span> <span data-ttu-id="8dd65-202">Essas mensagens ocultas são armazenadas em uma pasta e usadas para armazenar diversas configurações e dados auxiliares.</span><span class="sxs-lookup"><span data-stu-id="8dd65-202">These hidden messages are stored in a folder and are used to store a variety of settings and auxiliary data.</span></span>  <span data-ttu-id="8dd65-203">Tentar usar a operação UpdateItem gera um erro ErrorAccessDenied: "A extensão do Office não tem permissão para atualizar esse item".</span><span class="sxs-lookup"><span data-stu-id="8dd65-203">Attempting to use the UpdateItem operation will throw an ErrorAccessDenied error: "Office extension is not allowed to update this type of item".</span></span> <span data-ttu-id="8dd65-204">Se preferir, use a [API Gerenciada do EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) para atualizar esses itens usando um cliente do Windows ou um aplicativo para servidores.</span><span class="sxs-lookup"><span data-stu-id="8dd65-204">As an alternative, you may use the [EWS Managed API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) to update these items from a Windows client or a server application.</span></span> <span data-ttu-id="8dd65-205">Recomenda-se cuidado já que as estruturas de dados internos de tipo de serviço estão sujeitas a alterações e podem invalidar sua solução.</span><span class="sxs-lookup"><span data-stu-id="8dd65-205">Caution is recommended as internal, service-type data structures are subject to change and could break your solution.</span></span>


## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a><span data-ttu-id="8dd65-206">Considerações sobre autenticação e permissão para makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="8dd65-206">Authentication and permission considerations for makeEwsRequestAsync</span></span>

<span data-ttu-id="8dd65-207">Ao usar o método **makeEwsRequestAsync**, a solicitação é autenticada usando as credenciais de conta de email do usuário atual.</span><span class="sxs-lookup"><span data-stu-id="8dd65-207">When you use the **makeEwsRequestAsync** method, the request is authenticated by using the email account credentials of the current user.</span></span> <span data-ttu-id="8dd65-208">O método **makeEwsRequestAsync** gerencia as credenciais para você e, portanto, não é preciso fornecer credenciais de autenticação com a sua solicitação.</span><span class="sxs-lookup"><span data-stu-id="8dd65-208">The **makeEwsRequestAsync** method manages the credentials for you so that you do not have to provide authentication credentials with your request.</span></span>

> [!NOTE]
> <span data-ttu-id="8dd65-209">O administrador do servidor deve usar o cmldet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) para definir o parâmetro _OAuthAuthentication_ como **verdadeiro** no diretório EWS do servidor de Acesso do Cliente e, assim, habilitar o método **makeEwsRequestAsync** para fazer solicitações EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-209">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) cmldet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the **makeEwsRequestAsync** method to make EWS requests.</span></span>

<span data-ttu-id="8dd65-210">Seu suplemento deve especificar a permissão **ReadWriteMailbox** no manifesto do suplemento para usar o método **makeEwsRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="8dd65-210">Your add-in must specify the **ReadWriteMailbox** permission in its add-in manifest to use the **makeEwsRequestAsync** method.</span></span> <span data-ttu-id="8dd65-211">Para saber mais sobre como usar a permissão **ReadWriteMailbox**, confira a seção [Permissão ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) em [Noções básicas sobre permissões de suplementos do Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="8dd65-211">For information about using the **ReadWriteMailbox** permission, see the section [ReadWriteMailbox permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) in [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

> [!NOTE]
> <span data-ttu-id="8dd65-212">O administrador do servidor deve usar o cmldet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) para definir o parâmetro _OAuthAuthentication_ como **verdadeiro** no diretório EWS do servidor de Acesso do Cliente e, assim, habilitar o método **makeEwsRequestAsync** para fazer solicitações EWS.</span><span class="sxs-lookup"><span data-stu-id="8dd65-212">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) cmldet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the **makeEwsRequestAsync** method to make EWS requests.</span></span>



## <a name="see-also"></a><span data-ttu-id="8dd65-213">Confira também</span><span class="sxs-lookup"><span data-stu-id="8dd65-213">See also</span></span>

- [<span data-ttu-id="8dd65-214">Privacidade e segurança para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8dd65-214">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)   
- [<span data-ttu-id="8dd65-215">Como lidar com limitações de política de mesma origem nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8dd65-215">Addressing same-origin policy limitations in Office Add-ins</span></span>](../develop/addressing-same-origin-policy-limitations.md)
- [<span data-ttu-id="8dd65-216">Referência do EWS para Exchange</span><span class="sxs-lookup"><span data-stu-id="8dd65-216">EWS reference for Exchange</span></span>](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)   
- [<span data-ttu-id="8dd65-217">Aplicativos de email para Outlook e EWS no Exchange</span><span class="sxs-lookup"><span data-stu-id="8dd65-217">Mail apps for Outlook and EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)
   
<span data-ttu-id="8dd65-218">Veja os tópicos a seguir para criar serviços de back-end para suplementos usando a API Web ASP.NET:</span><span class="sxs-lookup"><span data-stu-id="8dd65-218">See the following for creating backend services for add-ins using ASP.NET Web API:</span></span>

- [<span data-ttu-id="8dd65-219">Criar um serviço Web para um suplemento do Office usando a API Web ASP.NET</span><span class="sxs-lookup"><span data-stu-id="8dd65-219">Create a web service for an Office Add-in using the ASP.NET Web API</span></span>](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)    
- [<span data-ttu-id="8dd65-220">Noções básicas sobre a criação de um serviço HTTP usando a API Web ASP.NET</span><span class="sxs-lookup"><span data-stu-id="8dd65-220">The basics of building an HTTP service using ASP.NET Web API</span></span>](https://www.asp.net/web-api)
    