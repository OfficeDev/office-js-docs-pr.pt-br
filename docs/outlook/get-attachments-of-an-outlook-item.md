---
title: Obter anexos em um suplemento do Outlook
description: Seu suplemento pode usar a API de anexos para enviar informações sobre os anexos a um serviço remoto.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: bcb8226ab0755351b9e3a365e40623d258887d3f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612077"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a><span data-ttu-id="4e9c7-103">Obter anexos de um item do Outlook a partir do servidor</span><span class="sxs-lookup"><span data-stu-id="4e9c7-103">Get attachments of an Outlook item from the server</span></span>

<span data-ttu-id="4e9c7-p101">Um suplemento do Outlook não pode passar os anexos de um item selecionado diretamente para o serviço remoto que é executado em seu servidor. Em vez disso, o suplemento pode usar a API de anexos para enviar informações sobre os anexos ao serviço remoto. Em seguida, o serviço pode contatar o Exchange Server diretamente para recuperar os anexos.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-p101">An Outlook add-in cannot pass the attachments of a selected item directly to the remote service that runs on your server. Instead, the add-in can use the attachments API to send information about the attachments to the remote service. The service can then contact the Exchange server directly to retrieve the attachments.</span></span>

<span data-ttu-id="4e9c7-107">Para enviar informações de anexo para o serviço remoto, você pode usar as propriedades e a função abaixo:</span><span class="sxs-lookup"><span data-stu-id="4e9c7-107">To send attachment information to the remote service, you use the following properties and function:</span></span>

- <span data-ttu-id="4e9c7-p102">Propriedade [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities)&ndash;: fornece a URL dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio. Seu serviço usa essa URL para chamar o método [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) do EWS.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-p102">[Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) property &ndash; Provides the URL of Exchange Web Services (EWS) on the Exchange server that hosts the mailbox. Your service uses this URL to call the [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation.</span></span>

- <span data-ttu-id="4e9c7-110">Propriedade [Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) &ndash; obtém uma matriz de objetos [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails), uma para cada anexo do item.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-110">[Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property &ndash; Gets an array of [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) objects, one for each attachment to the item.</span></span>

- <span data-ttu-id="4e9c7-111">Função [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) &ndash; faz uma chamada assíncrona ao Exchange Server que hospeda a caixa de correio para obter um token de retorno de chamada que o servidor envia de volta ao Exchange Server para autenticar uma solicitação de um anexo.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-111">[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) function &ndash; Makes an asynchronous call to the Exchange server that hosts the mailbox to get a callback token that the server sends back to the Exchange server to authenticate a request for an attachment.</span></span>

## <a name="using-the-attachments-api"></a><span data-ttu-id="4e9c7-112">Usar a API de anexos</span><span class="sxs-lookup"><span data-stu-id="4e9c7-112">Using the attachments API</span></span>

<span data-ttu-id="4e9c7-113">Para usar a API de anexos e obter anexos de uma caixa de correio do Exchange, execute as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="4e9c7-113">To use the attachments API to get attachments from an Exchange mailbox, perform the following steps:</span></span>

1. <span data-ttu-id="4e9c7-114">Mostre o suplemento quando o usuário estiver exibindo uma mensagem ou um compromisso que contém um anexo.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-114">Show the add-in when the user is viewing a message or appointment that contains an attachment.</span></span>

1. <span data-ttu-id="4e9c7-115">Obtenha o token de retorno de chamada do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-115">Get the callback token from the Exchange server.</span></span>

1. <span data-ttu-id="4e9c7-116">Envie informações do token de retorno de chamada e do anexo para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-116">Send the callback token and attachment information to the remote service.</span></span>

1. <span data-ttu-id="4e9c7-117">Obtenha os anexos do Exchange Server usando o método `ExchangeService.GetAttachments` ou a operação `GetAttachment`.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-117">Get the attachments from the Exchange server by using the `ExchangeService.GetAttachments` method or the `GetAttachment` operation.</span></span>

<span data-ttu-id="4e9c7-118">Cada uma dessas etapas é abordada em detalhes nas seções a seguir usando o código do exemplo [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).</span><span class="sxs-lookup"><span data-stu-id="4e9c7-118">Each of these steps is covered in detail in the following sections using code from the [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) sample.</span></span>

> [!NOTE]
> <span data-ttu-id="4e9c7-119">O código nesses exemplos foi reduzido para enfatizar as informações do anexo.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-119">The code in these examples has been shortened to emphasize the attachment information.</span></span> <span data-ttu-id="4e9c7-120">O exemplo contém código adicional para autenticar o suplemento com o servidor remoto e gerenciar o estado da solicitação.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-120">The sample contains additional code for authenticating the add-in with the remote server and managing the state of the request.</span></span>

## <a name="get-a-callback-token"></a><span data-ttu-id="4e9c7-121">Obter um token de retorno de chamada</span><span class="sxs-lookup"><span data-stu-id="4e9c7-121">Get a callback token</span></span>

<span data-ttu-id="4e9c7-122">O objeto [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) fornece a função `getCallbackTokenAsync` para obter um token que o servidor remoto pode usar para se autenticar com o Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-122">The [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) object provides the `getCallbackTokenAsync` function to get a token that the remote server can use to authenticate with the Exchange server.</span></span> <span data-ttu-id="4e9c7-123">O código a seguir mostra uma função em um suplemento que inicia a solicitação assíncrona para obter o token de retorno de chamada e a função de retorno de chamada que obtém a resposta.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-123">The following code shows a function in an add-in that starts the asynchronous request to get the callback token, and the callback function that gets the response.</span></span> <span data-ttu-id="4e9c7-124">O token de retorno de chamada é armazenado no objeto de solicitação de serviço que é definido na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-124">The callback token is stored in the service request object that is defined in the next section.</span></span>

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}
```

## <a name="send-attachment-information-to-the-remote-service"></a><span data-ttu-id="4e9c7-125">Enviar informações de anexo ao serviço remoto</span><span class="sxs-lookup"><span data-stu-id="4e9c7-125">Send attachment information to the remote service</span></span>

<span data-ttu-id="4e9c7-p105">O serviço remoto que seu suplemento chama define as especificações de como você deve enviar as informações de anexo para o serviço. Neste exemplo, o serviço remoto é um aplicativo de API Web criado usando o Visual Studio 2013. O serviço remoto espera as informações de anexo em um objeto JSON. O código a seguir inicializa um objeto que contém as informações do anexo.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-p105">The remote service that your add-in calls defines the specifics of how you should send the attachment information to the service. In this example, the remote service is a Web API application created by using Visual Studio 2013. The remote service expects the attachment information in a JSON object. The following code initializes an object that contains the attachment information.</span></span>

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

<br/>

<span data-ttu-id="4e9c7-130">A propriedade `Office.context.mailbox.item.attachments` contém um conjunto de objetos `AttachmentDetails`, um para cada anexo do item.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-130">The `Office.context.mailbox.item.attachments` property contains a collection of `AttachmentDetails` objects, one for each attachment to the item.</span></span> <span data-ttu-id="4e9c7-131">Na maioria dos casos, o suplemento pode passar apenas a propriedade de ID de anexo de um objeto `AttachmentDetails` para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-131">In most cases, the add-in can pass just the attachment ID property of an `AttachmentDetails` object to the remote service.</span></span> <span data-ttu-id="4e9c7-132">Se o serviço remoto precisar de mais detalhes sobre o anexo, você poderá passar todo ou parte do objeto `AttachmentDetails`.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-132">If the remote service needs more details about the attachment, you can pass all or part of the `AttachmentDetails` object.</span></span> <span data-ttu-id="4e9c7-133">O código a seguir define um método que coloca toda a matriz `AttachmentDetails` no objeto `serviceRequest` e envia uma solicitação para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-133">The following code defines a method that puts the entire `AttachmentDetails` array in the `serviceRequest` object and sends a request to the remote service.</span></span>

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (var i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      var names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## <a name="get-the-attachments-from-the-exchange-server"></a><span data-ttu-id="4e9c7-134">Obter os anexos do Exchange Server</span><span class="sxs-lookup"><span data-stu-id="4e9c7-134">Get the attachments from the Exchange server</span></span>

<span data-ttu-id="4e9c7-p107">Seu serviço remoto pode usar o método de API gerenciada por EWS [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou a operação dos EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) para recuperar anexos do servidor. O aplicativo de serviço precisa de dois objetos para desserializar a cadeia de caracteres JSON em objetos .NET Framework que podem ser usados no servidor. O código a seguir mostra as definições dos objetos de desserialização.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-p107">Your remote service can use either the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API method or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation to retrieve attachments from the server. The service application needs two objects to deserialize the JSON string into .NET Framework objects that can be used on the server. The following code shows the definitions of the deserialization objects.</span></span>

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a><span data-ttu-id="4e9c7-138">Usar a API gerenciada por EWS para obter os anexos</span><span class="sxs-lookup"><span data-stu-id="4e9c7-138">Use the EWS Managed API to get the attachments</span></span>

<span data-ttu-id="4e9c7-p108">Se você usar a [API gerenciada por EWS](https://go.microsoft.com/fwlink/?LinkID=255472) no seu serviço remoto, poderá usar o método [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) que irá construir, enviar e receber uma solicitação SOAP dos EWS para obter os anexos. Recomendamos que você use a API gerenciada por EWS porque ela requer menos linhas de código e fornece uma interface mais intuitiva para fazer chamadas aos EWS. O código a seguir faz uma solicitação para recuperar todos os anexos e retorna a contagem e os nomes dos anexos processados.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-p108">If you use the [EWS Managed API](https://go.microsoft.com/fwlink/?LinkID=255472) in your remote service, you can use the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, which will construct, send, and receive an EWS SOAP request to get the attachments. We recommend that you use the EWS Managed API because it requires fewer lines of code and provides a more intuitive interface for making calls to EWS. The following code makes one request to retrieve all the attachments, and returns the count and names of the attachments processed.</span></span>

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### <a name="use-ews-to-get-the-attachments"></a><span data-ttu-id="4e9c7-142">Usar os EWS para obter os anexos</span><span class="sxs-lookup"><span data-stu-id="4e9c7-142">Use EWS to get the attachments</span></span>

<span data-ttu-id="4e9c7-143">Se você usar os EWS em seu serviço remoto, precisará criar uma solicitação SOAP [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) para obter os anexos do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-143">If you use EWS in your remote service, you need to construct a [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="4e9c7-144">O código a seguir retorna uma cadeia de caracteres que fornece a solicitação SOAP.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-144">The following code returns a string that provides the SOAP request.</span></span> <span data-ttu-id="4e9c7-145">O serviço remoto usa o método `String.Format` para inserir a ID de um anexo na cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-145">The remote service uses the `String.Format` method to insert the attachment ID for an attachment into the string.</span></span>


```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

<br/>

<span data-ttu-id="4e9c7-146">Por fim, o método a seguir usa uma solicitação dos EWS `GetAttachment` para obter os anexos do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-146">Finally, the following method does the work of using an EWS `GetAttachment` request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="4e9c7-147">Essa implementação faz uma solicitação individual para cada anexo e retorna a contagem dos anexos processados.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-147">This implementation makes an individual request for each attachment, and returns the count of attachments processed.</span></span> <span data-ttu-id="4e9c7-148">Cada resposta é processada em um método `ProcessXmlResponse` separado, definido a seguir.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-148">Each response is processed in a separate `ProcessXmlResponse` method, defined next.</span></span>

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

<br/>

<span data-ttu-id="4e9c7-149">Cada resposta da operação `GetAttachment` é enviada ao método `ProcessXmlResponse`.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-149">Each response from the `GetAttachment` operation is sent to the `ProcessXmlResponse` method.</span></span> <span data-ttu-id="4e9c7-150">Esse método verifica erros na resposta.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-150">This method checks the response for errors.</span></span> <span data-ttu-id="4e9c7-151">Se não encontrar erros, ele processará anexos de arquivo e de item.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-151">If it doesn't find any errors, it processes file attachments and item attachments.</span></span> <span data-ttu-id="4e9c7-152">O método `ProcessXmlResponse` executa a maior parte do trabalho para processar o anexo.</span><span class="sxs-lookup"><span data-stu-id="4e9c7-152">The `ProcessXmlResponse` method performs the bulk of the work to process the attachment.</span></span>

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

## <a name="see-also"></a><span data-ttu-id="4e9c7-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="4e9c7-153">See also</span></span>

- [<span data-ttu-id="4e9c7-154">Criar suplementos do Outlook para formulários de leitura</span><span class="sxs-lookup"><span data-stu-id="4e9c7-154">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="4e9c7-155">Explorar os recursos do EWS Managed API, do EWS e dos serviços Web no Exchange</span><span class="sxs-lookup"><span data-stu-id="4e9c7-155">Explore the EWS Managed API, EWS, and web services in Exchange</span></span>](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [<span data-ttu-id="4e9c7-156">Introdução aos aplicativos clientes de API gerenciada por EWS</span><span class="sxs-lookup"><span data-stu-id="4e9c7-156">Get started with EWS Managed API client applications</span></span>](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [<span data-ttu-id="4e9c7-157">Exemplo de AttachmentsDemo de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="4e9c7-157">AttachmentsDemo Sample Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-attachments-demo)
