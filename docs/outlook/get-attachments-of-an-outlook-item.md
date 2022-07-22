---
title: Obter anexos em um suplemento do Outlook
description: Seu suplemento pode usar a API de anexos para enviar informações sobre os anexos a um serviço remoto.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 637513a5ee94f4a3b9fa6b913f4c419dd5ec4d8e
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958822"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>Obter anexos de um item do Outlook a partir do servidor

Você pode obter os anexos de um item do Outlook de algumas maneiras, mas qual opção você usa depende do seu cenário.

1. Envie as informações de anexo para o serviço remoto.

    Seu suplemento pode usar a API de anexos para enviar informações sobre os anexos para o serviço remoto. Em seguida, o serviço pode contatar o Exchange Server diretamente para recuperar os anexos.

1. Use a API [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) , disponível no conjunto de requisitos 1.8. Formatos com suporte: [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat).

    Essa API pode ser útil se o EWS/REST não estiver disponível (por exemplo, devido à configuração de administrador do servidor Exchange) ou se o suplemento quiser usar o conteúdo base64 diretamente em HTML ou JavaScript. Além disso, `getAttachmentContentAsync` a API está disponível em cenários de composição em que o anexo pode ainda não ter sido sincronizado com o Exchange; consulte Gerenciar [anexos](add-and-remove-attachments-to-an-item-in-a-compose-form.md) de um item em um formulário de composição no Outlook para saber mais.

Este artigo descreve a primeira opção. Para enviar informações de anexo para o serviço remoto, use as propriedades e o método a seguir.

- Propriedade [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities)&ndash;: fornece a URL dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio. Seu serviço usa essa URL para chamar o método [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) do EWS.

- Propriedade [Office.context.mailbox.item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) &ndash; obtém uma matriz de objetos [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails), uma para cada anexo do item.

- [O método Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) &ndash; faz uma chamada assíncrona para o servidor Exchange que hospeda a caixa de correio para obter um token de retorno de chamada que o servidor envia de volta ao servidor Exchange para autenticar uma solicitação de um anexo.

## <a name="using-the-attachments-api"></a>Usar a API de anexos

Para usar a API de anexos para obter anexos de uma caixa de correio do Exchange, execute as etapas a seguir.

1. Mostre o suplemento quando o usuário estiver exibindo uma mensagem ou um compromisso que contém um anexo.

1. Obtenha o token de retorno de chamada do Exchange Server.

1. Envie informações do token de retorno de chamada e do anexo para o serviço remoto.

1. Obtenha os anexos do Exchange Server usando o método `ExchangeService.GetAttachments` ou a operação `GetAttachment`.

Cada uma dessas etapas é abordada em detalhes nas seções a seguir usando o código do exemplo [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).

> [!NOTE]
> O código nesses exemplos foi reduzido para enfatizar as informações do anexo. O exemplo contém código adicional para autenticar o suplemento com o servidor remoto e gerenciar o estado da solicitação.

## <a name="get-a-callback-token"></a>Obter um token de retorno de chamada

O [objeto Office.context.mailbox](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox) `getCallbackTokenAsync` fornece o método para obter um token que o servidor remoto pode usar para autenticar com o servidor Exchange. O código a seguir mostra uma função em um suplemento que inicia a solicitação assíncrona para obter o token de retorno de chamada e a função de retorno de chamada que obtém a resposta. O token de retorno de chamada é armazenado no objeto de solicitação de serviço que é definido na próxima seção.

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

## <a name="send-attachment-information-to-the-remote-service"></a>Enviar informações de anexo ao serviço remoto

O serviço remoto que seu suplemento chama define as especificações de como você deve enviar as informações de anexo para o serviço. Neste exemplo, o serviço remoto é um aplicativo de API Web criado usando o Visual Studio 2013. O serviço remoto espera as informações de anexo em um objeto JSON. O código a seguir inicializa um objeto que contém as informações do anexo.

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 const serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

A propriedade `Office.context.mailbox.item.attachments` contém um conjunto de objetos `AttachmentDetails`, um para cada anexo do item. Na maioria dos casos, o suplemento pode passar apenas a propriedade de ID de anexo de um objeto `AttachmentDetails` para o serviço remoto. Se o serviço remoto precisar de mais detalhes sobre o anexo, você poderá passar todo ou parte do objeto `AttachmentDetails`. O código a seguir define um método que coloca toda a matriz `AttachmentDetails` no objeto `serviceRequest` e envia uma solicitação para o serviço remoto.

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (let i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      const names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (let i = 0; i < response.attachmentNames.length; i++) {
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

## <a name="get-the-attachments-from-the-exchange-server"></a>Obter os anexos do Exchange Server

Seu serviço remoto pode usar o método de API gerenciada por EWS [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) ou a operação dos EWS [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) para recuperar anexos do servidor. O aplicativo de serviço precisa de dois objetos para desserializar a cadeia de caracteres JSON em objetos .NET Framework que podem ser usados no servidor. O código a seguir mostra as definições dos objetos de desserialização.

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

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>Usar a API gerenciada por EWS para obter os anexos

Se você usar a [API gerenciada por EWS](/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange) no seu serviço remoto, poderá usar o método [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) que irá construir, enviar e receber uma solicitação SOAP dos EWS para obter os anexos. Recomendamos que você use a API gerenciada por EWS porque ela requer menos linhas de código e fornece uma interface mais intuitiva para fazer chamadas aos EWS. O código a seguir faz uma solicitação para recuperar todos os anexos e retorna a contagem e os nomes dos anexos processados.

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

### <a name="use-ews-to-get-the-attachments"></a>Usar os EWS para obter os anexos

Se você usar os EWS em seu serviço remoto, precisará criar uma solicitação SOAP [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) para obter os anexos do Exchange Server. O código a seguir retorna uma cadeia de caracteres que fornece a solicitação SOAP. O serviço remoto usa o método `String.Format` para inserir a ID de um anexo na cadeia de caracteres.

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

Por fim, o método a seguir usa uma solicitação dos EWS `GetAttachment` para obter os anexos do Exchange Server. Essa implementação faz uma solicitação individual para cada anexo e retorna a contagem dos anexos processados. Cada resposta é processada em um método `ProcessXmlResponse` separado, definido a seguir.

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

Cada resposta da operação `GetAttachment` é enviada ao método `ProcessXmlResponse`. Esse método verifica erros na resposta. Se não encontrar erros, ele processará anexos de arquivo e de item. O método `ProcessXmlResponse` executa a maior parte do trabalho para processar o anexo.

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

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de leitura](read-scenario.md)
- [Explorar os recursos do EWS Managed API, do EWS e dos serviços Web no Exchange](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [Introdução aos aplicativos clientes de API gerenciada por EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [SSO do Suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
