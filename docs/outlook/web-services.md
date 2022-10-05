---
title: Usar os Serviços Web do Exchange a partir de um suplemento do Outlook
description: Fornece um exemplo que mostra como um suplemento do Outlook pode solicitar informações dos Serviços Web do Exchange.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 94fff26fc7f9c16e2e385d6c44c128e4b03f968e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467010"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Chamar serviços Web de um suplemento do Outlook

Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.

The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.

**Tabela 1. Maneiras de chamar serviços Web de um suplemento do Outlook**

|**Local do serviço Web**|**Maneira de chamar o serviço Web**|
|:-----|:-----|
|O servidor Exchange que hospeda a caixa de correio do cliente|Use the [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.|
|O servidor Web que fornece o local de origem para a interface do usuário|Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.|
|Todos os outros locais|Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Usar o método makeEwsRequestAsync para acessar operações dos EWS

Você pode usar o método [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para fazer uma solicitação dos EWS ao servidor Exchange que hospeda a caixa de correio do usuário.

EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.

Para usar o `makeEwsRequestAsync` método para iniciar uma operação EWS, forneça o seguinte:

- O XML para a solicitação SOAP dessa operação dos EWS, como um argumento para o parâmetro _data_

- Uma função de retorno de chamada (como o  _argumento de retorno de_ chamada)

- Quaisquer dados de entrada opcionais para essa função de retorno de chamada (como o  _argumento userContext_ )

Quando a solicitação SOAP do EWS for concluída, o Outlook chamará a função de retorno de chamada com um argumento, que é um [objeto AsyncResult](/javascript/api/office/office.asyncresult) . `AsyncResult` A função de retorno de chamada pode acessar duas propriedades do objeto: `value` a propriedade, que contém a resposta XML SOAP da operação EWS e, opcionalmente, `asyncContext` a propriedade, `userContext` que contém todos os dados passados como o parâmetro. Normalmente, a função de retorno de chamada analisa o XML na resposta SOAP para obter informações relevantes e processa essas informações adequadamente.

## <a name="tips-for-parsing-ews-responses"></a>Dicas para analisar respostas dos EWS

Ao analisar uma resposta SOAP de uma operação EWS, observe os seguintes problemas dependentes do navegador.

- Especifique o prefixo de um nome de marca ao usar o método DOM `getElementsByTagName`, para incluir suporte para o Internet Explorer.

  `getElementsByTagName` se comporta de maneira diferente, dependendo do tipo de navegador. Por exemplo, uma resposta EWS pode conter o XML a seguir (formatado e abreviado para fins de exibição).

   ```XML
   <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
   PropertyName="MyProperty" 
   PropertyType="String"/>
   <t:Value>{
   ...
   }</t:Value></t:ExtendedProperty>
   ```

   O código, como mostrado a seguir, funcionaria em um navegador como o Chrome para colocar o XML entre as `ExtendedProperty` marcas.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("ExtendedProperty")
   });
   ```

   No Internet Explorer, você deve incluir o `t:` prefixo do nome da marca, da seguinte maneira.

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("t:ExtendedProperty")
   });
   ```

- Use a propriedade DOM `textContent` para obter o conteúdo de uma marca em uma resposta EWS, da seguinte maneira.

   ```js
   content = $.parseJSON(value.textContent);
   ```

   Outras propriedades, como podem `innerHTML` não funcionar no Internet Explorer para algumas marcas em uma resposta EWS.

## <a name="example"></a>Exemplo

O exemplo a seguir chama `makeEwsRequestAsync` para usar a [operação GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para obter o assunto de um item. Este exemplo inclui as três funções a seguir.

- `getSubjectRequest`&ndash; Usa uma ID de item como entrada e retorna o XML para a solicitação SOAP chamar `GetItem` para o item especificado.

- `sendRequest`&ndash; Chamadas `getSubjectRequest` para obter a solicitação SOAP para o item selecionado e, em seguida, passa a solicitação SOAP e a função de retorno de chamada, `callback``makeEwsRequestAsync` para obter o assunto do item especificado.

- `callback` &ndash; processa a resposta SOAP que inclui o assunto e outras informações sobre o item especificado.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   const result = 
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
   const mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   const result = asyncResult.value;
   const context = asyncResult.context;

   // Process the returned response here.
}
```

## <a name="ews-operations-that-add-ins-support"></a>Operações dos EWS compatíveis com suplementos

Os suplementos do Outlook podem acessar um subconjunto de operações que estão disponíveis no EWS por meio do `makeEwsRequestAsync` método. Se você não estiver familiarizado com as operações do EWS `makeEwsRequestAsync` e como usar o método para acessar uma operação, comece com um exemplo de solicitação SOAP para personalizar o _argumento de_ dados.

O exemplo a seguir descreve como você pode usar o `makeEwsRequestAsync` método.

1. No XML, substitua as IDs de item e atributos relevantes da operação dos EWS por valores apropriados.

1. Inclua a solicitação SOAP como um argumento para o  _parâmetro de_ dados de `makeEwsRequestAsync`.

1. Especifique uma função de retorno de chamada e uma chamada `makeEwsRequestAsync`.

1. Na função de retorno de chamada, verifique os resultados da operação na resposta SOAP.

1. Use os resultados da operação dos EWS de acordo com as suas necessidades.

The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Tabela 2. Operações compatíveis do EWS**

|**Operação do EWS**|**Descrição**|
|:-----|:-----|
|[Operação CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)|Copia os itens especificados e coloca os novos itens em uma pasta designada no repositório do Exchange.|
|[Operação CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)|Cria pastas no local especificado no repositório do Exchange.|
|[Operação CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)|Cria os itens especificados no repositório do Exchange.|
|[Operação ExpandDL](/exchange/client-developer/web-service-reference/expanddl-operation)|Exibe a associação completa das listas de distribuição.|
|[Operação FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)|Enumera uma lista de conversas na pasta especificada no repositório do Exchange.|
|[Operação FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)|Localiza subpastas de uma pasta identificada e retorna um conjunto de propriedades que descreve o conjunto de subpastas.|
|[Operação FindItem](/exchange/client-developer/web-service-reference/finditem-operation)|Identifica os itens que estão localizados em uma pasta especificada no repositório do Exchange.|
|[Operação GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)|Obtém um ou mais conjuntos de itens que estão organizados em nós em uma conversa.|
|[Operação GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)|Obtém as propriedades especificadas e o conteúdo de pastas do repositório do Exchange.|
|[Operação GetItem](/exchange/client-developer/web-service-reference/getitem-operation)|Obtém as propriedades especificadas e o conteúdo de itens do repositório do Exchange.|
|[Operação GetUserAvailability](/exchange/client-developer/web-service-reference/getuseravailability-operation)|Fornece informações detalhadas sobre a disponibilidade de um conjunto de usuários, salas e recursos em um período especificado.|
|[Operação MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)|Move mensagens de email para a pasta Lixo Eletrônico e adiciona ou remove, adequadamente, remetentes das mensagens na lista de remetentes bloqueados.|
|[Operação MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)|Move itens para uma única pasta de destino no repositório do Exchange.|
|[Operação ResolveNames](/exchange/client-developer/web-service-reference/resolvenames-operation)|Resolve endereços de email e nomes de exibição ambíguos.|
|[Operação SendItem](/exchange/client-developer/web-service-reference/senditem-operation)|Envia mensagens de email que estão localizadas no repositório do Exchange.|
|[Operação UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)|Modifica as propriedades de pastas existentes no repositório do Exchange.|
|[Operação UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)|Modifica as propriedades de itens existentes no repositório do Exchange.|

 > [!NOTE]
 > Não é possível atualizar (ou criar) itens FAI (Informações Associadas da Pasta) usando um suplemento. Essas mensagens ocultas são armazenadas em uma pasta e usadas para armazenar diversas configurações e dados auxiliares.  Tentar usar a operação UpdateItem gera um erro ErrorAccessDenied: "A extensão do Office não tem permissão para atualizar esse item". Se preferir, use a [API Gerenciada do EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) para atualizar esses itens usando um cliente do Windows ou um aplicativo para servidores. Recomenda-se cuidado já que as estruturas de dados internos de tipo de serviço estão sujeitas a alterações e podem invalidar sua solução.

## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a>Considerações sobre autenticação e permissão para makeEwsRequestAsync

Quando você usa o `makeEwsRequestAsync` método, a solicitação é autenticada usando as credenciais da conta de email do usuário atual. O `makeEwsRequestAsync` método gerencia as credenciais para você para que você não tenha que fornecer credenciais de autenticação com sua solicitação.

> [!NOTE]
> O administrador do servidor deve usar o cmdlet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) para definir o parâmetro _OAuthAuthentication_ `true` como no diretório EWS `makeEwsRequestAsync` do servidor de Acesso ao Cliente para permitir que o método faça solicitações EWS.

Para usar o método `makeEwsRequestAsync` , o suplemento deve solicitar a permissão de caixa de correio de leitura **/** gravação no manifesto. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: defina o **\<Permissions\>** elemento **como ReadWriteMailbox**.
- **Manifesto do Teams (** versão prévia):defina a propriedade "name" de um objeto na matriz "authorization.permissions.resourceSpecific" como "Mailbox.ReadWrite.User".

Para obter informações sobre como usar a **permissão de caixa de correio** de leitura/gravação, consulte [a permissão de caixa de correio de leitura/gravação](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission).

## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
- [Como lidar com limitações de política de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md)
- [Referência do EWS para Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Aplicativos de email para Outlook e EWS no Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

Consulte o seguinte para criar serviços de back-end para suplementos usando ASP.NET Web API.

- [Criar um serviço Web para um suplemento do Office usando a API Web ASP.NET](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [Noções básicas sobre a criação de um serviço HTTP usando a API Web ASP.NET](https://dotnet.microsoft.com/apps/aspnet/apis)
