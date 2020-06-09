---
title: Usar os Serviços Web do Exchange a partir de um suplemento do Outlook
description: Fornece um exemplo que mostra como um suplemento do Outlook pode solicitar informações dos Serviços Web do Exchange.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: a1141570c14b6905584f9398b629a75b477d3870
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604505"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Chamar serviços Web de um suplemento do Outlook

O suplemento pode usar os EWS (Serviços Web do Exchange) de um computador que esteja executando o Exchange Server 2013, um serviço Web que está disponível no servidor que fornece o local de origem para interface do usuário do suplemento ou um serviço Web que está disponível na Internet. Este artigo fornece um exemplo que mostra como um suplemento do Outlook pode solicitar informações dos EWS.

A maneira usada para chamar um serviço Web varia com base em onde o serviço Web está localizado. A Tabela 1 lista as diferentes maneiras que podem ser usadas para chamar um serviço Web baseado no local.


**Tabela 1. Maneiras de chamar serviços Web de um suplemento do Outlook**

<br/>

|**Local do serviço Web**|**Maneira de chamar o serviço Web**|
|:-----|:-----|
|O servidor Exchange que hospeda a caixa de correio do cliente|Use o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para chamar operações EWS com suporte dos suplementos. O servidor Exchange que hospeda a caixa de correio também expõe os EWS.|
|O servidor Web que fornece o local de origem para a interface do usuário|Chame o serviço Web usando técnicas JavaScript padrão. O código JavaScript no quadro da interface do usuário é executado no contexto do servidor Web que fornece a interface do usuário. Portanto, ele pode chamar serviços Web nesse servidor sem causar um erro de script entre sites.|
|Todos os outros locais|Crie um proxy para o serviço Web no servidor Web que fornece o local de origem para a interface do usuário. Se você não fornecer um proxy, erros de script entre sites impedirão a execução do suplemento. Uma maneira de fornecer um proxy é usar JSON/P. Para saber mais, confira [Privacidade e segurança para suplementos do Office](../develop/privacy-and-security.md).|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>Usar o método makeEwsRequestAsync para acessar operações dos EWS

Você pode usar o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para fazer uma solicitação dos EWS ao servidor Exchange que hospeda a caixa de correio do usuário.

Os EWS oferecem suporte a diferentes operações em um servidor Exchange, por exemplo, operações no nível do item para copiar, localizar, atualizar ou enviar um item e operações no nível da pasta para criar, acessar ou atualizar uma pasta. Para executar uma operação dos EWS, crie uma solicitação SOAP XML para a operação. Quando a operação termina, você recebe uma resposta SOAP XML que contém dados que são relevantes para a operação. As solicitações e respostas SOAP dos EWS seguem o esquema definido no arquivo Messages.xsd. Como outros arquivos de esquema dos EWS, o arquivo Message.xsd está localizado no diretório virtual do IIS que hospeda os EWS.

Para usar o `makeEwsRequestAsync` método para iniciar uma operação do EWS, forneça o seguinte:

- O XML para a solicitação SOAP dessa operação dos EWS, como um argumento para o parâmetro _data_

- Um método de retorno (como o argumento _callback_)

- Outros dados de entrada opcionais para esse método de retorno de chamada (como o argumento _userContext_)

Quando a solicitação SOAP do EWS estiver concluída, o Outlook chamará o método callback com um argumento, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult) . O método de retorno de chamada pode acessar duas propriedades do `AsyncResult` objeto: a `value` propriedade, que contém a resposta SOAP XML da operação do EWS e, opcionalmente, a `asyncContext` propriedade, que contém todos os dados passados como o `userContext` parâmetro. Normalmente, o método de retorno de chamada analisa o XML na resposta SOAP para obter qualquer informação relevante e processa essas informações de forma adequada.


## <a name="tips-for-parsing-ews-responses"></a>Dicas para analisar respostas dos EWS

Ao analisar uma resposta SOAP de uma operação dos EWS, observe os seguintes problemas que dependem do navegador:


- Especifique o prefixo para um nome de marca ao usar o método DOM `getElementsByTagName` , para incluir suporte para o Internet Explorer.

  `getElementsByTagName`comporta de forma diferente dependendo do tipo de navegador. Por exemplo, uma resposta do EWS pode conter o seguinte XML (formatado e abreviado para fins de exibição):

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   O código, como no seguinte, funcionaria em um navegador como o Chrome para obter o XML delimitado pelas `ExtendedProperty` marcas:

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   No Internet Explorer, você precisa incluir o prefixo `t:` do nome da marca, conforme mostrado abaixo:

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- Use a propriedade DOM `textContent` para obter o conteúdo de uma marca em uma resposta do EWS, conforme mostrado abaixo:

   ```js
      content = $.parseJSON(value.textContent);
   ```

   Outras propriedades como o `innerHTML` podem não funcionar no Internet Explorer para algumas marcas em uma resposta do EWS.


## <a name="example"></a>Exemplo

O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para obter o assunto de um item. Este exemplo inclui as seguintes três funções:

-  `getSubjectRequest`&ndash;Obtém uma ID de item como entrada e retorna o XML para que a solicitação SOAP chame `GetItem` o item especificado.

-  `sendRequest`&ndash;Chamadas `getSubjectRequest` para obter a solicitação SOAP para o item selecionado e, em seguida, passa a solicitação SOAP e o método de retorno de chamada, `callback` , para `makeEwsRequestAsync` obter o assunto do item especificado.

-  `callback` &ndash; processa a resposta SOAP que inclui o assunto e outras informações sobre o item especificado.


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


## <a name="ews-operations-that-add-ins-support"></a>Operações dos EWS compatíveis com suplementos

Os suplementos do Outlook podem acessar um subconjunto de operações disponíveis no EWS por meio do `makeEwsRequestAsync` método. Se você não estiver familiarizado com as operações do EWS e como usar o `makeEwsRequestAsync` método para acessar uma operação, comece com um exemplo de solicitação SOAP para personalizar o argumento de _dados_ .

O procedimento a seguir descreve como você pode usar o `makeEwsRequestAsync` método:

1. No XML, substitua as IDs de item e atributos relevantes da operação dos EWS por valores apropriados.

2. Inclua a solicitação SOAP como um argumento para o parâmetro de _dados_ de `makeEwsRequestAsync` .

3. Especifique um método de retorno de chamada e chame `makeEwsRequestAsync` .

4. No método de retorno de chamada, verifique os resultados da operação na resposta SOAP.

5. Use os resultados da operação dos EWS de acordo com as suas necessidades.

A tabela a seguir lista as operações dos EWS compatíveis com suplementos. Para ver exemplos de solicitações e respostas SOAP, escolha o link para cada operação. Para saber mais sobre operações dos EWS, confira [Operações dos EWS no Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).

**Tabela 2. Operações compatíveis do EWS**

<br/>

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

Quando você usa o `makeEwsRequestAsync` método, a solicitação é autenticada usando as credenciais da conta de email do usuário atual. O `makeEwsRequestAsync` método gerencia as credenciais para você para que você não precise fornecer credenciais de autenticação com a solicitação.

> [!NOTE]
> O administrador do servidor deve usar o cmdlet [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) ou [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) para definir o parâmetro _OAUTHAUTHENTICATION_ como **true** no diretório EWS do servidor de acesso para cliente a fim de habilitar o `makeEwsRequestAsync` método para fazer solicitações do EWS.

O suplemento deve especificar a `ReadWriteMailbox` permissão em seu manifesto do suplemento para usar o `makeEwsRequestAsync` método. Para saber mais sobre como usar a `ReadWriteMailbox` permissão, confira a seção [ReadWriteMailbox permissão](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) em [noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md).

## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../develop/privacy-and-security.md)
- [Como lidar com limitações de política de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md)
- [Referência do EWS para Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Aplicativos de email para Outlook e EWS no Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

Veja os tópicos a seguir para criar serviços de back-end para suplementos usando a API Web ASP.NET:

- [Criar um serviço Web para um suplemento do Office usando a API Web ASP.NET](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)
- [Noções básicas sobre a criação de um serviço HTTP usando a API Web ASP.NET](https://www.asp.net/web-api)
