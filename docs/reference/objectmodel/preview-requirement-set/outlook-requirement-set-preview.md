---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512857"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los. Você também precisará ingressar no [programa Office Insider](https://products.office.com/office-insider).

O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="add-in-commands"></a>Comandos de suplemento

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.

**Disponível em**: Outlook na web (clássico)

### <a name="attachments"></a>Attachments

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Adicionado um novo objeto que representa o conteúdo de um anexo.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

Adicionar um novo método para acessar o conteúdo de um anexo específico.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

Adicionado um novo método que obtém um item anexo no modo de redação.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

Adicionado `AttachmentsChanged` evento `Item`.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

### <a name="delegate-access"></a>Acesso de representante

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[Elemento manifesto SupportsSharedFolders](../../manifest/supportssharedfolders.md)

Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md). Define se o suplemento está disponível nos cenários de representante.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

### <a name="enhanced-location"></a>Local aprimorado

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Adicionado um novo objeto que representa o conjunto de locais em um compromisso.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

Adicionado um novo objeto que representa um local. Somente leitura.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

Adicionado um novo objeto que representa a id de um local.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

Adicionada uma nova enumeração que especifica o tipo de local do compromisso.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

Adicionado `EnhancedLocationsChanged` evento `Item`.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em**: Office 2019 para Windows (assinatura do Office 365), Outlook na web (clássico)

### <a name="internet-headers"></a>Cabeçalhos de Internet

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

Adicionado um novo objeto que representa os cabeçalhos de Internet de um item de mensagem.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

Adicionada uma nova propriedade que representa os cabeçalhos de Internet de um item de mensagem.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365)

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.

**Disponível em**: Outlook 2019 para Windows (assinatura do Office 365), Outlook 2019 para Mac, Outlook na Web (Office 365 e Outlook.com), Outlook na web (clássico)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)
