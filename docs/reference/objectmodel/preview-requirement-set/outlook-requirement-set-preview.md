---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: f2030ca40a83694c1c1c92ca4b956a9294cb3e0a
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804987"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los. Você também precisará ingressar no [programa Office Insider](https://products.office.com/office-insider).

O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="attachments"></a>Anexos

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Adicionado um novo objeto que representa o conteúdo de um anexo.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

Adicionar um novo método para acessar o conteúdo de um anexo específico.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

Adicionado um novo método que obtém um item anexo no modo de redação.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

Evento `AttachmentsChanged` adicionado ao `Item`.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

---

### <a name="block-on-send"></a>Bloquear ao enviar

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.

**Disponível em**: Outlook na Web (classic)

---

### <a name="categories"></a>Categorias

No Outlook, um usuário pode agrupar mensagens e compromissos usando uma categoria para codificá-los por cor. O usuário define as categorias em uma lista mestra em sua caixa de correio. Ele pode, em seguida, aplicar uma ou mais categorias a um item.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[Categories](/javascript/api/outlook/office.categories)

Adicionou um novo objeto que representa a categoria de um item.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[CategoryDetails](/javascript/api/outlook/office.categorydetails)

Adicionou um novo objeto que representa os detalhes de uma categoria (seu nome e cor associada).

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[MasterCategories](/javascript/api/outlook/office.mastercategories)

Adicionou um novo objeto que representa a lista mestra de categorias em uma caixa de correio.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

Adicionou uma nova propriedade que representa a lista mestra de categorias em uma caixa de correio.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

Adicionou uma nova propriedade que representa o conjunto de categorias em um item.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

Adicionou uma nova enumeração que especifica as cores disponíveis a serem associadas a categorias. 

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

---

### <a name="delegate-access"></a>Acesso de representante

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

Adicionado um novo método que obtém o ID de um compromisso ou item de mensagem salvo.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[Elemento manifesto SupportsSharedFolders](../../manifest/supportssharedfolders.md)

Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md). Define se o suplemento está disponível nos cenários de representante.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

---

### <a name="enhanced-location"></a>Local aprimorado

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Adicionado um novo objeto que representa o conjunto de locais em um compromisso.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

Adicionado um novo objeto que representa um local. Somente leitura.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

Adicionado um novo objeto que representa a id de um local.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

Adicionada uma nova enumeração que especifica o tipo de local do compromisso.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

Evento `EnhancedLocationsChanged` adicionado ao `Item`.

**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)

---

### <a name="internet-headers"></a>Cabeçalhos de Internet

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

Adicionado um novo objeto que representa os cabeçalhos de internet personalizados de um item de mensagem.

**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

Adicionada uma nova propriedade que representa os cabeçalhos de internet personalizados de um item de mensagem.

**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)

---

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.

**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)
- [Conjuntos de requisitos e clientes com suporte](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
