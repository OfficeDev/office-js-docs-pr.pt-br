---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 02/26/2019
localization_priority: Priority
ms.openlocfilehash: 233bc6770faefaa0e101fd01c353e7ce0df972a1
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359244"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento. Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.

O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent): adicionado um novo objeto que representa o conteúdo de um anexo.
- [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) - Adicionar um novo objeto que representa o conjunto de locais em um compromisso.
- [InternetHeaders](/javascript/api/outlook/office.internetheaders): Adicionado um novo objeto que representa os cabeçalhos de Internet de um item de mensagem.
- [LocationDetails](/javascript/api/outlook/office.locationdetails) - Adicionar um novo objeto que representa um local. Somente leitura.
- [LocationIdentifier](/javascript/api/outlook/office.locationidentifier) - Adicionar um novo objeto que representa a id de um local.
- [SharedProperties](/javascript/api/outlook/office.sharedproperties): adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback): adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.
- [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) – adicionar uma nova propriedade que representa o conjunto de locais em um compromisso.
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) – Adicionado um novo método para acessar o conteúdo de um anexo específico.
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) – Adicionado um novo método que obtém os anexos de um item no modo de redação.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – Adicionada uma nova função que retorna os dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback): adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders): adicionado uma nova propriedade que representa os cabeçalhos de Internet de um item de mensagem.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus): adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions): adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.
- [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) - Adicionar uma nova enumeração que especifica o tipo de local do compromisso.
- [Office.EventType](/javascript/api/office/office.eventtype) — Modificado para dar suporte aos eventos AttachmentsChanged, EnhancedLocationsChanged e OfficeThemeChanged por meio da adição das entradas `AttachmentsChanged`, `EnhancedLocationsChanged` e `OfficeThemeChanged` respectivamente.
- [Elemento do manifesto SupportsSharedFolders](../../manifest/supportssharedfolders.md): adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md). Define se o suplemento está disponível nos cenários de representante.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)
