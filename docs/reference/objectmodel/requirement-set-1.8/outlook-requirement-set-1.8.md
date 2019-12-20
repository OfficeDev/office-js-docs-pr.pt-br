---
title: Conjunto de requisitos 1.8 da API de suplemento do Outlook
description: ''
ms.date: 12/17/2019
localization_priority: Priority
ms.openlocfilehash: 1a542a0b00bd04674c4da4c45909ff03a152983c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814156"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Conjunto de requisitos 1.8 da API de suplemento do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs de JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

## <a name="whats-new-in-18"></a>O que há de novo no 1.8?

O conjunto de requisitos 1.8 inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs para anexos, categorias, acesso de representante, local aprimorado, cabeçalhos da Internet e recursos de bloqueio ao enviar.
- Adicionado o parâmetro opcional `options` ao Event.completed.
- Adicionado suporte para os eventos AttachmentsChanged e EnhancedLocationsChanged.

### <a name="change-log"></a>Log de mudanças

- Adicionado [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8): adiciona um novo objeto que representa o conteúdo de um anexo.
- Adicionado [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8): adiciona um novo objeto que representa as categorias de um item.
- Adicionado [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8): adiciona um novo objeto que representa os detalhes de uma categoria (nome e cor associada).
- Adicionado [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8): adiciona um novo objeto que representa o conjunto de locais em um compromisso.
- Adicionado [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8): adiciona um novo objeto que representa os cabeçalhos de Internet personalizados de um item de mensagem. Somente modo de redação.
- Adicionado [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8): adiciona um novo objeto que representa um local. Somente leitura.
- Adicionado [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8): adiciona um novo objeto que representa a ID de um local.
- Adicionado [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8): adiciona um novo objeto que representa a lista mestre de categorias em uma caixa de correio.
- Adicionado [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8): adiciona um novo objeto que representa as propriedades de um compromisso ou item de mensagem em uma pasta, calendário ou caixa de correio compartilhados.
- Adicionado [elemento de manifesto SupportsSharedFolders](../../manifest/supportssharedfolders.md): adiciona um elemento filho ao elemento de manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md). Define se o suplemento está disponível nos cenários de representante.
- Adicionado [Office.context.mailbox.masterCategories](office.context.mailbox.md#properties): adiciona uma nova propriedade que representa a lista principal de categorias em uma caixa de correio.
- Adicionado [Office.context.mailbox.item.categories](office.context.mailbox.item.md#properties): adiciona uma nova propriedade que representa o conjunto de categorias em um item.
- Adicionado [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#methods): adiciona um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada em base64 a uma mensagem ou compromisso.
- Adicionado [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#properties): adiciona uma nova propriedade que representa o conjunto de locais em um compromisso.
- Adicionado [Office.context.mailbox.item.getAllInternetHeadersAsync](office.context.mailbox.item.md#methods): adiciona um novo método que obtém todos os cabeçalhos da Internet para um item de mensagem. Somente modo de leitura.
- Adicionado [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#methods): adiciona um novo método para obter o conteúdo de um anexo específico.
- Adicionado [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#methods): adiciona um novo método que obtém os anexos de um item no modo de redação.
- Adicionado [Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#methods): adiciona um novo método que obtém a ID de um compromisso ou item de mensagem salvo.
- Adicionado [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods): adiciona um novo método que obtém um objeto que representa as sharedProperties de um compromisso ou item de mensagem.
- Adicionado [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#properties): adiciona uma nova propriedade que representa os cabeçalhos da Internet personalizados em um item de mensagem. Somente modo de redação.
- Modificado [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): adiciona um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`. Esse valor é usado para cancelar a execução de um evento.
- Adicionado [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): adiciona uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.
- Adicionado [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8): adiciona uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.
- Adicionado [Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8): adiciona uma nova enumeração que especifica as cores disponíveis para serem associadas às categorias.
- Adicionado [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8): adiciona uma nova enumeração de sinalizador de bit que especifica as permissões de representante.
- Adicionado [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8): adiciona uma nova enumeração que especifica o tipo de um local de compromisso.
- Modificado [Office.EventType](/javascript/api/office/office.eventtype): adiciona suporte para os eventos `AttachmentsChanged` e `EnhancedLocationsChanged`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
