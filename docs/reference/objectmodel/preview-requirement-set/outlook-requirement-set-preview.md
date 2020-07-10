---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: d91d1e16382a9ada71210657d6111f548c85ccfd
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094418"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center). "Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.
>
> Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview). "Solicitar acesso de visualização" é observado nesses recursos.

O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="additional-calendar-properties"></a>Propriedades de calendário adicionais

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. isAllDayEvent](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. sensibilidade](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

<br>

---

---

### <a name="append-on-send"></a>Anexar ao enviar

Para saber mais sobre como usar o recurso Append-on-Send, confira [implementar anexar ao enviar em seu suplemento do Outlook](../../../outlook/append-on-send.md).

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="async-versions-of-display-apis"></a>Versões assíncronas de `display` APIs

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[Office. Context. Mailbox. displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

Foi adicionada uma nova função ao `Mailbox` objeto que exibe um compromisso existente. Esta é a versão assíncrona do `displayAppointmentForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[Office. Context. Mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

Foi adicionada uma nova função ao `Mailbox` objeto que exibe uma mensagem existente. Esta é a versão assíncrona do `displayMessageForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[Office. Context. Mailbox. displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

Foi adicionada uma nova função ao `Mailbox` objeto que exibe um novo formulário de compromisso. Esta é a versão assíncrona do `displayNewAppointmentForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[Office. Context. Mailbox. displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

Foi adicionada uma nova função ao `Mailbox` objeto que exibe um novo formulário de mensagem. Esta é a versão assíncrona do `displayNewMessageForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[Office. Context. Mailbox. Item. displayReplyAllFormAsync](office.context.mailbox.item.md#methods)

Foi adicionada uma nova função ao `Item` objeto que exibe o formulário "responder a todos" no modo de leitura. Esta é a versão assíncrona do `displayReplyAllForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[Office. Context. Mailbox. Item. displayReplyFormAsync](office.context.mailbox.item.md#methods)

Foi adicionada uma nova função ao `Item` objeto que exibe o formulário "responder" no modo de leitura. Esta é a versão assíncrona do `displayReplyForm` método.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Ativação baseada em evento

Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.

#### <a name="launchevent-extension-point"></a>[Ponto de extensão LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto. Ele configura a funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))

#### <a name="launchevents-manifest-element"></a>[Elemento de manifesto LaunchEvents](../../manifest/launchevents.md)

`LaunchEvents`Elemento adicionado ao manifesto. Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))

#### <a name="runtimes-manifest-element"></a>[Elemento de manifesto de runtimes](../../manifest/runtimes.md)

Adicionado suporte do Outlook ao `Runtimes` elemento manifest. Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))

<br>

---

---

### <a name="get-all-custom-properties"></a>Obter todas as propriedades personalizadas

#### <a name="custompropertiesgetall"></a>[CustomProperties. getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

Foi adicionada uma nova função ao `CustomProperties` objeto que obtém todas as propriedades personalizadas.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno), Outlook no Mac (conectado à assinatura do Microsoft 365), Outlook no Android, Outlook no Ios

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (clássico)

<br>

---

---

### <a name="mail-signature"></a>Assinatura de email

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. composetype](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Adição de uma nova enumeração `ComposeType` disponível no modo de composição.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365)

<br>

---

---

### <a name="single-sign-on-sso"></a>SSO (logon único)

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.

**Disponível em**: Outlook no Windows (conectado à assinatura do Microsoft 365), Outlook no Mac (conectado à assinatura do Microsoft 365), Outlook na Web (moderno), Outlook na Web (clássico)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
