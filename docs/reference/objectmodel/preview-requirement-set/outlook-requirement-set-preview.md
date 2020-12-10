---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 2f83f81dcf7aa7ab0e3a48fff4279c1e08ba6286
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612747"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center). "Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.
>
> Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview). "Solicitar acesso de visualização" é observado nesses recursos.

O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1,9](../requirement-set-1.9/outlook-requirement-set-1.9.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Ativação de suplementos em itens protegidos por IRM (gerenciamento de direitos de informação)

Agora, os suplementos podem ser ativados em itens protegidos por IRM. Para ativar esse recurso, um administrador de locatários precisa habilitar o `OBJMODEL` direito de uso, configurando a opção permitir política personalizada de **acesso programático** no Office. Confira os [direitos de uso e as descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.

**Disponível em**: Outlook no Windows, começando com a compilação 13229,10000 (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriedades de calendário adicionais

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. isAllDayEvent](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. sensibilidade](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Ativação baseada em evento

Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.

#### <a name="launchevent-extension-point"></a>[Ponto de extensão LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto. Ele configura a funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="launchevents-manifest-element"></a>[Elemento de manifesto LaunchEvents](../../manifest/launchevents.md)

`LaunchEvents`Elemento adicionado ao manifesto. Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="runtimes-manifest-element"></a>[Elemento de manifesto de runtimes](../../manifest/runtimes.md)

Adicionado suporte do Outlook ao `Runtimes` elemento manifest. Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.

**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="mail-signature"></a>Assinatura de email

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. composetype](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

Adição de uma nova enumeração `ComposeType` disponível no modo de composição.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="notification-messages-with-actions"></a>Mensagens de notificação com ações

Este recurso permite que o suplemento inclua uma mensagem de notificação com uma ação personalizada além da ação padrão de **ignorar** . No Outlook moderno na Web, este recurso está disponível somente no modo de composição.

#### <a name="officenotificationmessagedetailsactions"></a>[Office. NotificationMessageDetails. Actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Adicionada uma nova propriedade que permite que você adicione uma `InsightMessage` notificação com uma ação personalizada.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officenotificationmessageaction"></a>[Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsactiontype"></a>[Office. MailboxEnums. ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Foi adicionada uma nova enumeração `ActionType` .

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office. MailboxEnums. ItemNotificationMessageType. InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Adicionado um novo tipo `InsightMessage` à `ItemNotificationMessageType` enumeração.

**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Os dados da sessão

#### <a name="officesessiondata"></a>[Office. SessionData](/javascript/api/outlook/office.sessiondata)

Adicionado um novo objeto que representa os dados de sessão de um item.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office. Context. Mailbox. Item. sessionData](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo de composição.

**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
