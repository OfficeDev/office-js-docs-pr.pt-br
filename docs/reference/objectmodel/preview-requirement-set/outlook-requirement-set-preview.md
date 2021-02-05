---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em visualização para os complementos do Outlook.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: 39dd1221f4dea9674c89cdaad20024ce408f8db3
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104837"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto da API de complemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um complemento do Outlook.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Você pode visualizar recursos no Outlook na Web configurando o lançamento direcionado [no locatário do Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) "Configurar o acesso de visualização" é notado nesta página para recursos aplicáveis.
>
> For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview). "Solicitar acesso de visualização" é notado nesses recursos.

O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Ativação de um complemento em itens protegidos pelo Gerenciamento de Direitos de Informação (IRM)

Os complementos agora podem ser ativados em itens protegidos por IRM. Para ativar esse recurso, um administrador de locatários precisa habilitar o direito de uso definindo a opção Permitir acesso `OBJMODEL` **programático** personalizado de política no Office. Consulte [Direitos de uso e descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.

**Disponível em:** Outlook no Windows, a partir do build 13229.10000 (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriedades de calendário adicionais

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Adicionada uma nova enum `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Ativação baseada em evento

Adicionado suporte para a funcionalidade de ativação baseada em eventos em complementos do Outlook. Confira [Configurar seu complemento do Outlook para ativação baseada em eventos](../../../outlook/autolaunch.md) para saber mais.

#### <a name="launchevent-extension-point"></a>[Ponto de extensão LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Adicionado `LaunchEvent` suporte ao ponto de extensão para manifesto. Ele configura a funcionalidade de ativação baseada em eventos.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="launchevents-manifest-element"></a>[Elemento de manifesto LaunchEvents](../../manifest/launchevents.md)

Elemento `LaunchEvents` adicionado ao manifesto. Ele dá suporte à configuração da funcionalidade de ativação baseada em eventos.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="runtimes-manifest-element"></a>[Elemento de manifesto runtimes](../../manifest/runtimes.md)

Adicionado suporte do Outlook ao elemento `Runtimes` de manifesto. Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="mail-signature"></a>Assinatura de email

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

Adicionada uma nova função ao objeto que adiciona ou substitui a `Body` assinatura no corpo do item no modo Redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo Redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

Adicionada uma nova função que obtém o tipo de composição de uma mensagem no modo redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no item no modo redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officemailboxenumscomposetype"></a>[Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

Adicionada uma nova enum `ComposeType` disponível no modo Redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="notification-messages-with-actions"></a>Mensagens de notificação com ações

Esse recurso permite que o seu complemento inclua uma mensagem de notificação com uma ação personalizada além da ação **Padrão Descartar.** No Outlook na Web moderno, esse recurso está disponível somente no modo Redação.

#### <a name="officenotificationmessagedetailsactions"></a>[Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Adicionada uma nova propriedade que permite adicionar uma `InsightMessage` notificação com uma ação personalizada.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officenotificationmessageaction"></a>[Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsactiontype"></a>[Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Adicionada uma nova enum `ActionType` .

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Adicionado um novo `InsightMessage` tipo à `ItemNotificationMessageType` enum.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Os dados da sessão

#### <a name="officesessiondata"></a>[Office.SessionData](/javascript/api/outlook/office.sessiondata)

Adicionado um novo objeto que representa os dados de sessão de um item.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo redação.

**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
