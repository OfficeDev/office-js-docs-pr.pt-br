---
title: Outlook conjunto de requisitos de visualização de API de complemento
description: Recursos e APIs que estão atualmente em visualização para Outlook de complementos.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2d1efa2b2dca5a88a56fb5f54a84b790e08745ec
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681645"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook conjunto de requisitos de visualização de API de complemento

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Você pode ser capaz de visualizar recursos no Outlook na Web configurando a versão direcionada [em seu locatário Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) "Configure preview access" é notado nesta página para recursos aplicáveis.
>
> Para outros recursos, você pode solicitar acesso aos bits de visualização para Outlook na Web usando sua conta Microsoft 365 concluindo e enviando [esse formulário](https://aka.ms/OWAPreview). "Solicitar acesso de visualização" é notado nesses recursos.

O conjunto de requisitos de visualização inclui todos os recursos do conjunto [de requisitos 1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Ativação do complemento em itens protegidos pelo IRM (Gerenciamento de Direitos de Informação)

Os complementos agora podem ser ativados em itens protegidos por IRM. Para ativar esse recurso, um administrador de locatário precisa habilitar o direito de uso definindo a opção Permitir política personalizada de acesso `OBJMODEL` programático em  Office. Confira [Direitos de uso e descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.

**Disponível em**: Outlook no Windows, começando com a com build 13229.10000 (conectada a uma assinatura Microsoft 365 de terceiros)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriedades de calendário adicionais

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a propriedade de evento de todos os dias de um compromisso no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Adicionado um novo número `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="delay-delivery-time"></a>Atrasar o tempo de entrega

#### <a name="officecontextmailboxitemdelaydeliverytime"></a>[Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que retorna um objeto que permite gerenciar a data e a hora de entrega de uma mensagem no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officedelaydeliverytime"></a>[Office. DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que permite gerenciar a data e a hora de entrega de uma mensagem no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Ativação baseada em evento

Esse recurso foi lançado no [conjunto de requisitos 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). No entanto, eventos adicionais agora estão disponíveis na visualização. Para saber mais, consulte [Eventos com suporte.](../../../outlook/autolaunch.md#supported-events)

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#officeTheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="shared-mailboxes"></a>Caixas de correio compartilhadas

O suporte a recursos para pastas compartilhadas (ou seja, acesso de representante) foi lançado no conjunto [de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). No entanto, o suporte para caixas de correio compartilhadas agora está disponível na visualização. Para saber mais, consultar [Habilitar pastas compartilhadas e cenários de caixas de correio compartilhada](../../../outlook/delegate-access.md).

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365), Outlook na Web (moderno), Outlook no Mac

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
