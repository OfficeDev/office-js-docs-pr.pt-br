---
title: Outlook conjunto de requisitos de API de complemento 1.11
description: Conjunto de requisitos 1.11 para Outlook api de complemento.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 384e872b44b213b60a1b651f85ac315cd06cf082
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744140"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook conjunto de requisitos de API de complemento 1.11

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

## <a name="whats-new-in-111"></a>Novidades no 1.11?

O conjunto de requisitos 1.11 inclui todos os recursos do conjunto de [requisitos 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Ele adicionou os seguintes recursos.

- Adicionados novos eventos para a [ativação baseada em eventos](../../../outlook/autolaunch.md#supported-events).
- Adicionadas APIs SessionData.

### <a name="change-log"></a>Log de mudanças

- Adicionado [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties): adiciona uma nova propriedade para gerenciar os dados de sessão de um item no modo Redação.
- Adicionado [Office. SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true): adiciona um novo objeto que representa os dados de sessão de um item de composição.
- Adicionados novos eventos para [a ativação baseada em eventos](../../../outlook/autolaunch.md#supported-events): adiciona suporte para os seguintes eventos.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- Adicionado [Office. AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true): adiciona um objeto que dá suporte ao `OnAppointmentTimeChanged` evento.
- Adicionado [Office. AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true): adiciona um objeto que dá suporte aos `OnAppointmentAttachmentsChanged` eventos e`OnMessageAttachmentsChanged`.
- Adicionado [Office. InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true): adiciona um objeto que dá suporte ao `OnInfoBarDismissClicked` evento.
- Adicionado [Office. RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true): adiciona um objeto que dá suporte aos `OnAppointmentAttendeesChanged` eventos e`OnMessageRecipientsChanged`.
- Adicionado [Office. RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true): adiciona um objeto que dá suporte ao `OnAppointmentRecurrenceChanged` evento.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
