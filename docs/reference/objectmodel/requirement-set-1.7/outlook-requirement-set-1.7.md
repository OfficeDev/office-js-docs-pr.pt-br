---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.7
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 793a1e1c2c3dd014f104ab264f4954369b591162
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597023"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.7

O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o mais recente.

## <a name="whats-new-in-17"></a>Quais as novidades da versão 1.7?

O conjunto de requisitos versão 1.7 inclui todos os recursos do [Conjunto de requisitos versão 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs referentes ao padrão de recorrência em compromissos e mensagens que são solicitações de reunião.
- Foi modificada a propriedade item.from também estar disponível no modo Redação.
- Adicionado suporte para eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged.

### <a name="change-log"></a>Log de mudanças

- Adicionado o [From](/javascript/api/outlook/office.from?view=outlook-js-1.7): adiciona um novo objeto que fornece um método para obter o valor from.
- Adicionado o [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7): adiciona um novo objeto que fornece um método para obter o valor organizer.
- Adicionado o [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7): adiciona um novo objeto que fornece métodos para obter e definir o padrão de recorrência de compromissos, mas obtém apenas o padrão de recorrência de mensagens de solicitações de reunião.
- Adicionado o [RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone?view=outlook-js-1.7): adiciona um novo objeto que representa a configuração de fuso horário do padrão de recorrência.
- Adicionado o [SeriesTime](/javascript/api/outlook/office.seriestime?view=outlook-js-1.7): adiciona um novo objeto que fornece métodos para obter e definir as datas e horas de compromissos em uma série recorrente e obter as datas e horas de solicitações de reunião em uma série recorrente.
- Adicionado o [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#methods): adiciona um novo método que adiciona um manipulador de eventos para um evento com suporte.
- Modificado [Office.context.mailbox.item.from](office.context.mailbox.item.md#properties): Adiciona a capacidade de adquirir o valor from no modo de Redação.
- Modificado [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#properties): Adiciona a capacidade de adquirir o valor organizer no modo de Redação.
- Adicionado o [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#properties): adiciona uma nova propriedade que obtém ou define um objeto que fornece métodos de gerenciamento do padrão de recorrência de um item de compromisso. Essa propriedade também pode ser usada para obter o padrão de recorrência de um item de solicitação de reunião.
- Adicionado o [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#methods): adiciona um novo método que remove um manipulador de eventos.
- Adicionado o [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#properties): adiciona uma nova propriedade que obtém a ID da série à qual uma ocorrência pertence.
- Adicionado o [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days?view=outlook-js-1.7): adiciona uma nova enumeração que especifica o dia da semana ou o tipo de dia.
- Adicionado o [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month?view=outlook-js-1.7): adiciona uma nova enumeração que especifica o mês.
- Adicionado o [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone?view=outlook-js-1.7): adiciona uma nova enumeração que especifica o fuso horário aplicado à recorrência.
- Adicionado o [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype?view=outlook-js-1.7): adiciona uma nova enumeração que especifica o tipo de recorrência.
- Adicionado o [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber?view=outlook-js-1.7): adiciona uma nova enumeração que especifica a semana do mês.
- Modificado [Office.EventType](/javascript/api/office/office.eventtype): Adiciona suporte para eventos `RecurrenceChanged`, `RecipientsChanged`, e `AppointmentTimeChanged`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
