---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.7
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2e233c614a902a724ead0240c4e5229e1053ee81
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432309"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.7

O subconjunto de APIs de suplemento do Outlook para as APIs de JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

## <a name="whats-new-in-17"></a>Quais as novidades da versão 1.7?

O conjunto de requisitos versão 1.7 inclui todos os recursos do [Conjunto de requisitos versão 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Ele adicionou os seguintes recursos.

- Adicionadas novas APIs referentes ao padrão de recorrência em compromissos e mensagens que são solicitações de reunião.
- Foi modificada a propriedade item.from também estar disponível no modo Redação.
- Adicionado suporte para eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged.

### <a name="change-log"></a>Log de mudanças

- Adicionado o [From](/javascript/api/outlook_1_7/office.from): adiciona um novo objeto que fornece um método para obter o valor from.
- Adicionado o [Organizer](/javascript/api/outlook_1_7/office.organizer): adiciona um novo objeto que fornece um método para obter o valor organizer.
- Adicionado o [Recurrence](/javascript/api/outlook_1_7/office.recurrence): adiciona um novo objeto que fornece métodos para obter e definir o padrão de recorrência de compromissos, mas obtém apenas o padrão de recorrência de mensagens de solicitações de reunião.
- Adicionado o [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): adiciona um novo objeto que representa a configuração de fuso horário do padrão de recorrência.
- Adicionado o [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): adiciona um novo objeto que fornece métodos para obter e definir as datas e horas de compromissos em uma série recorrente e obter as datas e horas de solicitações de reunião em uma série recorrente.
- Adicionado o [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): adiciona um novo método que adiciona um manipulador de eventos para um evento com suporte.
- Modificado o [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): modifica para obter o valor from no modo Redação.
- Modificado o [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer): modifica para obter o valor organizer no modo Redação.
- Adicionado o [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): adiciona uma nova propriedade que obtém ou define um objeto que fornece métodos de gerenciamento do padrão de recorrência de um item de compromisso. Essa propriedade também pode ser usada para obter o padrão de recorrência de um item de solicitação de reunião.
- Adicionado o [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): adiciona um novo método que remove um manipulador de eventos.
- Adicionado o [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): adiciona uma nova propriedade que obtém a ID da série à qual uma ocorrência pertence.
- Adicionado o [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): adiciona uma nova enumeração que especifica o dia da semana ou o tipo de dia.
- Adicionado o [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): adiciona uma nova enumeração que especifica o mês.
- Adicionado o [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): adiciona uma nova enumeração que especifica o fuso horário aplicado à recorrência.
- Adicionado o [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): adiciona uma nova enumeração que especifica o tipo de recorrência.
- Adicionado o [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): adiciona uma nova enumeração que especifica a semana do mês.
- Modificado o [Office.EventType](/javascript/api/office/office.eventtype): modifica para dar suporte a eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged por meio da adição de entradas `RecurrenceChanged`, `RecipientsChanged` e `AppointmentTimeChanged`, respectivamente.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)