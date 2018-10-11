# <a name="outlook-add-in-api-requirement-set-17"></a>Conjunto de requisitos de API versão 1.7 para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript para Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

## <a name="whats-new-in-17"></a>Novidades na versão 1.7?

O conjunto de requisitos versão 1.7 inclui todos os recursos do [conjunto de requisitos versão 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Ele adicionou os recursos a seguir.

- Novas APIs adicionadas com relação ao padrão de recorrência em compromissos e mensagens que são solicitações de reunião.
- A propriedade item.from foi modificada para também ficar disponível no modo Redigir.
- Foi adicionado suporte para os eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged.

### <a name="change-log"></a>Log de alterações

- [From](/javascript/api/outlook_1_7/office.from) adicionado: adiciona um novo objeto que fornece um método para obter o valor from
- [Organizer](/javascript/api/outlook_1_7/office.organizer) adicionado: adiciona um novo objeto que fornece um método para obter o valor organizer.
- [Recurrence](/javascript/api/outlook_1_7/office.recurrence) adicionado: adiciona um novo objeto que fornece métodos para obter e definir o padrão de recorrência de compromissos, mas apenas obtém o padrão de recorrência de mensagens que são solicitações de reunião.
- [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone) adicionado: adiciona um novo objeto que representa a configuração de fuso horário do padrão de recorrência.
-  [SeriesTime](/javascript/api/outlook_1_7/office.seriestime) adicionado: adiciona um novo objeto que fornece os métodos para obter e definir as datas e horários de compromissos em uma série recorrente e para obter as datas e horários de solicitações de reunião em uma série recorrente.
- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) adicionado: adiciona um novo método que adiciona um manipulador de eventos a um evento com suporte.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) modificado: modifica para o valor from no modo Redigir.
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) modificado - modifica para obter o valor organizer no modo Redigir.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) adicionado: adiciona uma nova propriedade que obtém ou define um objeto que fornece os métodos para gerenciar o padrão de recorrência de um item de compromisso. Essa propriedade também pode ser usada para obter o padrão de recorrência de um item de solicitação de reunião.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) adicionado: adiciona um novo método que remove um manipulador de eventos.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) adicionado: adiciona uma nova propriedade que obtém a id da série a qual uma ocorrência pertence.
- [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days) adicionado: adiciona um nova enumeração que especifica o dia da semana ou o tipo de dia.
- [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month) adicionado: adiciona um nova enumeração que especifica o mês.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone) adicionado: adiciona um nova enumeração que especifica o fuso horário aplicado à recorrência.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype) adicionado: adiciona um nova enumeração que especifica o tipo de recorrência.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber) adicionado: adiciona um nova enumeração que especifica a semana do mês.
- [Office.EventType](/javascript/api/office/office.eventtype) modificado: modifica para dar suporte aos eventos RecurrenceChanged, RecipientsChanged e AppointmentTimeChanged por meio da adição das entradas `RecurrenceChanged`, `RecipientsChanged` e `AppointmentTimeChanged`, respectivamente.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)