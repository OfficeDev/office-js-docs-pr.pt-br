---
title: Obter e definir uma recorrência em um suplemento do Outlook
description: Este tópico mostra como usar a API JavaScript do Office para obter e definir várias propriedades de recorrência de um item em um suplemento do Outlook.
ms.date: 08/18/2020
ms.localizationpriority: medium
ms.openlocfilehash: 0b211e72304e22874f847f2231e3a800efaceb4d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151858"
---
# <a name="get-and-set-recurrence"></a>Obter e definir uma recorrência

Às vezes, você precisa criar e atualizar um compromisso recorrente, como uma reunião de status semanal para um projeto da equipe ou um lembrete anual de aniversário. Você pode usar Office API JavaScript para gerenciar os padrões de recorrência de uma série de compromissos no seu complemento.

> [!NOTE]
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.7. Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="available-recurrence-patterns"></a>Padrões de recorrência disponíveis

Para configurar o padrão de recorrência, você precisa combinar o [tipo de recorrência](/javascript/api/outlook/office.mailboxenums.recurrencetype) e as [propriedades da recorrência](/javascript/api/outlook/office.recurrenceproperties) aplicáveis (se houver).

**Tabela 1. Tipos de recorrência e as propriedades aplicáveis delas**

|Tipo de recorrência|Propriedades válidas das recorrências|Uso|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|Um compromisso ocorre a cada *interval* (intervalo) de dias. Exemplo: um compromisso ocorre a cada **_2_** dias.|
|`weekday`|Nenhum.|Um compromisso ocorre todos os dias úteis.|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|– Ocorre um compromisso no dia *dayOfMonth* (diaDoMês) a cada *interval* (intervalo) de meses. Exemplo: um compromisso ocorre no dia **_5_** a cada **_4_** meses.<br/><br/>– Ocorre um compromisso na *weekNumber* (númeroDaSemana) do *dayOfWeek* (diaDoMês) a cada *interval* (intervalo) de meses. Exemplo: um compromisso ocorre na **_terceira_** **_quinta-feira_** a cada **_2_** meses.|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|Ocorre um compromisso nos *days* (dias) a cada *interval* (intervalo) de semanas. Exemplo: um compromisso ocorre na **_terça-feira_ e na _quinta-feira_** a cada **_2_** semanas.|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|– Ocorre um compromisso no dia *dayOfMonth* (diaDoMês) do *month* (mês) a cada *interval* (intervalo) de anos. Exemplo: um compromisso ocorre no dia **_7_** de **_setembro_** a cada **_4_** anos.<br/><br/>– Ocorre um compromisso na *weekNumber* (númeroDaSemana) *dayOfWeek* (diaDaSemana) do *month* (mês) a cada *interval* (intervalo) de anos. Exemplo: um compromisso ocorre na **_primeira_** **_quinta-feira_** de **_setembro_** a cada **_2_** anos.|

> [!NOTE]
> Você também pode usar a propriedade [`firstDayOfWeek`][firstDayOfWeek link] com o tipo de recorrência `weekly`. O dia especificado iniciará a lista de dias exibidos na caixa de diálogo de recorrência.

## <a name="access-recurrence"></a>Acessar a recorrência

Como você acessa o padrão de recorrência e o que pode fazer com ele depende de você ser o organizador ou um participante do compromisso.

**Tabela 2. Estados do compromisso aplicáveis**

|Estado do compromisso|A recorrência é editável?|A recorrência é visível?|
|---|---|---|
|Organizador de compromisso – redigir a série|Sim ( [`setAsync`][setAsync link] )|Sim ( [`getAsync`][getAsync link] )|
|Organizador de compromisso – redigir a instância|Não (`setAsync` retorna um erro)|Sim ( [`getAsync`][getAsync link] )|
|Participante de compromisso – ler a série|Não (`setAsync` não está disponível)|Sim ( [`item.recurrence`][item.recurrence link] )|
|Participante de compromisso – ler a instância|Não (`setAsync` não está disponível)|Sim ( [`item.recurrence`][item.recurrence link] )|
|Solicitação de reunião – ler a série|Não (`setAsync` não está disponível)|Sim ( [`item.recurrence`][item.recurrence link] )|
|Solicitação de reunião – ler a instância|Não (`setAsync` não está disponível)|Sim ( [`item.recurrence`][item.recurrence link] )|

## <a name="set-recurrence-as-the-organizer"></a>Definir recorrência como o organizador

Com o padrão de recorrência, também é necessário determinar os horários e as datas de início e término da série de compromissos. O objeto [`SeriesTime`][SeriesTime link] é usado para gerenciar essas informações.

O organizador de compromisso só pode especificar o padrão de recorrência para uma série de compromissos no modo de redação. No exemplo a seguir, a série de compromissos está definida para ocorrer das 10:30 às 11:00 toda terça-feira e quinta-feira durante o período de 2 de novembro de 2019 a 2 de dezembro de 2019.

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## <a name="change-recurrence-as-the-organizer"></a>Alterar a recorrência como organizador

No exemplo a seguir, no modo de redação, o organizador do compromisso obtém o objeto de recorrência de uma série de compromissos dada a série ou uma instância dessa série e define uma nova duração de recorrência.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var recurrencePattern = asyncResult.value;
  recurrencePattern.seriesTime.setDuration(60);
  Office.context.mailbox.item.recurrence.setAsync(recurrencePattern, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("failed");
      return;
    }

    console.log("success");
  });
}
```

## <a name="get-recurrence"></a>Obter recorrência

### <a name="get-recurrence-as-the-organizer"></a>Obter recorrência como o organizador

No exemplo a seguir, no modo de redação, o organizador de compromisso acessa o objeto de recorrência de uma série de compromissos relacionados à série ou a uma instância daquela série.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

O exemplo a seguir mostra os resultados da chamada `getAsync` que recupera a recorrência de uma série.

> [!NOTE]
> Neste exemplo, `seriesTimeObject` é um espaço reservado para o JSON que representa a propriedade `recurrence.seriesTime`. Você deve usar os métodos [`SeriesTime`][SeriesTime link] para acessar as propriedades de data e hora da recorrência.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-recurrence-as-an-attendee"></a>Obter recorrência como um participante

No exemplo a seguir, o participante do compromisso pode acessar o objeto de recorrência de uma série de compromissos relacionados à série, a uma instância daquela série ou a uma solicitação de reunião.

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

O exemplo a seguir mostra o valor da propriedade `item.recurrence` para uma série de compromissos.

> [!NOTE]
> Neste exemplo, `seriesTimeObject` é um espaço reservado para o JSON que representa a propriedade `recurrence.seriesTime`. Você deve usar os métodos [`SeriesTime`][SeriesTime link] para acessar as propriedades de data e hora da recorrência.

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-the-recurrence-details"></a>Obter os detalhes de recorrência

Depois que você recuperou o objeto de recorrência (seja do retorno de chamada de `getAsync` ou de `item.recurrence`), é possível obter as propriedades específicas da recorrência. Por exemplo, você pode usar os horários e as datas de início e término da série usando os [métodos][SeriesTime link] na propriedade `recurrence.seriesTime`.

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a>Confira também

[Evento RecurrenceChanged](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getAsync_options__callback_
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setAsync_recurrencePattern__options__callback_

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayOfMonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayOfWeek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstDayOfWeek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weekNumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
