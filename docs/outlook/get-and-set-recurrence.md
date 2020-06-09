---
title: Obter e definir uma recorrência em um suplemento do Outlook
description: Este tópico mostra como usar a API JavaScript do Office para obter e definir várias propriedades de recorrência de um item em um suplemento do Outlook.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6a50ba5eab39145d8e50a5a888a6ed0900200bc4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44606451"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="2c92e-103">Obter e definir uma recorrência</span><span class="sxs-lookup"><span data-stu-id="2c92e-103">Get and set recurrence</span></span>

<span data-ttu-id="2c92e-104">Às vezes, você precisa criar e atualizar um compromisso recorrente, como uma reunião de status semanal para um projeto da equipe ou um lembrete anual de aniversário.</span><span class="sxs-lookup"><span data-stu-id="2c92e-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="2c92e-105">Você pode usar a API JavaScript do Office para gerenciar os padrões de recorrência de uma série de compromissos no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="2c92e-105">You can use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2c92e-106">O suporte para esse recurso foi introduzido no conjunto de requisitos 1,7.</span><span class="sxs-lookup"><span data-stu-id="2c92e-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="2c92e-107">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="2c92e-108">Padrões de recorrência disponíveis</span><span class="sxs-lookup"><span data-stu-id="2c92e-108">Available recurrence patterns</span></span>

<span data-ttu-id="2c92e-109">Para configurar o padrão de recorrência, você precisa combinar o [tipo de recorrência](/javascript/api/outlook/office.mailboxenums.recurrencetype) e as [propriedades da recorrência](/javascript/api/outlook/office.recurrenceproperties) aplicáveis (se houver).</span><span class="sxs-lookup"><span data-stu-id="2c92e-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="2c92e-110">**Tabela 1. Tipos de recorrência e as propriedades aplicáveis delas**</span><span class="sxs-lookup"><span data-stu-id="2c92e-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="2c92e-111">Tipo de recorrência</span><span class="sxs-lookup"><span data-stu-id="2c92e-111">Recurrence type</span></span>|<span data-ttu-id="2c92e-112">Propriedades válidas das recorrências</span><span class="sxs-lookup"><span data-stu-id="2c92e-112">Valid recurrence properties</span></span>|<span data-ttu-id="2c92e-113">Uso</span><span class="sxs-lookup"><span data-stu-id="2c92e-113">Usage</span></span>|
|---|---|---|
|`daily`|- [`interval`][interval link]|<span data-ttu-id="2c92e-114">Um compromisso ocorre a cada *interval* (intervalo) de dias.</span><span class="sxs-lookup"><span data-stu-id="2c92e-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="2c92e-115">Exemplo: um compromisso ocorre a cada **_2_** dias.</span><span class="sxs-lookup"><span data-stu-id="2c92e-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="2c92e-116">Nenhum.</span><span class="sxs-lookup"><span data-stu-id="2c92e-116">None.</span></span>|<span data-ttu-id="2c92e-117">Um compromisso ocorre todos os dias úteis.</span><span class="sxs-lookup"><span data-stu-id="2c92e-117">An appointment occurs every weekday.</span></span>|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|<span data-ttu-id="2c92e-118">– Ocorre um compromisso no dia *dayOfMonth* (diaDoMês) a cada *interval* (intervalo) de meses.</span><span class="sxs-lookup"><span data-stu-id="2c92e-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="2c92e-119">Exemplo: um compromisso ocorre no dia **_5_** a cada **_4_** meses.</span><span class="sxs-lookup"><span data-stu-id="2c92e-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="2c92e-120">– Ocorre um compromisso na *weekNumber* (númeroDaSemana) do *dayOfWeek* (diaDoMês) a cada *interval* (intervalo) de meses.</span><span class="sxs-lookup"><span data-stu-id="2c92e-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="2c92e-121">Exemplo: um compromisso ocorre na **_terceira_** **_quinta-feira_** a cada **_2_** meses.</span><span class="sxs-lookup"><span data-stu-id="2c92e-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|<span data-ttu-id="2c92e-122">Ocorre um compromisso nos *days* (dias) a cada *interval* (intervalo) de semanas.</span><span class="sxs-lookup"><span data-stu-id="2c92e-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="2c92e-123">Exemplo: um compromisso ocorre na **_terça-feira_ e na _quinta-feira_** a cada **_2_** semanas.</span><span class="sxs-lookup"><span data-stu-id="2c92e-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|<span data-ttu-id="2c92e-124">– Ocorre um compromisso no dia *dayOfMonth* (diaDoMês) do *month* (mês) a cada *interval* (intervalo) de anos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="2c92e-125">Exemplo: um compromisso ocorre no dia **_7_** de **_setembro_** a cada **_4_** anos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="2c92e-126">– Ocorre um compromisso na *weekNumber* (númeroDaSemana) *dayOfWeek* (diaDaSemana) do *month* (mês) a cada *interval* (intervalo) de anos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="2c92e-127">Exemplo: um compromisso ocorre na **_primeira_** **_quinta-feira_** de **_setembro_** a cada **_2_** anos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="2c92e-128">Você também pode usar a propriedade [`firstDayOfWeek`][firstDayOfWeek link] com o tipo de recorrência `weekly`.</span><span class="sxs-lookup"><span data-stu-id="2c92e-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="2c92e-129">O dia especificado iniciará a lista de dias exibidos na caixa de diálogo de recorrência.</span><span class="sxs-lookup"><span data-stu-id="2c92e-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="2c92e-130">Acessar a recorrência</span><span class="sxs-lookup"><span data-stu-id="2c92e-130">Access recurrence</span></span>

<span data-ttu-id="2c92e-131">Como você acessa o padrão de recorrência e o que pode fazer com ele depende de você ser o organizador ou um participante do compromisso.</span><span class="sxs-lookup"><span data-stu-id="2c92e-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="2c92e-132">**Tabela 2. Estados do compromisso aplicáveis**</span><span class="sxs-lookup"><span data-stu-id="2c92e-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="2c92e-133">Estado do compromisso</span><span class="sxs-lookup"><span data-stu-id="2c92e-133">Appointment state</span></span>|<span data-ttu-id="2c92e-134">A recorrência é editável?</span><span class="sxs-lookup"><span data-stu-id="2c92e-134">Is recurrence editable?</span></span>|<span data-ttu-id="2c92e-135">A recorrência é visível?</span><span class="sxs-lookup"><span data-stu-id="2c92e-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="2c92e-136">Organizador de compromisso – redigir a série</span><span class="sxs-lookup"><span data-stu-id="2c92e-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="2c92e-137">Sim ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="2c92e-138">Sim ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="2c92e-139">Organizador de compromisso – redigir a instância</span><span class="sxs-lookup"><span data-stu-id="2c92e-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="2c92e-140">Não (`setAsync` retorna um erro)</span><span class="sxs-lookup"><span data-stu-id="2c92e-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="2c92e-141">Sim ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="2c92e-142">Participante de compromisso – ler a série</span><span class="sxs-lookup"><span data-stu-id="2c92e-142">Appointment attendee - read series</span></span>|<span data-ttu-id="2c92e-143">Não (`setAsync` não está disponível)</span><span class="sxs-lookup"><span data-stu-id="2c92e-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="2c92e-144">Sim ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2c92e-145">Participante de compromisso – ler a instância</span><span class="sxs-lookup"><span data-stu-id="2c92e-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="2c92e-146">Não (`setAsync` não está disponível)</span><span class="sxs-lookup"><span data-stu-id="2c92e-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="2c92e-147">Sim ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2c92e-148">Solicitação de reunião – ler a série</span><span class="sxs-lookup"><span data-stu-id="2c92e-148">Meeting request - read series</span></span>|<span data-ttu-id="2c92e-149">Não (`setAsync` não está disponível)</span><span class="sxs-lookup"><span data-stu-id="2c92e-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="2c92e-150">Sim ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2c92e-151">Solicitação de reunião – ler a instância</span><span class="sxs-lookup"><span data-stu-id="2c92e-151">Meeting request - read instance</span></span>|<span data-ttu-id="2c92e-152">Não (`setAsync` não está disponível)</span><span class="sxs-lookup"><span data-stu-id="2c92e-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="2c92e-153">Sim ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2c92e-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="2c92e-154">Definir recorrência como o organizador</span><span class="sxs-lookup"><span data-stu-id="2c92e-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="2c92e-155">Com o padrão de recorrência, também é necessário determinar os horários e as datas de início e término da série de compromissos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="2c92e-156">O objeto [`SeriesTime`][SeriesTime link] é usado para gerenciar essas informações.</span><span class="sxs-lookup"><span data-stu-id="2c92e-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="2c92e-157">O organizador de compromisso só pode especificar o padrão de recorrência para uma série de compromissos no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="2c92e-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="2c92e-158">No exemplo a seguir, a série de compromissos está definida para ocorrer das 10:30 às 11:00 toda terça-feira e quinta-feira durante o período de 2 de novembro de 2019 a 2 de dezembro de 2019.</span><span class="sxs-lookup"><span data-stu-id="2c92e-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

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

## <a name="get-recurrence"></a><span data-ttu-id="2c92e-159">Obter recorrência</span><span class="sxs-lookup"><span data-stu-id="2c92e-159">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="2c92e-160">Obter recorrência como o organizador</span><span class="sxs-lookup"><span data-stu-id="2c92e-160">Get recurrence as the organizer</span></span>

<span data-ttu-id="2c92e-161">No exemplo a seguir, no modo de redação, o organizador de compromisso acessa o objeto de recorrência de uma série de compromissos relacionados à série ou a uma instância daquela série.</span><span class="sxs-lookup"><span data-stu-id="2c92e-161">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

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

<span data-ttu-id="2c92e-162">O exemplo a seguir mostra os resultados da chamada `getAsync` que recupera a recorrência de uma série.</span><span class="sxs-lookup"><span data-stu-id="2c92e-162">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="2c92e-163">Neste exemplo, `seriesTimeObject` é um espaço reservado para o JSON que representa a propriedade `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="2c92e-163">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="2c92e-164">Você deve usar os métodos [`SeriesTime`][SeriesTime link] para acessar as propriedades de data e hora da recorrência.</span><span class="sxs-lookup"><span data-stu-id="2c92e-164">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="2c92e-165">Obter recorrência como um participante</span><span class="sxs-lookup"><span data-stu-id="2c92e-165">Get recurrence as an attendee</span></span>

<span data-ttu-id="2c92e-166">No exemplo a seguir, o participante do compromisso pode acessar o objeto de recorrência de uma série de compromissos relacionados à série, a uma instância daquela série ou a uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="2c92e-166">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

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

<span data-ttu-id="2c92e-167">O exemplo a seguir mostra o valor da propriedade `item.recurrence` para uma série de compromissos.</span><span class="sxs-lookup"><span data-stu-id="2c92e-167">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="2c92e-168">Neste exemplo, `seriesTimeObject` é um espaço reservado para o JSON que representa a propriedade `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="2c92e-168">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="2c92e-169">Você deve usar os métodos [`SeriesTime`][SeriesTime link] para acessar as propriedades de data e hora da recorrência.</span><span class="sxs-lookup"><span data-stu-id="2c92e-169">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="2c92e-170">Obter os detalhes de recorrência</span><span class="sxs-lookup"><span data-stu-id="2c92e-170">Get the recurrence details</span></span>

<span data-ttu-id="2c92e-171">Depois que você recuperou o objeto de recorrência (seja do retorno de chamada de `getAsync` ou de `item.recurrence`), é possível obter as propriedades específicas da recorrência.</span><span class="sxs-lookup"><span data-stu-id="2c92e-171">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="2c92e-172">Por exemplo, você pode usar os horários e as datas de início e término da série usando os [métodos][SeriesTime link] na propriedade `recurrence.seriesTime`.</span><span class="sxs-lookup"><span data-stu-id="2c92e-172">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2c92e-173">Confira também</span><span class="sxs-lookup"><span data-stu-id="2c92e-173">See also</span></span>

[<span data-ttu-id="2c92e-174">Evento RecurrenceChanged</span><span class="sxs-lookup"><span data-stu-id="2c92e-174">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
