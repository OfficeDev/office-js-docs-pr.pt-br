---
title: Dicas para lidar com valores de data em Suplementos do Outlook
description: A API JavaScript do Office usa o objeto Data do JavaScript para a maioria do armazenamento e recuperação de datas e horas.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49de8db712400e006dc919e9ad62ae6cbaaa11cf
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713074"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Dicas para lidar com valores de data em suplementos do Outlook

A API JavaScript do Office usa o objeto [Data](https://www.w3schools.com/jsref/jsref_obj_date.asp) do JavaScript para a maioria do armazenamento e recuperação de datas e horas.

`Date` Esse objeto fornece métodos como [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) e [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), que retornam o valor de data ou hora solicitado de acordo com a hora UTC (Tempo Coordenado Universal).

O `Date` objeto também fornece outros métodos, como [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) e [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), que retornam a data ou hora solicitada de acordo com a "hora local".

O conceito de "hora local" é basicamente determinado pelo navegador e pelo sistema operacional no computador cliente. Por exemplo, na maioria dos navegadores em execução em um computador cliente baseado no Windows, uma chamada JavaScript `getDate`para retorna uma data com base no fuso horário definido no Windows no computador cliente.

O exemplo a seguir cria um `Date` objeto `myLocalDate` na hora local e `toUTCString` chama para converter essa data em uma cadeia de caracteres de data em UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Embora você possa usar o objeto JavaScript `Date` para obter um valor de data ou hora com base em UTC ou no fuso horário do computador cliente, o objeto **Date** é limitado em um aspecto – ele não fornece métodos para retornar um valor de data ou hora para qualquer outro fuso horário específico. Por exemplo, se o computador cliente estiver definido para estar no Horário Padrão do Leste (EST), `Date` não haverá nenhum método que permita que você obtenha o valor de hora diferente de EST ou UTC, como Hora Padrão do Pacífico (PST).

## <a name="date-related-features-for-outlook-add-ins"></a>Recursos relacionados a data para suplementos do Outlook

A limitação de JavaScript mencionada anteriormente tem uma implicação para você, quando você usa a API JavaScript do Office para lidar com valores de data ou hora em suplementos do Outlook que são executados em um cliente avançado do Outlook e em Outlook na Web ou dispositivos móveis.

### <a name="time-zones-for-outlook-clients"></a>Fusos horários para clientes do Outlook

Para maior clareza, vamos definir os fusos horários em questão.

|**Fuso horário**|**Descrição**|
|:-----|:-----|
|Fuso horário do computador cliente|Isso é definido no sistema operacional do computador cliente. A maioria dos navegadores usa o fuso horário do computador cliente para exibir valores de data ou hora do objeto JavaScript `Date` .<br/><br/>Um cliente avançado do Outlook usa esse fuso horário para exibir os valores de data ou hora na interface do usuário. <br/><br/>Por exemplo, em um computador cliente executando o Windows, o Outlook usa o fuso horário definido no Windows como o fuso horário local. No Mac, se o usuário alterar o fuso horário no computador cliente, o Outlook no Mac solicitará que o usuário atualize o fuso horário no Outlook também.|
|Fuso horário do EAC (Centro de Administração do Exchange)|O usuário define esse valor de fuso horário (e o idioma preferencial) quando faz logon em Outlook na Web ou dispositivos móveis pela primeira vez. <br/><br/>Outlook na Web e dispositivos móveis usam esse fuso horário para exibir valores de data ou hora na interface do usuário.|

Como um cliente avançado do Outlook usa o fuso horário do computador cliente e a interface do usuário do Outlook na Web e dispositivos móveis usa o fuso horário do EAC, a hora local para o mesmo suplemento instalado para a mesma caixa de correio pode ser diferente ao ser executada em um cliente avançado do Outlook e em dispositivos Outlook na Web ou móveis. Como desenvolvedor de suplementos do Outlook, você deve fornecer valores de data de entrada e saída de forma que sejam sempre consistentes com o fuso horário que o usuário espera no cliente correspondente.

### <a name="date-related-api"></a>API relacionada à data

A seguir estão as propriedades e os métodos na API JavaScript do Office que dão suporte a recursos relacionados à data.

|Membro da API|Representação de fuso horário|Exemplo em um cliente avançado do Outlook|Exemplo em Outlook na Web ou dispositivos móveis|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|Em um cliente avançado do Outlook, essa propriedade retorna o fuso horário do computador cliente. Em Outlook na Web dispositivos móveis, essa propriedade retorna o fuso horário do EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Cada uma dessas propriedades retorna um objeto JavaScript `Date` . Esse `Date` valor é utc-correto, conforme mostrado no exemplo a seguir - `myUTCDate` tem o mesmo valor em um cliente avançado do Outlook, Outlook na Web dispositivos móveis.<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>No entanto, `myDate.getDate` chamar retorna um valor de data no fuso horário do computador cliente, que é consistente com o fuso horário usado para exibir valores de data e hora na interface avançada do cliente do Outlook, mas pode ser diferente do fuso horário do EAC que Outlook na Web e dispositivos móveis usam em sua interface do usuário.|Se o item for criado às 9h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` é retornado às 6h EST.|Se a hora de criação do item for 9:00 UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` é retornado às 6h EST.<br/><br/>Observe que se você quer exibir a hora de criação ou de alteração na interface do usuário, convém primeiro converter a hora em PST para ficar consistente com o resto da interface do usuário.|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|Cada um dos _parâmetros Start_ _e End_ requer um objeto JavaScript `Date` . Os argumentos devem ser corretos em UTC, independentemente do fuso horário usado na interface do usuário de um cliente avançado do Outlook ou Outlook na Web ou dispositivos móveis.|Se as horas de início e término do formulário de compromisso forem 9:00 UTC e 11:00 UTC, `start` `end` você deverá garantir que os argumentos e os argumentos estejam corretos em UTC, o que significa:<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>|Se as horas de início e término do formulário de compromisso forem 9:00 UTC e 11:00 UTC, `start` `end` você deverá garantir que os argumentos e os argumentos estejam corretos em UTC, o que significa:<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Métodos auxiliares para cenários de data

Conforme descrito nas seções anteriores, como a "hora local" de um usuário no Outlook na Web ou em dispositivos móveis pode ser diferente em um cliente avançado do Outlook, mas o objeto Data **javaScript dá** suporte à conversão apenas para o fuso horário do computador cliente ou UTC, a API JavaScript do Office fornece dois métodos auxiliares: [Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) e [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

Esses métodos auxiliares cuidam de qualquer necessidade de lidar com data ou hora de forma diferente para os dois cenários relacionados a datas a seguir, em um cliente avançado do Outlook, Outlook na Web e dispositivos móveis, reforçando assim "write-once" para diferentes clientes do seu suplemento.

### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Cenário A: exibir a criação de item ou a hora da alteração

Se você estiver exibindo a hora de criação do item (`Item.dateTimeCreated`) ou o tempo de modificação (`Item.dateTimeModified`na interface do usuário, primeiro use `convertToLocalClientTime` `Date` para converter o objeto fornecido por essas propriedades para obter uma representação de dicionário na hora local apropriada. Em seguida, exiba as partes da data do dicionário. A seguir está um exemplo desse cenário.

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Observe que `convertToLocalClientTime` cuida da diferença entre um cliente avançado do Outlook e Outlook na Web ou dispositivos móveis:

- Se `convertToLocalClientTime` detectar que o aplicativo atual é um cliente avançado, `Date` o método converterá a representação em uma representação de dicionário no mesmo fuso horário do computador cliente, consistente com o restante da interface do usuário do cliente avançado.

- `convertToLocalClientTime` Se detectar que o aplicativo atual é Outlook na Web ou dispositivos móveis, o método converterá a representação utc-correcta `Date` em um formato de dicionário no fuso horário EAC, consistente com o restante da interface do usuário Outlook na Web ou dispositivos móveis.

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Cenário B: exibir datas de início e de término em um formulário de novo compromisso

Se você estiver obtendo como partes de entrada diferentes de um valor de data/hora representado na hora local e quiser fornecer esse valor de entrada de dicionário como uma hora de início ou de término em um formulário de compromisso, primeiro use `convertToUtcClientTime` o método auxiliar para converter o valor do dicionário em um objeto UTC-correct `Date` .

No exemplo a seguir, assuma que `myLocalDictionaryStartDate` e `myLocalDictionaryEndDate` são valores de data e hora em formato de dicionário que você obteve do usuário. Esses valores são baseados na hora local, dependendo da plataforma do cliente.

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Os valores resultantes, `myUTCCorrectStartDate` e `myUTCCorrectEndDate`, são corrigidos para UTC. Em seguida, passe esses `Date` objetos como argumentos para os parâmetros _Start_ e _End_ do `Mailbox.displayNewAppointmentForm` método para exibir o novo formulário de compromisso.

Observe que `convertToUtcClientTime` cuida da diferença entre um cliente avançado do Outlook e Outlook na Web ou dispositivos móveis:

- Se `convertToUtcClientTime` detectar que o aplicativo atual é um cliente avançado do Outlook, o método simplesmente converterá a representação do dicionário em um `Date` objeto. Este `Date` objeto está correto em UTC, conforme o esperado por `displayNewAppointmentForm`.

- Se `convertToUtcClientTime` detectar que o aplicativo atual Outlook na Web dispositivos móveis, o método converterá o formato de dicionário dos valores de data e hora expressos no fuso horário EAC em um `Date` objeto. Este `Date` objeto está correto em UTC, conforme o esperado por `displayNewAppointmentForm`.

## <a name="see-also"></a>Confira também

- [Implantar e instalar suplementos do Outlook para teste](testing-and-tips.md)
