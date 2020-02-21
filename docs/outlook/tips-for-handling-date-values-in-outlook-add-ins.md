---
title: Dicas para lidar com valores de data em Suplementos do Outlook
description: A API JavaScript para Office usa o objeto JavaScript Date para a maioria dos processos de armazenamento e recuperação de datas e horas.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5718839ebda433df6fb14886da34d734f81eb5f2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165695"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Dicas para lidar com valores de data em suplementos do Outlook

A API JavaScript para Office usa o objeto JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) para a maioria dos processos de armazenamento e recuperação de datas e horas. 

Esse objeto **Date** fornece métodos como [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) e [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), que retornam o valor de data ou hora selecionado de acordo com o UTC (Tempo Universal Coordenado).

O objeto **Date** também fornece outros métodos como [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) e [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), que retornam a data ou a hora solicitada de acordo com a "hora local".

O conceito de "hora local" é basicamente determinado pelo navegador e pelo sistema operacional no computador cliente. Por exemplo, na maioria dos navegadores em execução em um computador cliente baseado no Windows, uma chamada JavaScript para **getDate** retorna uma data com base no fuso horário definido no Windows do computador cliente.

O exemplo a seguir cria um objeto **Date** `myLocalDate` na hora local e chama **toUTCString** para converter essa data em uma cadeia de caracteres de data em UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Embora você possa usar o objeto JavaScript **Date** para obter um valor de data ou hora baseado em UTC ou no fuso horário do computador cliente, o objeto **Date** é limitado em um aspecto: ele não oferece métodos para retornar um valor de data ou hora para qualquer outro fuso horário específico. Por exemplo, se seu computador cliente estiver definido para EST (hora oficial do leste dos EUA), não existe método **Date** que permita que você obtenha o valor de hora diferente de EST ou UTC, como PST (hora oficial do Pacífico).


## <a name="date-related-features-for-outlook-add-ins"></a>Recursos relacionados a data para suplementos do Outlook

A limitação JavaScript mencionada tem uma implicação para você, quando você usa a API JavaScript para Office para lidar com valores de data ou hora em suplementos do Outlook que são executados em um cliente avançado do Outlook e no Outlook na Web ou em dispositivos móveis.


### <a name="time-zones-for-outlook-clients"></a>Fusos horários para clientes do Outlook

Para maior clareza, vamos definir os fusos horários em questão.

|**Fuso horário**|**Descrição**|
|:-----|:-----|
|Fuso horário do computador cliente|Isso é definido no sistema operacional do computador cliente. A maioria dos navegadores usa o fuso horário do computador cliente para exibir os valores de data ou hora do objeto JavaScript **Date**.<br/><br/>Um cliente avançado do Outlook usa esse fuso horário para exibir os valores de data ou hora na interface do usuário. <br/><br/>Por exemplo, em um computador cliente executando o Windows, o Outlook usa o fuso horário definido no Windows como o fuso horário local. No Mac, se o usuário alterar o fuso horário no computador cliente, o Outlook no Mac solicitará que o usuário atualize o fuso horário no Outlook também.|
|Fuso horário do EAC (Centro de Administração do Exchange)|O usuário define esse valor de fuso horário (e o idioma preferido) quando faz logon no Outlook na Web ou em dispositivos móveis pela primeira vez. <br/><br/>O Outlook na Web e dispositivos móveis usam esse fuso horário para exibir valores de data ou hora na interface do usuário.|

Como um cliente avançado do Outlook usa o fuso horário do computador cliente e a interface do usuário do Outlook na Web e dispositivos móveis usa o fuso horário do Eat, a hora local para o mesmo suplemento instalada para a mesma caixa de correio pode ser diferente ao executar em um clie avançado do Outlook NT e no Outlook na Web ou dispositivos móveis. Como desenvolvedor de suplementos do Outlook, você deve fornecer valores de data de entrada e saída de forma que sejam sempre consistentes com o fuso horário que o usuário espera no cliente correspondente.


### <a name="date-related-api"></a>API relacionada à data

A seguir estão as propriedades e métodos da API JavaScript para Office que dão suporte a recursos relacionados a datas.

**Membro da API**|**Representação de fuso horário**|**Exemplo em um cliente avançado do Outlook**|**Exemplo no Outlook na Web ou em dispositivos móveis**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|Em um cliente avançado do Outlook, essa propriedade retorna o fuso horário do computador cliente. No Outlook na Web e dispositivos móveis, essa propriedade retorna o fuso horário da Eat. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Cada uma dessas propriedades retorna um objeto JavaScript **Date**. Esse valor de **Data** está correto, conforme mostrado no exemplo a seguir: `myUTCDate` tem o mesmo valor em um cliente avançado do Outlook, Outlook na Web e dispositivos móveis.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>No entanto `myDate.getDate` , chamar retorna um valor de data no fuso horário do computador cliente, que é consistente com o fuso horário usado para exibir os valores de data e hora na interface de cliente avançado do Outlook, mas pode ser diferente do fuso horário da Eat que o Outlook na Web e dispositivos móveis usam em sua interface do usuário.|Se o item é criado às 9h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` é retornado às 6h EST.|Se a hora de criação do item for às 9h UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` é retornado às 6h EST.<br/><br/>Observe que se você quer exibir a hora de criação ou de alteração na interface do usuário, convém primeiro converter a hora em PST para ficar consistente com o resto da interface do usuário.
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Cada um dos parâmetros _Start_ e _End_ requer um objeto JavaScript **Date**. Os argumentos devem ser corrigidos por UTC, independentemente do fuso horário usado na interface do usuário de um cliente avançado do Outlook ou do Outlook na Web ou dispositivos móveis.|Se as horas de início e de término para o formulário de compromisso são 9h UTC e 11h UTC, você deve fazer com que os argumentos `start` e `end` estejam corretos em relação à UTC, o que significa que :<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>|Se as horas de início e de término para o formulário de compromisso são 9h UTC e 11h UTC, você deve fazer com que os argumentos `start` e `end` estejam corretos em relação à UTC, o que significa que :<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>Métodos auxiliares para cenários de data


Conforme descrito nas seções anteriores, como a "hora local" para um usuário no Outlook na Web ou dispositivos móveis pode ser diferente em um cliente avançado do Outlook, mas o objeto de **Data** JavaScript oferece suporte à conversão somente no fuso horário do computador cliente ou UTC, a API JavaScript para Office fornece dois métodos auxiliares: [Office. Context. Mailbox. convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) e [Office. Context. Mailbox. convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Esses métodos auxiliares cuidam de qualquer necessidade de lidar com a data ou a hora de forma diferente para os dois cenários relacionados a datas a seguir, em um cliente avançado do Outlook, no Outlook na Web e em dispositivos móveis, reforçando "somente uma vez" para clientes diferentes do seu suplemento.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Cenário A: exibir a criação de item ou a hora da alteração

Se você estiver exibindo a hora de criação do item (**Item.dateTimeCreated**) ou a hora da modificação (**Item.dateTimeModified**) na interface do usuário, primeiro use **convertToLocalClientTime** para converter o objeto **Date** fornecido por essas propriedades a fim de obter uma representação de dicionário no horário local apropriado. Em seguida, exiba as partes da data do dicionário. A seguir, um exemplo desse cenário:


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Observe que **convertToLocalClientTime** se encarrega da diferença entre um cliente avançado do Outlook e o Outlook na Web ou dispositivos móveis:


- Se **convertToLocalClientTime** detecta que o host atual é um cliente avançado, o método converte a representação **Date** em uma representação de dicionário no mesmo fuso horário do computador cliente, consistente com o resto da interface do usuário do cliente avançado.
    
- Se o **convertToLocalClientTime** detectar que o host atual é o Outlook na Web ou dispositivos móveis, o método converterá a representação de **Data** correta para UTC em um formato de dicionário no fuso horário da Eat, consistente com o restante da interface de usuário do Outlook na Web ou dispositivos móveis.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Cenário B: exibir datas de início e de término em um formulário de novo compromisso

Se você está obtendo como entrada diferentes partes de um valor de data e hora representadas na hora local e quer fornecer esse valor de entrada de dicionário como uma hora de início ou de término em um formulário de compromisso, primeiro use o método auxiliar **convertToUtcClientTime** para converter o valor do dicionário em um objeto **Date** com correção para UTC.

No exemplo a seguir, supinha que `myLocalDictionaryStartDate` e `myLocalDictionaryEndDate` são valores de data e hora em formato de dicionário que você obteve do usuário. Esses valores se baseiam no horário local, dependendo do aplicativo host.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Os valores resultantes, `myUTCCorrectStartDate` e `myUTCCorrectEndDate`, são corrigidos para UTC. Passe esses objetos **Date** como argumentos para os parâmetros _Start_ e _End_ do método **Mailbox.displayNewAppointmentForm** para exibir o formulário do novo compromisso.

Observe que **convertToUtcClientTime** se encarrega da diferença entre um cliente avançado do Outlook e o Outlook na Web ou dispositivos móveis:


- Se **convertToUtcClientTime** detecta que o host atual é um cliente avançado do Outlook, o método simplesmente converte a representação de dicionário em um objeto **Date**. Esse objeto **Date** é corrigido para UTC, conforme o esperado por **displayNewAppointmentForm**.
    
- Se o **convertToUtcClientTime** detectar que o host atual é o Outlook na Web ou dispositivos móveis, o método converte o formato de dicionário dos valores de data e hora expressos no fuso horário da Eat em um objeto **Date** . Esse objeto **Date** é corrigido para UTC, conforme o esperado por **displayNewAppointmentForm**.
    

## <a name="see-also"></a>Confira também

- [Implantar e instalar suplementos do Outlook para teste](testing-and-tips.md)
    


