---
title: Regras de ativação para suplementos do Outlook
description: O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: b9baf3c813dcb1aefc6554e8e295d50045803dd9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165854"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Regras de ativação para suplementos contextuais do Outlook

O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento. Isso é verdadeiro para todos os suplementos que usam o esquema de manifesto 1.1. O usuário pode escolher o suplemento na interface de usuário do Outlook para iniciá-lo em relação ao item atual.

A figura a seguir mostra suplementos do Outlook ativados na barra de suplementos da mensagem que está no painel de leitura. 

![Barra de aplicativos mostrando aplicativos de email de leitura ativados](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>Especificar regras de ativação em um manifesto


Para que o Outlook ative um suplemento em condições específicas, especifique as regras de ativação no manifesto do suplemento usando um dos seguintes elementos **Rule**:

- [Elemento Rule (MailApp complexType)](../reference/manifest/rule.md) - especifica uma regra individual.
- [Elemento Rule (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - combina várias regras usando operações lógicas.
    

 > [!NOTE]
 > O elemento **Rule** que você usa para especificar uma regra individual é do tipo complexo [Rule](../reference/manifest/rule.md) abstrato. Cada um dos tipos de regra a seguir estende esse tipo complexo **Rule** abstrato. Portanto, ao especificar uma regra individual em um manifesto, é preciso usar o atributo [xsi:type](https://www.w3.org/TR/xmlschema-1/) para definir um dos tipos de regra a seguir. 
 > 
 > Por exemplo, a seguinte regra define uma regra [ItemIs](../reference/manifest/rule.md#itemis-rule): `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > O atributo **FormType** se aplica às regras de ativação na versão 1.1 do manifesto, mas não está definido na versão 1.0 do **VersionOverrides**. Portanto, não pode ser usado quando [ItemIs](../reference/manifest/rule.md#itemis-rule) é usado no nó **VersionOverrides**.

A tabela a seguir lista os tipos de regra disponíveis. Veja mais informações após a tabela e nos artigos especificados em [Criar suplementos do Outlook para formulários de leitura](read-scenario.md).

<br/>

|**Nome da regra**|**Formulários aplicáveis**|**Descrição**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|Ler, Redigir|Verifica se o item atual é do tipo especificado (compromisso ou mensagem). Pode também verificar a classe do item e o tipo de formulário e, opcionalmente, a classe de mensagem do item.|
|[ItemHasAttachment](#itemhasattachment-rule)|Leitura|Verifica se o item selecionado contém um anexo.|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|Leitura|Verifica se o item selecionado contém uma ou mais entidades conhecidas. Mais informações: [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|Leitura|Verifica se o endereço de email do remetente, o assunto e/ou o corpo do item selecionado contêm uma correspondência para uma expressão regular. Mais informações: [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](#rulecollection-rule)|Ler, Redigir|Combina uma coleção de regras para que você forme regras mais complexas.|

## <a name="itemis-rule"></a>Regra ItemIs

O tipo complexo **ItemIs** define uma regra que avalia **true** se o item atual coincidir com o tipo de item e, opcionalmente, a classe de mensagens do item, se estiver declarada na regra.

Especifique um dos tipos de item a seguir no atributo **ItemType** de uma regra **ItemIs**. Você pode especificar mais de uma regra **ItemIs** em um manifesto. O tipo simples ItemType define os tipos de itens do Outlook que dão suporte aos suplementos do Outlook.

<br/>

|**Valor**|**Descrição**|
|:-----|:-----|
|**Compromisso**|Especifica um item em um calendário do Outlook. Isso inclui um item de reunião que foi respondido e que tem um organizador e participantes, ou um compromisso que não tem um organizador ou participantes e é simplesmente um item no calendário. Isso corresponde à classe de mensagens IPM.Appointment no Outlook.|
|**Mensagem**|Especifica um dos seguintes itens recebidos normalmente na Caixa de Entrada: <ul><li><p>Uma mensagem de email. Isso corresponde à classe de mensagem IPM.Note no Outlook.</p></li><li><p>Uma solicitação de reunião, resposta ou cancelamento. Isso corresponde às seguintes classes de mensagem no Outlook:</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

O atributo **FormType** é usado para especificar o modo (leitura ou redação) no qual o suplemento deve ser ativado.


 > [!NOTE]
 > O atributo ItemIs **FormType** está definido no esquema v1.1 e versões posteriores, mas não no **VersionOverrides** v1.0. Não inclua o atributo **FormType** ao definir comandos do suplemento.

Depois que um suplemento é ativado, você pode usar a propriedade [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) para obter o item selecionado atualmente no Outlook e a propriedade [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para obter o tipo do item atual.

Opcionalmente, você pode usar o atributo **ItemClass** para especificar a classe de mensagens do item e o atributo **IncludeSubClasses** para especificar se a regra deve ser **true** quando o item é uma subclasse da classe especificada.

Para saber mais sobre classes de mensagens, confira [Tipos de item e classes de mensagens](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).

O exemplo a seguir é uma regra **ItemIs** que permite que os usuários vejam o suplemento na barra de suplementos do Outlook quando o usuário está lendo uma mensagem:

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

O exemplo a seguir é uma regra **ItemIs** que permite que os usuários vejam o suplemento na barra de suplementos do Outlook quando o usuário está lendo uma mensagem ou compromisso.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a>Regra ItemHasAttachment


O tipo complexo **ItemHasAttachment** define uma regra que verifica se o item selecionado contém um anexo.

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>Regra ItemHasKnownEntity

Antes de um item ser disponibilizado para um suplemento, o servidor o examina para determinar se o assunto e o corpo contêm texto que provavelmente é uma das entidades conhecidas. Se uma dessas entidades for encontrada, ela é colocada em uma coleção de entidades conhecidas que você acessa usando o método **getEntities** ou **getEntitiesByType** desse item.

Você pode especificar uma regra usando o **ItemHasKnownEntity** que mostra seu suplemento quando uma entidade do tipo especificado está presente no item. Você pode especificar as seguintes entidades conhecidas no atributo **EntityType** de uma regra **ItemHasKnownEntity**:

-  Endereço
-  Contato
-  EmailAddress
-  MeetingSuggestion
-  PhoneNumber
-  TaskSuggestion
-  URL
    
Opcionalmente, você pode incluir uma expressão regular no atributo **RegularExpression** para que seu suplemento seja exibido somente quando uma entidade que corresponde à expressão regular está presente. Para obter correspondências às expressões regulares especificadas em regras **ItemHasKnownEntity**, você pode usar os métodos **getRegExMatches** ou **getFilteredEntitiesByName** do item do Outlook selecionado no momento.

O exemplo a seguir mostra uma coleção de elementos **Rule** que mostram o suplemento quando uma das entidades conhecidas especificadas está presente na mensagem.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

O exemplo a seguir mostra uma regra **ItemHasKnownEntity** com um atributo **RegularExpression** que ativa o suplemento quando uma URL que contém a palavra "contoso" está presente em uma mensagem.


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Para saber mais sobre entidades nas regras de ativação, confira [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).


## <a name="itemhasregularexpressionmatch-rule"></a>Regra ItemHasRegularExpressionMatch

O tipo complexo **ItemHasRegularExpressionMatch** define uma regra que usa uma expressão regular para corresponder o conteúdo da propriedade especificada de um item. Se o texto que corresponde à expressão regular for encontrado na propriedade especificada do item, o Outlook ativa a barra de suplementos e exibe o suplemento. Você pode usar os métodos **getRegExMatches** ou **getRegExMatchesByName** do objeto que representa o item selecionado atualmente a fim de obter correspondências para a expressão regular especificada.

O exemplo a seguir mostra uma **ItemHasRegularExpressionMatch** que ativa o suplemento quando o corpo do item selecionado contém "apple", "banana" ou "coconut", ignorando maiúsculas e minúsculas.

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Para saber mais sobre como usar a regra **ItemHasRegularExpressionMatch**, confira [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).


## <a name="rulecollection-rule"></a>Regra RuleCollection


O tipo complexo **RuleCollection** combina várias regras em uma única regra. Você pode especificar se as regras na coleção devem ser combinadas com um OU lógico ou um E lógico usando o atributo **Mode**.

Quando um E lógico é especificado, um item deve corresponder a todas as regras especificadas na coleção para mostrar o suplemento. Quando um OU lógico é especificado, um item que corresponde a qualquer das regras especificadas na coleção mostra o suplemento.

Você pode combinar regras **RuleCollection** para formar regras complexas. O exemplo a seguir ativa o suplemento quando o usuário está exibindo um compromisso ou um item de mensagem e o assunto ou corpo do item contém um endereço.

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

O exemplo a seguir ativa o suplemento quando o usuário está redigindo uma mensagem ou quando o usuário está exibindo um compromisso e o assunto ou corpo do compromisso contém um endereço.

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a>Limites para regras e expressões regulares


Para oferecer uma experiência satisfatória com suplementos do Outlook, você deve seguir as diretrizes de ativação e de uso da API. A tabela a seguir mostra os limites gerais para expressões regulares e regras, mas existem regras específicas para hosts diferentes. Para saber mais, confira [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) e [Solucionar problemas de ativação de suplemento do Outlook](troubleshoot-outlook-add-in-activation.md).

<br/>

|**Elemento do suplemento**|**Diretrizes**|
|:-----|:-----|
|Tamanho do manifesto|Não pode exceder 256 KB.|
|Regras|Máximo de 15 regras.|
|ItemHasKnownEntity|Um cliente avançado do Outlook aplicará a regra em relação ao primeiro megabyte do corpo, e não no restante do corpo.|
|Expressões Regulares|Para regras ItemHasKnownEntity ou ItemHasRegularExpressionMatch de todos os hosts do Outlook:<br><ul><li>Especifique no máximo cinco expressões regulares em regras de ativação de um suplemento do Outlook. Não será possível instalar um suplemento se você exceder esse limite.</li><li>Especifica expressões regulares cujos resultados previstos sejam retornados pela chamada de método <b>getRegExMatches</b> nas primeiras 50 correspondências. </li><li>Especifica declarações look-ahead em expressões regulares, mas não look-behind, `(?<=text)` e negative look-behind `(?<!text)`.</li><li>Especifica expressões regulares cuja correspondência não exceda os limites da tabela a seguir.<br/><br/><table><tr><th>Limite de comprimento de uma correspondência de regex</th><th>Clientes avançados do Outlook</th><th>Outlook no iOS e no Android</th></tr><tr><td>O corpo do item é texto sem formatação</td><td>1,5 KB</td><td>3 KB</td></tr><tr><td>Corpo do item em HTML</td><td>3 KB</td><td>3 KB</td></tr></table>|

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de redação](compose-scenario.md)
- [Limites de ativação e da API do JavaScript API para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)
    
