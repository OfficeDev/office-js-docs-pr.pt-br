---
title: Usar regras de ativação de expressões regulares para mostrar um suplemento
description: Saiba como usar as regras de ativação de expressões regulares para suplementos contextuais do Outlook.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: d334ba6b2e0f044fc8d876cd6edd218743ccb390
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937627"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook

Você poderá especificar regras de expressão regulares para ativar um [suplemento contextual](contextual-outlook-add-ins.md) quando houver uma correspondência em campos específicos da mensagem. Os suplementos contextuais só são ativados no modo de leitura. O Outlook não ativa os suplementos contextuais quando o usuário está redigindo um item. Há também outros cenários em que o Outlook não ativa os complementos, por exemplo, itens assinados digitalmente. Saiba mais em [Regras de ativação para suplementos do Outlook](activation-rules.md).

Você pode especificar uma expressão regular como parte de uma regra [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ou de uma regra [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) no manifesto XML do suplemento. As regras são especificadas em um ponto de extensão [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).

O Outlook avalia expressões regulares com base em regras para o intérprete de JavaScript usado pelo navegador no computador cliente. O Outlook dá suporte à mesma lista de caracteres especiais que têm suporte em todos os processadores XML. A tabela a seguir lista os caracteres especiais. Você pode usar esses caracteres em uma expressão regular especificando a sequência de escape para o caractere correspondente, conforme descrito na tabela a seguir.

<br/>

|Caractere|Descrição|Sequência de escape a ser usada|
|:-----|:-----|:-----|
|`"`|Aspas duplas|`&quot;`|
|`&`|E comercial|`&amp;`|
|`'`|Apóstrofo|`&apos;`|
|`<`|Sinal menor que|`&lt;`|
|`>`|Sinal maior que|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>Regra ItemHasRegularExpressionMatch

Uma regra `ItemHasRegularExpressionMatch` é útil para controlar a ativação do suplemento com base em valores específicos de uma propriedade compatível. A regra `ItemHasRegularExpressionMatch` tem os seguintes atributos.

<br/>

|Nome do atributo|Descrição|
|:-----|:-----|
|`RegExName`|Especifica o nome da expressão regular para que você possa referir-se à expressão no código de seu suplemento.|
|`RegExValue`|Especifica a expressão regular que será avaliada para determinar se o suplemento deve ser mostrado.|
|`PropertyName`|Especifica o nome da propriedade em relação à qual a expressão regular será avaliada. Os valores permitidos são `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` e `Subject`.<br/><br/>Se você especificar `BodyAsHTML`, o Outlook só aplicará a expressão regular se o corpo do item for HTML. Caso contrário, o Outlook não retornará nenhuma correspondência para essa expressão regular.<br/><br/>Se você especificar `BodyAsPlaintext`, o Outlook sempre aplicará a expressão regular no corpo do item.<br/><br/>**Observação:** Você deve configurar o `PropertyName` atributo para `BodyAsPlaintext` se você especificar o `Highlight` atributo para o `Rule` elemento.|
|`IgnoreCase`|Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por `RegExName`.|
| `Highlight` | Especifica como o cliente deve realçar texto correspondente. Esse elemento só pode aplicado em `Rule` elementos dentro de `ExtensionPoint` elementos. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.<br/><br/>**Observação:** Você deve configurar o `PropertyName` atributo para `BodyAsPlaintext` se você especificar o `Highlight` atributo para o `Rule` elemento. |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Práticas recomendadas para usar expressões regulares em regras

Preste atenção especial ao seguinte ao usar expressões regulares.

- Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. O uso de uma expressão regular como `.*` para tentar obter todo o corpo de um item nem sempre retorna os resultados esperados.
- O corpo de texto sem formatação retornado em um navegador pode ser sutilmente diferente do retornado em outro. Se você usa uma regra `ItemHasRegularExpressionMatch` com `BodyAsPlaintext` como atributo `PropertyName`, teste sua expressão regular em todos os navegadores compatíveis com o suplemento.

    Como diferentes navegadores usam diferentes maneiras de obter o corpo de texto de um item selecionado, você deve se certificar de que sua expressão regular dê suporte a diferenças sutis que possam ser retornadas como parte do corpo de texto. Por exemplo, alguns navegadores, como o Internet Explorer 9, usam a propriedade `innerText` do DOM. Outros, como o Firefox, usam o método `.textContent()` para obter o corpo de texto de um item. Além disso, navegadores diferentes podem retornar quebras de linha diferentes: uma quebra de linha é `\r\n` no Internet Explorer e `\n` no Firefox e no Chrome. Para saber mais, confira [Compatibilidade do DOM do W3C – HTML](https://quirksmode.org/dom/html/).

- O corpo HTML de um item é um pouco diferente entre um cliente avançado do Outlook e o Outlook na Web ou Outlook Mobile. Defina as expressões regulares com cuidado.

- Dependendo do cliente Outlook, tipo de dispositivo ou propriedade em que uma expressão regular está sendo aplicada, há outras práticas recomendadas e limites para cada um dos clientes que você deve estar ciente ao projetar expressões regulares como regras de ativação. Confira [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) para saber mais.

### <a name="examples"></a>Exemplos

A regra `ItemHasRegularExpressionMatch` a seguir ativa o suplemento sempre que o endereço de email SMTP do remetente corresponde a `@contoso`, independentemente dos caracteres em maiúsculas ou minúsculas.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

A seguir, temos outra maneira de especificar a mesma expressão regular usando o atributo `IgnoreCase`.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

A regra `ItemHasRegularExpressionMatch` a seguir ativa o suplemento sempre que um símbolo de ação estiver incluso no corpo do item atual.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>Regra ItemHasKnownEntity

Uma regra `ItemHasKnownEntity` ativa um suplemento com base na existência de uma entidade no assunto ou no corpo do item selecionado. O tipo [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) define as entidades compatíveis. A aplicação de uma expressão regular em uma regra `ItemHasKnownEntity` traz praticidade quando a ativação é baseada em um subconjunto de valores de uma entidade (por exemplo, um conjunto específico de URLs ou números de telefone com determinado código de área).

> [!NOTE]
> O Outlook só pode extrair cadeias de caracteres de entidade em inglês, independentemente da localidade padrão especificada no manifesto. Somente as mensagens são compatíveis com o tipo entidade `MeetingSuggestion`; os compromissos, não. Não é possível extrair entidades de itens na pasta **Itens enviados** nem é possível usar uma regra `ItemHasKnownEntity` para ativar um suplemento para itens na pasta **Itens enviados**.

A regra `ItemHasKnownEntity` é compatível com os atributos da tabela a seguir. Embora a especificação de uma expressão regular seja opcional em uma regra `ItemHasKnownEntity`, se você optar por usar uma expressão regular como filtro de entidade, deverá especificar ambos os atributos `RegExFilter` e `FilterName`.

<br/>

|Nome do atributo|Descrição|
|:-----|:-----|
|`EntityType`|Especifica o tipo de entidade que deve ser encontrado para que a regra seja avaliada como `true`. Use várias regras para especificar vários tipos de entidades.|
|`RegExFilter`|Especifica uma expressão regular que filtra mais instâncias da entidade especificada por `EntityType`.|
|`FilterName`|Especifica o nome das expressões regulares especificadas por `RegExFilter` para que seja possível consultá-lo posteriormente por código.|
|`IgnoreCase`|Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por `RegExFilter`.|

### <a name="examples"></a>Exemplos

A regra `ItemHasKnownEntity` a seguir ativa o suplemento sempre que há uma URL no assunto ou no corpo do item atual e a URL contém a cadeia de caracteres `youtube`, independentemente de maiúsculas e minúsculas na cadeia de caracteres.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>Usar resultados de expressões regulares no código

Você pode obter corresponde a uma expressão regular usando os métodos a seguir no item atual.

- [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna correspondências no item atual para todas as expressões regulares especificadas nas regras `ItemHasRegularExpressionMatch` e `ItemHasKnownEntity` do suplemento.

- [getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna correspondências no item atual para a expressão regular especificada na regra `ItemHasRegularExpressionMatch` do suplemento.

- [getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna instâncias inteiras de entidades que contêm correspondências para a expressão regular identificada especificada em uma regra `ItemHasKnownEntity` do suplemento.

Quando as expressões regulares são avaliadas, as correspondências são retornadas para seu suplemento em um objeto de matriz. Para `getRegExMatches`, esse objeto tem o identificador do nome da expressão regular.

> [!NOTE]
> O Outlook não retorna correspondências em uma ordem específica na matriz. Além disso, não considere que as correspondências são retornadas pela mesma ordem nessa matriz, ainda que você execute o mesmo suplemento em cada um desses clientes no mesmo item e na mesma caixa de correio.

### <a name="examples"></a>Exemplos

A seguir temos um exemplo de uma coleção de regras que contém uma regra `ItemHasRegularExpressionMatch` com uma expressão regular denominada `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

O exemplo a seguir usa `getRegExMatches` do item atual para definir uma variável `videos` nos resultados da regra `ItemHasRegularExpressionMatch` anterior.

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

Várias correspondências são armazenadas como elementos de matriz nesse objeto. O exemplo a seguir mostra como repetir correspondências para uma expressão regular denominada `reg1` a fim de construir uma cadeia de caracteres para exibir como HTML.

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

A seguir temos um exemplo de uma regra `ItemHasKnownEntity` que especifica a entidade `MeetingSuggestion` e uma expressão regular denominada `CampSuggestion`. O Outlook ativará o suplemento se detectar que o atual item selecionado contém uma sugestão de reunião e o assunto ou corpo contêm o termo `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

O exemplo de código a seguir usa `getFilteredEntitiesByName` do item atual para definir uma variável `suggestions` para uma matriz de sugestões de reunião detectadas para a regra `ItemHasKnownEntity` anterior.

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>Confira também

- [Suplemento do Outlook: número de ordem da Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) – um exemplo do suplemento contextual ativado com base em uma correspondência de expressão regular.
- [Criar suplementos do Outlook para formulários de leitura](read-scenario.md)
- [Regras de ativação para suplementos do Outlook](activation-rules.md)
- [Limites para ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)
- [Práticas recomendadas para expressões regulares no .NET Framework](/dotnet/standard/base-types/best-practices)
