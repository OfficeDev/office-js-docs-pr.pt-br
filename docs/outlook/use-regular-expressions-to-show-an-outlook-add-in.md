---
title: Usar regras de ativação de expressões regulares para mostrar um suplemento
description: Saiba como usar as regras de ativação de expressões regulares para suplementos contextuais do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: b697f1b0a4d20254986a7aa10a5cc7f25dbdd887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605238"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a><span data-ttu-id="83e1c-103">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="83e1c-103">Use regular expression activation rules to show an Outlook add-in</span></span>

<span data-ttu-id="83e1c-104">Você poderá especificar regras de expressão regulares para ativar um [suplemento contextual](contextual-outlook-add-ins.md) quando houver uma correspondência em campos específicos da mensagem.</span><span class="sxs-lookup"><span data-stu-id="83e1c-104">You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message.</span></span> <span data-ttu-id="83e1c-105">Os suplementos contextuais só são ativados no modo de leitura. O Outlook não ativa os suplementos contextuais quando o usuário está redigindo um item.</span><span class="sxs-lookup"><span data-stu-id="83e1c-105">Contextual add-ins activate only in read mode, Outlook does not activate contextual add-ins when the user is composing an item.</span></span> <span data-ttu-id="83e1c-106">Também há outras situações em que o Outlook não ativa suplementos, por exemplo, itens protegidos por IRM (Gerenciamento de Direitos de Informação).</span><span class="sxs-lookup"><span data-stu-id="83e1c-106">There are also other scenarios where Outlook does not activate add-ins, for example, items protected by Information Rights Management (IRM).</span></span> <span data-ttu-id="83e1c-107">Saiba mais em [Regras de ativação para suplementos do Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="83e1c-107">For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="83e1c-108">Você pode especificar uma expressão regular como parte de uma regra [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ou de uma regra [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) no manifesto XML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-108">You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule in the add-in XML manifest.</span></span> <span data-ttu-id="83e1c-109">As regras são especificadas em um ponto de extensão [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).</span><span class="sxs-lookup"><span data-stu-id="83e1c-109">The rules are specified in a [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) extension point.</span></span>

<span data-ttu-id="83e1c-110">O Outlook avalia expressões regulares com base em regras para o intérprete de JavaScript usado pelo navegador no computador cliente.</span><span class="sxs-lookup"><span data-stu-id="83e1c-110">Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer.</span></span> <span data-ttu-id="83e1c-111">O Outlook dá suporte à mesma lista de caracteres especiais que têm suporte em todos os processadores XML.</span><span class="sxs-lookup"><span data-stu-id="83e1c-111">Outlook supports the same list of special characters that all XML processors also support.</span></span> <span data-ttu-id="83e1c-112">A tabela a seguir lista os caracteres especiais.</span><span class="sxs-lookup"><span data-stu-id="83e1c-112">The following table lists these special characters.</span></span> <span data-ttu-id="83e1c-113">Você pode usar esses caracteres em uma expressão regular especificando a sequência de escape para o caractere correspondente, conforme descrito na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="83e1c-113">You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.</span></span>

<br/>

|<span data-ttu-id="83e1c-114">Caractere</span><span class="sxs-lookup"><span data-stu-id="83e1c-114">Character</span></span>|<span data-ttu-id="83e1c-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="83e1c-115">Description</span></span>|<span data-ttu-id="83e1c-116">Sequência de escape a ser usada</span><span class="sxs-lookup"><span data-stu-id="83e1c-116">Escape sequence to use</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="83e1c-117">Aspas duplas</span><span class="sxs-lookup"><span data-stu-id="83e1c-117">Double quotation mark</span></span>|`&quot;`|
|`&`|<span data-ttu-id="83e1c-118">E comercial</span><span class="sxs-lookup"><span data-stu-id="83e1c-118">Ampersand</span></span>|`&amp;`|
|`'`|<span data-ttu-id="83e1c-119">Apóstrofo</span><span class="sxs-lookup"><span data-stu-id="83e1c-119">Apostrophe</span></span>|`&apos;`|
|`<`|<span data-ttu-id="83e1c-120">Sinal menor que</span><span class="sxs-lookup"><span data-stu-id="83e1c-120">Less-than sign</span></span>|`&lt;`|
|`>`|<span data-ttu-id="83e1c-121">Sinal maior que</span><span class="sxs-lookup"><span data-stu-id="83e1c-121">Greater-than sign</span></span>|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="83e1c-122">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="83e1c-122">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="83e1c-123">Uma regra `ItemHasRegularExpressionMatch` é útil para controlar a ativação do suplemento com base em valores específicos de uma propriedade compatível.</span><span class="sxs-lookup"><span data-stu-id="83e1c-123">An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property.</span></span> <span data-ttu-id="83e1c-124">A regra `ItemHasRegularExpressionMatch` tem os seguintes atributos.</span><span class="sxs-lookup"><span data-stu-id="83e1c-124">The `ItemHasRegularExpressionMatch` rule has the following attributes.</span></span>

<br/>

|<span data-ttu-id="83e1c-125">Nome do atributo</span><span class="sxs-lookup"><span data-stu-id="83e1c-125">Attribute name</span></span>|<span data-ttu-id="83e1c-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="83e1c-126">Description</span></span>|
|:-----|:-----|
|`RegExName`|<span data-ttu-id="83e1c-127">Especifica o nome da expressão regular para que você possa referir-se à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-127">Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.</span></span>|
|`RegExValue`|<span data-ttu-id="83e1c-128">Especifica a expressão regular que será avaliada para determinar se o suplemento deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="83e1c-128">Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.</span></span>|
|`PropertyName`|<span data-ttu-id="83e1c-129">Especifica o nome da propriedade em relação à qual a expressão regular será avaliada.</span><span class="sxs-lookup"><span data-stu-id="83e1c-129">Specifies the name of the property that the regular expression will be evaluated against.</span></span> <span data-ttu-id="83e1c-130">Os valores permitidos são `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` e `Subject`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-130">The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.</span></span><br/><br/><span data-ttu-id="83e1c-131">Se você especificar `BodyAsHTML`, o Outlook só aplicará a expressão regular se o corpo do item for HTML.</span><span class="sxs-lookup"><span data-stu-id="83e1c-131">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="83e1c-132">Caso contrário, o Outlook não retornará nenhuma correspondência para essa expressão regular.</span><span class="sxs-lookup"><span data-stu-id="83e1c-132">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="83e1c-133">Se você especificar `BodyAsPlaintext`, o Outlook sempre aplicará a expressão regular no corpo do item.</span><span class="sxs-lookup"><span data-stu-id="83e1c-133">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="83e1c-134">**Observação:** Você deve configurar o `PropertyName` atributo para `BodyAsPlaintext` se você especificar o `Highlight` atributo para o `Rule` elemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-134">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span>|
|`IgnoreCase`|<span data-ttu-id="83e1c-135">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-135">Specifies whether to ignore case when matching the regular expression specified by `RegExName`.</span></span>|
| `Highlight` | <span data-ttu-id="83e1c-136">Especifica como o cliente deve realçar texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="83e1c-136">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="83e1c-137">Esse elemento só pode aplicado em `Rule` elementos dentro de `ExtensionPoint` elementos.</span><span class="sxs-lookup"><span data-stu-id="83e1c-137">This element can only be applied to `Rule` elements within `ExtensionPoint` elements.</span></span> <span data-ttu-id="83e1c-138">Pode ser um dos seguintes: `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-138">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="83e1c-139">Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-139">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="83e1c-140">**Observação:** Você deve configurar o `PropertyName` atributo para `BodyAsPlaintext` se você especificar o `Highlight` atributo para o `Rule` elemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-140">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span> |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a><span data-ttu-id="83e1c-141">Práticas recomendadas para usar expressões regulares em regras</span><span class="sxs-lookup"><span data-stu-id="83e1c-141">Best practices for using regular expressions in rules</span></span>

<span data-ttu-id="83e1c-142">Ao usar expressões regulares, preste bastante atenção ao seguinte:</span><span class="sxs-lookup"><span data-stu-id="83e1c-142">Pay special attention to the following when you use regular expressions:</span></span>

- <span data-ttu-id="83e1c-143">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item.</span><span class="sxs-lookup"><span data-stu-id="83e1c-143">If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item.</span></span> <span data-ttu-id="83e1c-144">O uso de uma expressão regular como `.*` para tentar obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="83e1c-144">Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.</span></span>
- <span data-ttu-id="83e1c-145">O corpo de texto sem formatação retornado em um navegador pode ser sutilmente diferente do retornado em outro.</span><span class="sxs-lookup"><span data-stu-id="83e1c-145">The plain text body returned on one browser can be different in subtle ways on another.</span></span> <span data-ttu-id="83e1c-146">Se você usa uma regra `ItemHasRegularExpressionMatch` com `BodyAsPlaintext` como atributo `PropertyName`, teste sua expressão regular em todos os navegadores compatíveis com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-146">If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.</span></span>

    <span data-ttu-id="83e1c-147">Como diferentes navegadores usam diferentes maneiras de obter o corpo de texto de um item selecionado, você deve se certificar de que sua expressão regular dê suporte a diferenças sutis que possam ser retornadas como parte do corpo de texto.</span><span class="sxs-lookup"><span data-stu-id="83e1c-147">Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text.</span></span> <span data-ttu-id="83e1c-148">Por exemplo, alguns navegadores, como o Internet Explorer 9, usam a propriedade `innerText` do DOM. Outros, como o Firefox, usam o método `.textContent()` para obter o corpo de texto de um item.</span><span class="sxs-lookup"><span data-stu-id="83e1c-148">For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item.</span></span> <span data-ttu-id="83e1c-149">Além disso, navegadores diferentes podem retornar quebras de linha diferentes: uma quebra de linha é `\r\n` no Internet Explorer e `\n` no Firefox e no Chrome.</span><span class="sxs-lookup"><span data-stu-id="83e1c-149">Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome.</span></span> <span data-ttu-id="83e1c-150">Para saber mais, confira [Compatibilidade do DOM do W3C – HTML](https://quirksmode.org/dom/html/).</span><span class="sxs-lookup"><span data-stu-id="83e1c-150">For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).</span></span>

- <span data-ttu-id="83e1c-151">O corpo HTML de um item é um pouco diferente entre um cliente avançado do Outlook e o Outlook na Web ou Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="83e1c-151">The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook mobile.</span></span> <span data-ttu-id="83e1c-152">Defina as expressões regulares com cuidado.</span><span class="sxs-lookup"><span data-stu-id="83e1c-152">Define your regular expressions carefully.</span></span>

- <span data-ttu-id="83e1c-p112">Dependendo do aplicativo host, do tipo de dispositivo ou da propriedade aplicada à expressão regular, há outras práticas recomendadas e limites para cada um dos hosts que você deve estar ciente durante a criação de expressões regulares como regras de ativação. Confira [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="83e1c-p112">Depending on the host application, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the hosts that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.</span></span>

### <a name="examples"></a><span data-ttu-id="83e1c-155">Exemplos</span><span class="sxs-lookup"><span data-stu-id="83e1c-155">Examples</span></span>

<span data-ttu-id="83e1c-156">A regra `ItemHasRegularExpressionMatch` a seguir ativa o suplemento sempre que o endereço de email SMTP do remetente corresponde a `@contoso`, independentemente dos caracteres em maiúsculas ou minúsculas.</span><span class="sxs-lookup"><span data-stu-id="83e1c-156">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

<span data-ttu-id="83e1c-157">A seguir, temos outra maneira de especificar a mesma expressão regular usando o atributo `IgnoreCase`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-157">The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

<span data-ttu-id="83e1c-158">A regra `ItemHasRegularExpressionMatch` a seguir ativa o suplemento sempre que um símbolo de ação estiver incluso no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="83e1c-158">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="83e1c-159">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="83e1c-159">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="83e1c-160">Uma regra `ItemHasKnownEntity` ativa um suplemento com base na existência de uma entidade no assunto ou no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="83e1c-160">An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item.</span></span> <span data-ttu-id="83e1c-161">O tipo [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) define as entidades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="83e1c-161">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities.</span></span> <span data-ttu-id="83e1c-162">A aplicação de uma expressão regular em uma regra `ItemHasKnownEntity` traz praticidade quando a ativação é baseada em um subconjunto de valores de uma entidade (por exemplo, um conjunto específico de URLs ou números de telefone com determinado código de área).</span><span class="sxs-lookup"><span data-stu-id="83e1c-162">Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).</span></span>

> [!NOTE]
> <span data-ttu-id="83e1c-163">O Outlook só pode extrair cadeias de caracteres de entidade em inglês, independentemente da localidade padrão especificada no manifesto.</span><span class="sxs-lookup"><span data-stu-id="83e1c-163">Outlook can only extract entity strings in English regardless of the default locale specified in the manifest.</span></span> <span data-ttu-id="83e1c-164">Somente as mensagens são compatíveis com o tipo entidade `MeetingSuggestion`; os compromissos, não.</span><span class="sxs-lookup"><span data-stu-id="83e1c-164">Only messages support the `MeetingSuggestion` entity type; appointments do not.</span></span> <span data-ttu-id="83e1c-165">Não é possível extrair entidades de itens na pasta **Itens enviados** nem é possível usar uma regra `ItemHasKnownEntity` para ativar um suplemento para itens na pasta **Itens enviados**.</span><span class="sxs-lookup"><span data-stu-id="83e1c-165">You cannot extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.</span></span>

<span data-ttu-id="83e1c-166">A regra `ItemHasKnownEntity` é compatível com os atributos da tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="83e1c-166">The `ItemHasKnownEntity` rule supports the attributes in the following table.</span></span> <span data-ttu-id="83e1c-167">Embora a especificação de uma expressão regular seja opcional em uma regra `ItemHasKnownEntity`, se você optar por usar uma expressão regular como filtro de entidade, deverá especificar ambos os atributos `RegExFilter` e `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-167">Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.</span></span>

<br/>

|<span data-ttu-id="83e1c-168">Nome do atributo</span><span class="sxs-lookup"><span data-stu-id="83e1c-168">Attribute name</span></span>|<span data-ttu-id="83e1c-169">Descrição</span><span class="sxs-lookup"><span data-stu-id="83e1c-169">Description</span></span>|
|:-----|:-----|
|`EntityType`|<span data-ttu-id="83e1c-170">Especifica o tipo de entidade que deve ser encontrado para que a regra seja avaliada como `true`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-170">Specifies the type of entity that must be found for the rule to evaluate to `true`.</span></span> <span data-ttu-id="83e1c-171">Use várias regras para especificar vários tipos de entidades.</span><span class="sxs-lookup"><span data-stu-id="83e1c-171">Use multiple rules to specify multiple types of entities.</span></span>|
|`RegExFilter`|<span data-ttu-id="83e1c-172">Especifica uma expressão regular que filtra mais instâncias da entidade especificada por `EntityType`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-172">Specifies a regular expression that further filters instances of the entity specified by `EntityType`.</span></span>|
|`FilterName`|<span data-ttu-id="83e1c-173">Especifica o nome das expressões regulares especificadas por `RegExFilter` para que seja possível consultá-lo posteriormente por código.</span><span class="sxs-lookup"><span data-stu-id="83e1c-173">Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.</span></span>|
|`IgnoreCase`|<span data-ttu-id="83e1c-174">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por `RegExFilter`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-174">Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.</span></span>|

### <a name="examples"></a><span data-ttu-id="83e1c-175">Exemplos</span><span class="sxs-lookup"><span data-stu-id="83e1c-175">Examples</span></span>

<span data-ttu-id="83e1c-176">A regra `ItemHasKnownEntity` a seguir ativa o suplemento sempre que há uma URL no assunto ou no corpo do item atual e a URL contém a cadeia de caracteres `youtube`, independentemente de maiúsculas e minúsculas na cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="83e1c-176">The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a><span data-ttu-id="83e1c-177">Usar resultados de expressões regulares no código</span><span class="sxs-lookup"><span data-stu-id="83e1c-177">Using regular expression results in code</span></span>

<span data-ttu-id="83e1c-178">Você pode obter correspondências para uma expressão regular usando os seguintes métodos no item atual:</span><span class="sxs-lookup"><span data-stu-id="83e1c-178">You can obtain matches to a regular expression by using the following methods on the current item:</span></span>

- <span data-ttu-id="83e1c-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna correspondências no item atual para todas as expressões regulares especificadas nas regras `ItemHasRegularExpressionMatch` e `ItemHasKnownEntity` do suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.</span></span>

- <span data-ttu-id="83e1c-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna correspondências no item atual para a expressão regular especificada na regra `ItemHasRegularExpressionMatch` do suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.</span></span>

- <span data-ttu-id="83e1c-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) retorna instâncias inteiras de entidades que contêm correspondências para a expressão regular identificada especificada em uma regra `ItemHasKnownEntity` do suplemento.</span><span class="sxs-lookup"><span data-stu-id="83e1c-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.</span></span>

<span data-ttu-id="83e1c-182">Quando as expressões regulares são avaliadas, as correspondências são retornadas para seu suplemento em um objeto de matriz.</span><span class="sxs-lookup"><span data-stu-id="83e1c-182">When the regular expressions are evaluated, the matches are returned to your add-in in an array object.</span></span> <span data-ttu-id="83e1c-183">Para `getRegExMatches`, esse objeto tem o identificador do nome da expressão regular.</span><span class="sxs-lookup"><span data-stu-id="83e1c-183">For `getRegExMatches`, that object has the identifier of the name of the regular expression.</span></span>

> [!NOTE]
> <span data-ttu-id="83e1c-184">O Outlook não retorna correspondências em uma ordem específica na matriz.</span><span class="sxs-lookup"><span data-stu-id="83e1c-184">Outlook does not return matches in any particular order in the array.</span></span> <span data-ttu-id="83e1c-185">Além disso, não considere que as correspondências são retornadas pela mesma ordem nessa matriz, ainda que você execute o mesmo suplemento em cada um desses clientes no mesmo item e na mesma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="83e1c-185">Also, you should not assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.</span></span>

### <a name="examples"></a><span data-ttu-id="83e1c-186">Exemplos</span><span class="sxs-lookup"><span data-stu-id="83e1c-186">Examples</span></span>

<span data-ttu-id="83e1c-187">A seguir temos um exemplo de uma coleção de regras que contém uma regra `ItemHasRegularExpressionMatch` com uma expressão regular denominada `videoURL`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-187">The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

<span data-ttu-id="83e1c-188">O exemplo a seguir usa `getRegExMatches` do item atual para definir uma variável `videos` nos resultados da regra `ItemHasRegularExpressionMatch` anterior.</span><span class="sxs-lookup"><span data-stu-id="83e1c-188">The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

<span data-ttu-id="83e1c-p119">Várias correspondências são armazenadas como elementos de matriz nesse objeto. O exemplo a seguir mostra como repetir correspondências para uma expressão regular denominada `reg1` a fim de construir uma cadeia de caracteres para exibir como HTML.</span><span class="sxs-lookup"><span data-stu-id="83e1c-p119">Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.</span></span>

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

<span data-ttu-id="83e1c-191">A seguir temos um exemplo de uma regra `ItemHasKnownEntity` que especifica a entidade `MeetingSuggestion` e uma expressão regular denominada `CampSuggestion`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-191">The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`.</span></span> <span data-ttu-id="83e1c-192">O Outlook ativará o suplemento se detectar que o atual item selecionado contém uma sugestão de reunião e o assunto ou corpo contêm o termo `WonderCamp`.</span><span class="sxs-lookup"><span data-stu-id="83e1c-192">Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

<span data-ttu-id="83e1c-193">O exemplo de código a seguir usa `getFilteredEntitiesByName` do item atual para definir uma variável `suggestions` para uma matriz de sugestões de reunião detectadas para a regra `ItemHasKnownEntity` anterior.</span><span class="sxs-lookup"><span data-stu-id="83e1c-193">The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a><span data-ttu-id="83e1c-194">Confira também</span><span class="sxs-lookup"><span data-stu-id="83e1c-194">See also</span></span>

- <span data-ttu-id="83e1c-195">[Suplemento do Outlook: número de ordem da Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) – um exemplo do suplemento contextual ativado com base em uma correspondência de expressão regular.</span><span class="sxs-lookup"><span data-stu-id="83e1c-195">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.</span></span>
- [<span data-ttu-id="83e1c-196">Criar suplementos do Outlook para formulários de leitura</span><span class="sxs-lookup"><span data-stu-id="83e1c-196">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="83e1c-197">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="83e1c-197">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="83e1c-198">Limites para ativação e API JavaScript para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="83e1c-198">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="83e1c-199">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="83e1c-199">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="83e1c-200">Práticas recomendadas para expressões regulares no .NET Framework</span><span class="sxs-lookup"><span data-stu-id="83e1c-200">Best Practices for Regular Expressions in the .NET Framework</span></span>](/dotnet/standard/base-types/best-practices)
