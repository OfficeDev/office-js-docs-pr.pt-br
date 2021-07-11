---
title: Regras de ativação para suplementos do Outlook
description: O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349018"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="ee567-103">Regras de ativação para suplementos contextuais do Outlook</span><span class="sxs-lookup"><span data-stu-id="ee567-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="ee567-p101">O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento. Isso é verdadeiro para todos os suplementos que usam o esquema de manifesto 1.1. O usuário pode escolher o suplemento na interface de usuário do Outlook para iniciá-lo em relação ao item atual.</span><span class="sxs-lookup"><span data-stu-id="ee567-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="ee567-107">A figura a seguir mostra suplementos do Outlook ativados na barra de suplementos da mensagem que está no painel de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee567-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![Barra de aplicativos mostrando aplicativos de email de leitura ativados.](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="ee567-109">Especificar regras de ativação em um manifesto</span><span class="sxs-lookup"><span data-stu-id="ee567-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="ee567-110">Para Outlook ativar um complemento para condições específicas, especifique regras de ativação no manifesto do complemento usando um dos seguintes `Rule` elementos.</span><span class="sxs-lookup"><span data-stu-id="ee567-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements.</span></span>

- <span data-ttu-id="ee567-111">[Elemento Rule (MailApp complexType)](../reference/manifest/rule.md) - especifica uma regra individual.</span><span class="sxs-lookup"><span data-stu-id="ee567-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="ee567-112">[Elemento Rule (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - combina várias regras usando operações lógicas.</span><span class="sxs-lookup"><span data-stu-id="ee567-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>


 > [!NOTE]
 > <span data-ttu-id="ee567-113">O `Rule` elemento que você usa para especificar uma regra individual é do tipo complexo [Rule](../reference/manifest/rule.md) abstrato.</span><span class="sxs-lookup"><span data-stu-id="ee567-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="ee567-114">Cada um dos seguintes tipos de regras estende esse tipo `Rule` complexo abstrato.</span><span class="sxs-lookup"><span data-stu-id="ee567-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="ee567-115">Portanto, ao especificar uma regra individual em um manifesto, é preciso usar o atributo [xsi:type](https://www.w3.org/TR/xmlschema-1/) para definir um dos tipos de regra a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee567-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="ee567-116">Por exemplo, a regra a seguir define uma [regra ItemIs.](../reference/manifest/rule.md#itemis-rule)</span><span class="sxs-lookup"><span data-stu-id="ee567-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule.</span></span>
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > <span data-ttu-id="ee567-117">O `FormType` atributo se aplica às regras de ativação no manifesto v1.1, mas não é definido em `VersionOverrides` v1.0.</span><span class="sxs-lookup"><span data-stu-id="ee567-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="ee567-118">Portanto, ele não pode ser usado [quando ItemIs](../reference/manifest/rule.md#itemis-rule) é usado no `VersionOverrides` nó.</span><span class="sxs-lookup"><span data-stu-id="ee567-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="ee567-p105">A tabela a seguir lista os tipos de regra disponíveis. Veja mais informações após a tabela e nos artigos especificados em [Criar suplementos do Outlook para formulários de leitura](read-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-p105">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="ee567-121">**Nome da regra**</span><span class="sxs-lookup"><span data-stu-id="ee567-121">**Rule name**</span></span>|<span data-ttu-id="ee567-122">**Formulários aplicáveis**</span><span class="sxs-lookup"><span data-stu-id="ee567-122">**Applicable forms**</span></span>|<span data-ttu-id="ee567-123">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ee567-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="ee567-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="ee567-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="ee567-125">Ler, Redigir</span><span class="sxs-lookup"><span data-stu-id="ee567-125">Read, Compose</span></span>|<span data-ttu-id="ee567-p106">Verifica se o item atual é do tipo especificado (compromisso ou mensagem). Pode também verificar a classe do item e o tipo de formulário e, opcionalmente, a classe de mensagem do item.</span><span class="sxs-lookup"><span data-stu-id="ee567-p106">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="ee567-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="ee567-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="ee567-129">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee567-129">Read</span></span>|<span data-ttu-id="ee567-130">Verifica se o item selecionado contém um anexo.</span><span class="sxs-lookup"><span data-stu-id="ee567-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="ee567-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ee567-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="ee567-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee567-132">Read</span></span>|<span data-ttu-id="ee567-p107">Verifica se o item selecionado contém uma ou mais entidades conhecidas. Mais informações: [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-p107">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="ee567-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="ee567-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="ee567-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee567-136">Read</span></span>|<span data-ttu-id="ee567-137">Verifica se o endereço de email do remetente, o assunto e/ou o corpo do item selecionado contêm uma correspondência para uma expressão regular. Mais informações: [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="ee567-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ee567-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="ee567-139">Ler, Redigir</span><span class="sxs-lookup"><span data-stu-id="ee567-139">Read, Compose</span></span>|<span data-ttu-id="ee567-140">Combina uma coleção de regras para que você forme regras mais complexas.</span><span class="sxs-lookup"><span data-stu-id="ee567-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="ee567-141">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="ee567-141">ItemIs rule</span></span>

<span data-ttu-id="ee567-142">O tipo complexo **ItemIs** define uma regra que avalia **true** se o item atual coincidir com o tipo de item e, opcionalmente, a classe de mensagens do item, se estiver declarada na regra.</span><span class="sxs-lookup"><span data-stu-id="ee567-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="ee567-143">Especifique um dos seguintes tipos de item `ItemType` no atributo de uma regra **ItemIs.**</span><span class="sxs-lookup"><span data-stu-id="ee567-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="ee567-144">Você pode especificar mais de uma regra **ItemIs** em um manifesto.</span><span class="sxs-lookup"><span data-stu-id="ee567-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="ee567-145">O tipo simples ItemType define os tipos de itens do Outlook que dão suporte aos suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee567-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="ee567-146">**Valor**</span><span class="sxs-lookup"><span data-stu-id="ee567-146">**Value**</span></span>|<span data-ttu-id="ee567-147">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ee567-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ee567-148">**Compromisso**</span><span class="sxs-lookup"><span data-stu-id="ee567-148">**Appointment**</span></span>|<span data-ttu-id="ee567-149">Especifica um item em um Outlook calendário.</span><span class="sxs-lookup"><span data-stu-id="ee567-149">Specifies an item in an Outlook calendar.</span></span> <span data-ttu-id="ee567-150">Isso inclui um item de reunião que foi respondido e tem um organizador e participantes, ou um compromisso que não tem um organizador ou participante e é simplesmente um item no calendário.</span><span class="sxs-lookup"><span data-stu-id="ee567-150">This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.</span></span> <span data-ttu-id="ee567-151">Isso corresponde ao IPM. Classe de mensagem de compromisso Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee567-151">This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="ee567-152">**Mensagem**</span><span class="sxs-lookup"><span data-stu-id="ee567-152">**Message**</span></span>|<span data-ttu-id="ee567-153">Especifica um dos seguintes itens recebidos normalmente na Caixa de Entrada.</span><span class="sxs-lookup"><span data-stu-id="ee567-153">Specifies one of the following items received in typically the Inbox.</span></span> <ul><li><p><span data-ttu-id="ee567-p110">Uma mensagem de email. Isso corresponde à classe de mensagem IPM.Note no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee567-p110">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="ee567-156">Uma solicitação de reunião, resposta ou cancelamento.</span><span class="sxs-lookup"><span data-stu-id="ee567-156">A meeting request, response, or cancellation.</span></span> <span data-ttu-id="ee567-157">Isso corresponde às seguintes classes de mensagem no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee567-157">This corresponds to the following message classes in Outlook.</span></span></p><p><span data-ttu-id="ee567-158">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="ee567-158">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="ee567-159">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="ee567-159">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="ee567-160">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="ee567-160">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="ee567-161">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="ee567-161">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="ee567-162">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="ee567-162">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="ee567-163">O atributo é usado para especificar o modo (leitura ou `FormType` redação) no qual o complemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="ee567-163">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="ee567-164">O atributo ItemIs `FormType` é definido no esquema v1.1 e posterior, mas não em `VersionOverrides` v1.0.</span><span class="sxs-lookup"><span data-stu-id="ee567-164">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="ee567-165">Não inclua o `FormType` atributo ao definir comandos de complemento.</span><span class="sxs-lookup"><span data-stu-id="ee567-165">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="ee567-166">Depois que um suplemento é ativado, você pode usar a propriedade [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) para obter o item selecionado atualmente no Outlook e a propriedade [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para obter o tipo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee567-166">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="ee567-167">Opcionalmente, você pode usar o atributo para especificar a classe de mensagem do item e o atributo para especificar se a regra deve ser verdadeira quando o item for uma `ItemClass` subclasse da `IncludeSubClasses` classe especificada. </span><span class="sxs-lookup"><span data-stu-id="ee567-167">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="ee567-168">Para saber mais sobre classes de mensagens, confira [Tipos de item e classes de mensagens](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span><span class="sxs-lookup"><span data-stu-id="ee567-168">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="ee567-169">O exemplo a seguir é uma **regra ItemIs** que permite que os usuários vejam o complemento na barra de Outlook de complementos quando o usuário estiver lendo uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee567-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message.</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="ee567-170">O exemplo a seguir é uma regra **ItemIs** que permite que os usuários vejam o suplemento na barra de suplementos do Outlook quando o usuário está lendo uma mensagem ou compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee567-170">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="ee567-171">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="ee567-171">ItemHasAttachment rule</span></span>


<span data-ttu-id="ee567-172">O `ItemHasAttachment` tipo complexo define uma regra que verifica se o item selecionado contém um anexo.</span><span class="sxs-lookup"><span data-stu-id="ee567-172">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="ee567-173">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ee567-173">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="ee567-174">Antes de um item ser disponibilizado para um suplemento, o servidor o examina para determinar se o assunto e o corpo contêm texto que provavelmente é uma das entidades conhecidas.</span><span class="sxs-lookup"><span data-stu-id="ee567-174">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities.</span></span> <span data-ttu-id="ee567-175">Se alguma dessas entidades for encontrada, ela será colocada em uma coleção de entidades conhecidas que você acessa usando o ou o método `getEntities` `getEntitiesByType` desse item.</span><span class="sxs-lookup"><span data-stu-id="ee567-175">If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="ee567-176">Você pode especificar uma regra usando que mostra o seu complemento quando uma entidade do `ItemHasKnownEntity` tipo especificado está presente no item.</span><span class="sxs-lookup"><span data-stu-id="ee567-176">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item.</span></span> <span data-ttu-id="ee567-177">Você pode especificar as seguintes entidades conhecidas no `EntityType` atributo de uma `ItemHasKnownEntity` regra.</span><span class="sxs-lookup"><span data-stu-id="ee567-177">You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule.</span></span>

- <span data-ttu-id="ee567-178">Endereço</span><span class="sxs-lookup"><span data-stu-id="ee567-178">Address</span></span>
- <span data-ttu-id="ee567-179">Contato</span><span class="sxs-lookup"><span data-stu-id="ee567-179">Contact</span></span>
- <span data-ttu-id="ee567-180">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="ee567-180">EmailAddress</span></span>
- <span data-ttu-id="ee567-181">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ee567-181">MeetingSuggestion</span></span>
- <span data-ttu-id="ee567-182">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ee567-182">PhoneNumber</span></span>
- <span data-ttu-id="ee567-183">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ee567-183">TaskSuggestion</span></span>
- <span data-ttu-id="ee567-184">URL</span><span class="sxs-lookup"><span data-stu-id="ee567-184">URL</span></span>

<span data-ttu-id="ee567-185">Opcionalmente, você pode incluir uma expressão regular no atributo para que o seu complemento seja mostrado somente quando uma entidade que corresponde à `RegularExpression` expressão regular presente.</span><span class="sxs-lookup"><span data-stu-id="ee567-185">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present.</span></span> <span data-ttu-id="ee567-186">Para obter combinações com expressões regulares especificadas em regras, você pode usar o método ou para o item Outlook `ItemHasKnownEntity` `getRegExMatches` selecionado no `getFilteredEntitiesByName` momento.</span><span class="sxs-lookup"><span data-stu-id="ee567-186">To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="ee567-187">O exemplo a seguir mostra uma coleção de elementos que mostram o complemento quando uma das entidades conhecidas especificadas está `Rule` presente na mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee567-187">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="ee567-188">O exemplo a seguir mostra uma regra com um atributo que ativa o complemento quando uma URL que contém a palavra `ItemHasKnownEntity` `RegularExpression` "contoso" está presente em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee567-188">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="ee567-189">Para saber mais sobre entidades nas regras de ativação, confira [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-189">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="ee567-190">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="ee567-190">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="ee567-191">O tipo complexo define uma regra que usa uma expressão regular para corresponder ao conteúdo da `ItemHasRegularExpressionMatch` propriedade especificada de um item.</span><span class="sxs-lookup"><span data-stu-id="ee567-191">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item.</span></span> <span data-ttu-id="ee567-192">Se o texto que corresponde à expressão regular for encontrado na propriedade especificada do item, o Outlook ativa a barra de suplementos e exibe o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ee567-192">If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in.</span></span> <span data-ttu-id="ee567-193">Você pode usar o ou o método do objeto que representa o item selecionado no momento para obter corresponde à `getRegExMatches` `getRegExMatchesByName` expressão regular especificada.</span><span class="sxs-lookup"><span data-stu-id="ee567-193">You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="ee567-194">O exemplo a seguir mostra um que ativa o complemento quando o corpo do item selecionado contém `ItemHasRegularExpressionMatch` "apple", "banana" ou "coco", ignorando o caso.</span><span class="sxs-lookup"><span data-stu-id="ee567-194">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="ee567-195">Para obter mais informações sobre como usar a `ItemHasRegularExpressionMatch` regra, consulte [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-195">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="ee567-196">Regra RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ee567-196">RuleCollection rule</span></span>


<span data-ttu-id="ee567-197">O `RuleCollection` tipo complexo combina várias regras em uma única regra.</span><span class="sxs-lookup"><span data-stu-id="ee567-197">The `RuleCollection` complex type combines multiple rules into a single rule.</span></span> <span data-ttu-id="ee567-198">Você pode especificar se as regras na coleção devem ser combinadas com um OR lógico ou um E lógico usando o `Mode` atributo.</span><span class="sxs-lookup"><span data-stu-id="ee567-198">You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="ee567-p118">Quando um E lógico é especificado, um item deve corresponder a todas as regras especificadas na coleção para mostrar o suplemento. Quando um OU lógico é especificado, um item que corresponde a qualquer das regras especificadas na coleção mostra o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ee567-p118">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="ee567-201">Você pode combinar `RuleCollection` regras para formar regras complexas.</span><span class="sxs-lookup"><span data-stu-id="ee567-201">You can combine `RuleCollection` rules to form complex rules.</span></span> <span data-ttu-id="ee567-202">O exemplo a seguir ativa o suplemento quando o usuário está exibindo um compromisso ou um item de mensagem e o assunto ou corpo do item contém um endereço.</span><span class="sxs-lookup"><span data-stu-id="ee567-202">The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="ee567-203">O exemplo a seguir ativa o suplemento quando o usuário está redigindo uma mensagem ou quando o usuário está exibindo um compromisso e o assunto ou corpo do compromisso contém um endereço.</span><span class="sxs-lookup"><span data-stu-id="ee567-203">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="ee567-204">Limites para regras e expressões regulares</span><span class="sxs-lookup"><span data-stu-id="ee567-204">Limits for rules and regular expressions</span></span>


<span data-ttu-id="ee567-205">Para oferecer uma experiência satisfatória com suplementos do Outlook, você deve seguir as diretrizes de ativação e de uso da API.</span><span class="sxs-lookup"><span data-stu-id="ee567-205">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines.</span></span> <span data-ttu-id="ee567-206">A tabela a seguir mostra limites gerais para expressões e regras regulares, mas há regras específicas para diferentes aplicativos.</span><span class="sxs-lookup"><span data-stu-id="ee567-206">The following table shows general limits for regular expressions and rules but there are specific rules for different applications.</span></span> <span data-ttu-id="ee567-207">Para saber mais, confira [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) e [Solucionar problemas de ativação de suplemento do Outlook](troubleshoot-outlook-add-in-activation.md).</span><span class="sxs-lookup"><span data-stu-id="ee567-207">For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="ee567-208">**Elemento do suplemento**</span><span class="sxs-lookup"><span data-stu-id="ee567-208">**Add-in element**</span></span>|<span data-ttu-id="ee567-209">**Diretrizes**</span><span class="sxs-lookup"><span data-stu-id="ee567-209">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ee567-210">Tamanho do manifesto</span><span class="sxs-lookup"><span data-stu-id="ee567-210">Manifest Size</span></span>|<span data-ttu-id="ee567-211">Não pode exceder 256 KB.</span><span class="sxs-lookup"><span data-stu-id="ee567-211">No larger than 256 KB.</span></span>|
|<span data-ttu-id="ee567-212">Regras</span><span class="sxs-lookup"><span data-stu-id="ee567-212">Rules</span></span>|<span data-ttu-id="ee567-213">Máximo de 15 regras.</span><span class="sxs-lookup"><span data-stu-id="ee567-213">No more than 15 rules.</span></span>|
|<span data-ttu-id="ee567-214">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ee567-214">ItemHasKnownEntity</span></span>|<span data-ttu-id="ee567-215">Um cliente avançado do Outlook aplicará a regra em relação ao primeiro megabyte do corpo, e não no restante do corpo.</span><span class="sxs-lookup"><span data-stu-id="ee567-215">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="ee567-216">Expressões Regulares</span><span class="sxs-lookup"><span data-stu-id="ee567-216">Regular Expressions</span></span>|<span data-ttu-id="ee567-217">Para regras ItemHasKnownEntity ou ItemHasRegularExpressionMatch para todos os Outlook aplicativos:</span><span class="sxs-lookup"><span data-stu-id="ee567-217">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="ee567-p121">Especifique no máximo cinco expressões regulares em regras de ativação de um suplemento do Outlook. Não será possível instalar um suplemento se você exceder esse limite.</span><span class="sxs-lookup"><span data-stu-id="ee567-p121">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="ee567-220">Especifica expressões regulares cujos resultados previstos sejam retornados pela chamada de método <b>getRegExMatches</b> nas primeiras 50 correspondências.</span><span class="sxs-lookup"><span data-stu-id="ee567-220">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="ee567-221">Especifica declarações look-ahead em expressões regulares, mas não look-behind, `(?<=text)` e negative look-behind `(?<!text)`.</span><span class="sxs-lookup"><span data-stu-id="ee567-221">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="ee567-222">Especifica expressões regulares cuja correspondência não exceda os limites da tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee567-222">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="ee567-223">Limite de comprimento de uma correspondência de regex</span><span class="sxs-lookup"><span data-stu-id="ee567-223">Limit on length of a regex match</span></span></th><th><span data-ttu-id="ee567-224">Clientes avançados do Outlook</span><span class="sxs-lookup"><span data-stu-id="ee567-224">Outlook rich clients</span></span></th><th><span data-ttu-id="ee567-225">Outlook no iOS e no Android</span><span class="sxs-lookup"><span data-stu-id="ee567-225">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="ee567-226">O corpo do item é texto sem formatação</span><span class="sxs-lookup"><span data-stu-id="ee567-226">Item body is plain text</span></span></td><td><span data-ttu-id="ee567-227">1,5 KB</span><span class="sxs-lookup"><span data-stu-id="ee567-227">1.5 KB</span></span></td><td><span data-ttu-id="ee567-228">3 KB</span><span class="sxs-lookup"><span data-stu-id="ee567-228">3 KB</span></span></td></tr><tr><td><span data-ttu-id="ee567-229">Corpo do item em HTML</span><span class="sxs-lookup"><span data-stu-id="ee567-229">Item body it HTML</span></span></td><td><span data-ttu-id="ee567-230">3 KB</span><span class="sxs-lookup"><span data-stu-id="ee567-230">3 KB</span></span></td><td><span data-ttu-id="ee567-231">3 KB</span><span class="sxs-lookup"><span data-stu-id="ee567-231">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="ee567-232">Confira também</span><span class="sxs-lookup"><span data-stu-id="ee567-232">See also</span></span>

- [<span data-ttu-id="ee567-233">Criar suplementos do Outlook para formulários de redação</span><span class="sxs-lookup"><span data-stu-id="ee567-233">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="ee567-234">Limites de ativação e da API do JavaScript API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ee567-234">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="ee567-235">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="ee567-235">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="ee567-236">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="ee567-236">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
