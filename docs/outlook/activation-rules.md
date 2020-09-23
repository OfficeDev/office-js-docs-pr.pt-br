---
title: Regras de ativação para suplementos do Outlook
description: O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: cdcdfbf3961ad9f627ba00f7366f49c77bba435d
ms.sourcegitcommit: fd110305c2be8660ab8a47c1da3e3969bd1ede86
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/23/2020
ms.locfileid: "48214593"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="7ca7e-103">Regras de ativação para suplementos contextuais do Outlook</span><span class="sxs-lookup"><span data-stu-id="7ca7e-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="7ca7e-p101">O Outlook ativa alguns tipos de suplementos se a mensagem ou o compromisso que o usuário está lendo ou redigindo satisfaz as regras de ativação do suplemento. Isso é verdadeiro para todos os suplementos que usam o esquema de manifesto 1.1. O usuário pode escolher o suplemento na interface de usuário do Outlook para iniciá-lo em relação ao item atual.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="7ca7e-107">A figura a seguir mostra suplementos do Outlook ativados na barra de suplementos da mensagem que está no painel de leitura.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![Barra de aplicativos mostrando aplicativos de email de leitura ativados](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="7ca7e-109">Especificar regras de ativação em um manifesto</span><span class="sxs-lookup"><span data-stu-id="7ca7e-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="7ca7e-110">Para que o Outlook ative um suplemento para condições específicas, especifique as regras de ativação no manifesto do suplemento usando um dos seguintes `Rule` elementos:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements:</span></span>

- <span data-ttu-id="7ca7e-111">[Elemento Rule (MailApp complexType)](../reference/manifest/rule.md) - especifica uma regra individual.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="7ca7e-112">[Elemento Rule (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - combina várias regras usando operações lógicas.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>
    

 > [!NOTE]
 > <span data-ttu-id="7ca7e-113">O `Rule` elemento que você usa para especificar uma regra individual é do tipo complexo de [regra](../reference/manifest/rule.md) abstrata.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="7ca7e-114">Cada um dos tipos de regra a seguir estende esse `Rule` tipo complexo abstrato.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="7ca7e-115">Portanto, ao especificar uma regra individual em um manifesto, é preciso usar o atributo [xsi:type](https://www.w3.org/TR/xmlschema-1/) para definir um dos tipos de regra a seguir.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="7ca7e-116">Por exemplo, a seguinte regra define uma regra [ItemIs](../reference/manifest/rule.md#itemis-rule): `<Rule xsi:type="ItemIs" ItemType="Message" />`</span><span class="sxs-lookup"><span data-stu-id="7ca7e-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule: `<Rule xsi:type="ItemIs" ItemType="Message" />`</span></span>
 > 
 > <span data-ttu-id="7ca7e-117">O `FormType` atributo se aplica às regras de ativação no manifesto v 1.1, mas não está definido na `VersionOverrides` v 1.0.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="7ca7e-118">Portanto, não pode ser usado quando [itemis](../reference/manifest/rule.md#itemis-rule) é usado no `VersionOverrides` nó.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="7ca7e-p104">A tabela a seguir lista os tipos de regra disponíveis. Veja mais informações após a tabela e nos artigos especificados em [Criar suplementos do Outlook para formulários de leitura](read-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p104">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="7ca7e-121">**Nome da regra**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-121">**Rule name**</span></span>|<span data-ttu-id="7ca7e-122">**Formulários aplicáveis**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-122">**Applicable forms**</span></span>|<span data-ttu-id="7ca7e-123">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="7ca7e-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="7ca7e-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="7ca7e-125">Ler, Redigir</span><span class="sxs-lookup"><span data-stu-id="7ca7e-125">Read, Compose</span></span>|<span data-ttu-id="7ca7e-p105">Verifica se o item atual é do tipo especificado (compromisso ou mensagem). Pode também verificar a classe do item e o tipo de formulário e, opcionalmente, a classe de mensagem do item.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p105">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="7ca7e-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="7ca7e-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="7ca7e-129">Leitura</span><span class="sxs-lookup"><span data-stu-id="7ca7e-129">Read</span></span>|<span data-ttu-id="7ca7e-130">Verifica se o item selecionado contém um anexo.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="7ca7e-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7ca7e-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="7ca7e-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="7ca7e-132">Read</span></span>|<span data-ttu-id="7ca7e-p106">Verifica se o item selecionado contém uma ou mais entidades conhecidas. Mais informações: [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p106">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="7ca7e-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="7ca7e-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="7ca7e-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="7ca7e-136">Read</span></span>|<span data-ttu-id="7ca7e-137">Verifica se o endereço de email do remetente, o assunto e/ou o corpo do item selecionado contêm uma correspondência para uma expressão regular. Mais informações: [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="7ca7e-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7ca7e-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="7ca7e-139">Ler, Redigir</span><span class="sxs-lookup"><span data-stu-id="7ca7e-139">Read, Compose</span></span>|<span data-ttu-id="7ca7e-140">Combina uma coleção de regras para que você forme regras mais complexas.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="7ca7e-141">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="7ca7e-141">ItemIs rule</span></span>

<span data-ttu-id="7ca7e-142">O tipo complexo **ItemIs** define uma regra que avalia **true** se o item atual coincidir com o tipo de item e, opcionalmente, a classe de mensagens do item, se estiver declarada na regra.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="7ca7e-143">Especifique um dos tipos de item a seguir no `ItemType` atributo de uma regra **itemis** .</span><span class="sxs-lookup"><span data-stu-id="7ca7e-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="7ca7e-144">Você pode especificar mais de uma regra **ItemIs** em um manifesto.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="7ca7e-145">O tipo simples ItemType define os tipos de itens do Outlook que dão suporte aos suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="7ca7e-146">**Valor**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-146">**Value**</span></span>|<span data-ttu-id="7ca7e-147">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7ca7e-148">**Compromisso**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-148">**Appointment**</span></span>|<span data-ttu-id="7ca7e-p108">Especifica um item em um calendário do Outlook. Isso inclui um item de reunião que foi respondido e que tem um organizador e participantes, ou um compromisso que não tem um organizador ou participantes e é simplesmente um item no calendário. Isso corresponde à classe de mensagens IPM.Appointment no Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p108">Specifies an item in an Outlook calendar. This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="7ca7e-151">**Mensagem**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-151">**Message**</span></span>|<span data-ttu-id="7ca7e-152">Especifica um dos seguintes itens recebidos normalmente na Caixa de Entrada:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-152">Specifies one of the following items received in typically the Inbox:</span></span> <ul><li><p><span data-ttu-id="7ca7e-p109">Uma mensagem de email. Isso corresponde à classe de mensagem IPM.Note no Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p109">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="7ca7e-p110">Uma solicitação de reunião, resposta ou cancelamento. Isso corresponde às seguintes classes de mensagem no Outlook:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p110">A meeting request, response, or cancellation. This corresponds to the following  message classes in Outlook:</span></span></p><p><span data-ttu-id="7ca7e-157">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="7ca7e-157">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="7ca7e-158">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="7ca7e-158">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="7ca7e-159">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="7ca7e-159">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="7ca7e-160">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="7ca7e-160">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="7ca7e-161">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="7ca7e-161">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="7ca7e-162">O `FormType` atributo é usado para especificar o modo (leitura ou composição) no qual o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-162">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="7ca7e-163">O atributo Itemis `FormType` é definido no esquema v 1.1 e posterior, mas não em `VersionOverrides` v 1.0.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-163">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="7ca7e-164">Não inclua o `FormType` atributo ao definir comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-164">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="7ca7e-165">Depois que um suplemento é ativado, você pode usar a propriedade [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) para obter o item selecionado atualmente no Outlook e a propriedade [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) para obter o tipo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-165">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="7ca7e-166">Opcionalmente, você pode usar o `ItemClass` atributo para especificar a classe de mensagem do item, e o `IncludeSubClasses` atributo para especificar se a regra deve ser **true** quando o item é uma subclasse da classe especificada.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-166">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="7ca7e-167">Para saber mais sobre classes de mensagens, confira [Tipos de item e classes de mensagens](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-167">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="7ca7e-168">O exemplo a seguir é uma regra **ItemIs** que permite que os usuários vejam o suplemento na barra de suplementos do Outlook quando o usuário está lendo uma mensagem:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-168">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message:</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="7ca7e-169">O exemplo a seguir é uma regra **ItemIs** que permite que os usuários vejam o suplemento na barra de suplementos do Outlook quando o usuário está lendo uma mensagem ou compromisso.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="7ca7e-170">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="7ca7e-170">ItemHasAttachment rule</span></span>


<span data-ttu-id="7ca7e-171">O `ItemHasAttachment` tipo complexo define uma regra que verifica se o item selecionado contém um anexo.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-171">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="7ca7e-172">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7ca7e-172">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="7ca7e-p112">Antes de um item ser disponibilizado para um suplemento, o servidor examina-o para determinar se o assunto e o corpo contêm qualquer texto que provavelmente seja uma das entidades conhecidas. Se qualquer uma dessas entidades for encontrada, ela será colocada em uma coleção de entidades conhecidas que você acessa usando o `getEntities` método ou `getEntitiesByType` desse item.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p112">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities. If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="7ca7e-p113">Você pode especificar uma regra usando `ItemHasKnownEntity` que mostre o suplemento quando uma entidade do tipo especificado estiver presente no item. Você pode especificar as seguintes entidades conhecidas no `EntityType` atributo de uma `ItemHasKnownEntity` regra:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p113">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item. You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule:</span></span>

- <span data-ttu-id="7ca7e-177">Endereço</span><span class="sxs-lookup"><span data-stu-id="7ca7e-177">Address</span></span>
- <span data-ttu-id="7ca7e-178">Contato</span><span class="sxs-lookup"><span data-stu-id="7ca7e-178">Contact</span></span>
- <span data-ttu-id="7ca7e-179">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="7ca7e-179">EmailAddress</span></span>
- <span data-ttu-id="7ca7e-180">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7ca7e-180">MeetingSuggestion</span></span>
- <span data-ttu-id="7ca7e-181">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7ca7e-181">PhoneNumber</span></span>
- <span data-ttu-id="7ca7e-182">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7ca7e-182">TaskSuggestion</span></span>
- <span data-ttu-id="7ca7e-183">URL</span><span class="sxs-lookup"><span data-stu-id="7ca7e-183">URL</span></span>
    
<span data-ttu-id="7ca7e-p114">Opcionalmente, você pode incluir uma expressão regular no `RegularExpression` atributo para que seu suplemento seja mostrado apenas quando uma entidade que corresponde à expressão regular no presente. Para obter correspondências com as expressões regulares especificadas nas `ItemHasKnownEntity` regras, você pode usar o `getRegExMatches` `getFilteredEntitiesByName` método ou para o item do Outlook selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p114">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present. To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="7ca7e-186">O exemplo a seguir mostra uma coleção de `Rule` elementos que mostram o suplemento quando uma das entidades conhecidas especificadas está presente na mensagem.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-186">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="7ca7e-187">O exemplo a seguir mostra uma `ItemHasKnownEntity` regra com um `RegularExpression` atributo que ativa o suplemento quando uma URL que contém a palavra "contoso" está presente em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-187">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="7ca7e-188">Para saber mais sobre entidades nas regras de ativação, confira [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-188">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="7ca7e-189">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="7ca7e-189">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="7ca7e-p115">O `ItemHasRegularExpressionMatch` tipo complexo define uma regra que usa uma expressão regular para corresponder ao conteúdo da propriedade especificada de um item. Se o texto que corresponde à expressão regular for encontrado na propriedade especificada do item, o Outlook ativará a barra de suplementos e exibirá o suplemento. Você pode usar o `getRegExMatches` `getRegExMatchesByName` método ou do objeto que representa o item selecionado no momento para obter correspondências para a expressão regular especificada.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p115">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item. If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in. You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="7ca7e-193">O exemplo a seguir mostra um `ItemHasRegularExpressionMatch` que ativa o suplemento quando o corpo do item selecionado contém "Apple", "banana" ou "Coconut", ignorando maiúsculas e minúsculas.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-193">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="7ca7e-194">Para obter mais informações sobre como usar a `ItemHasRegularExpressionMatch` regra, confira [usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-194">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="7ca7e-195">Regra RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7ca7e-195">RuleCollection rule</span></span>


<span data-ttu-id="7ca7e-p116">O `RuleCollection` tipo complexo combina várias regras em uma única regra. Você pode especificar se as regras na coleção devem ser combinadas com um lógica ou lógica e usando o `Mode` atributo.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p116">The `RuleCollection` complex type combines multiple rules into a single rule. You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="7ca7e-p117">Quando um E lógico é especificado, um item deve corresponder a todas as regras especificadas na coleção para mostrar o suplemento. Quando um OU lógico é especificado, um item que corresponde a qualquer das regras especificadas na coleção mostra o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p117">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="7ca7e-p118">Você pode combinar `RuleCollection` regras para formar regras complexas. O exemplo a seguir ativa o suplemento quando o usuário está exibindo um compromisso ou item de mensagem, e o assunto ou corpo do item contém um endereço.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p118">You can combine `RuleCollection` rules to form complex rules. The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="7ca7e-202">O exemplo a seguir ativa o suplemento quando o usuário está redigindo uma mensagem ou quando o usuário está exibindo um compromisso e o assunto ou corpo do compromisso contém um endereço.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-202">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="7ca7e-203">Limites para regras e expressões regulares</span><span class="sxs-lookup"><span data-stu-id="7ca7e-203">Limits for rules and regular expressions</span></span>


<span data-ttu-id="7ca7e-p119">Para fornecer uma experiência satisfatória com os suplementos do Outlook, você deve cumprir as diretrizes de ativação e uso da API. A tabela a seguir mostra os limites gerais para expressões regulares e regras, mas há regras específicas para aplicativos diferentes. Para obter mais informações, consulte [limites para ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) e [solução de problemas de ativação do suplemento do Outlook](troubleshoot-outlook-add-in-activation.md).</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p119">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines. The following table shows general limits for regular expressions and rules but there are specific rules for different applications. For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="7ca7e-207">**Elemento do suplemento**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-207">**Add-in element**</span></span>|<span data-ttu-id="7ca7e-208">**Diretrizes**</span><span class="sxs-lookup"><span data-stu-id="7ca7e-208">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7ca7e-209">Tamanho do manifesto</span><span class="sxs-lookup"><span data-stu-id="7ca7e-209">Manifest Size</span></span>|<span data-ttu-id="7ca7e-210">Não pode exceder 256 KB.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-210">No larger than 256 KB.</span></span>|
|<span data-ttu-id="7ca7e-211">Regras</span><span class="sxs-lookup"><span data-stu-id="7ca7e-211">Rules</span></span>|<span data-ttu-id="7ca7e-212">Máximo de 15 regras.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-212">No more than 15 rules.</span></span>|
|<span data-ttu-id="7ca7e-213">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7ca7e-213">ItemHasKnownEntity</span></span>|<span data-ttu-id="7ca7e-214">Um cliente avançado do Outlook aplicará a regra em relação ao primeiro megabyte do corpo, e não no restante do corpo.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-214">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="7ca7e-215">Expressões Regulares</span><span class="sxs-lookup"><span data-stu-id="7ca7e-215">Regular Expressions</span></span>|<span data-ttu-id="7ca7e-216">Para as regras do ItemHasKnownEntity ou do ItemHasRegularExpressionMatch para todos os aplicativos do Outlook:</span><span class="sxs-lookup"><span data-stu-id="7ca7e-216">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="7ca7e-p120">Especifique no máximo cinco expressões regulares em regras de ativação de um suplemento do Outlook. Não será possível instalar um suplemento se você exceder esse limite.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-p120">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="7ca7e-219">Especifica expressões regulares cujos resultados previstos sejam retornados pela chamada de método <b>getRegExMatches</b> nas primeiras 50 correspondências.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-219">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="7ca7e-220">Especifica declarações look-ahead em expressões regulares, mas não look-behind, `(?<=text)` e negative look-behind `(?<!text)`.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-220">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="7ca7e-221">Especifica expressões regulares cuja correspondência não exceda os limites da tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="7ca7e-221">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="7ca7e-222">Limite de comprimento de uma correspondência de regex</span><span class="sxs-lookup"><span data-stu-id="7ca7e-222">Limit on length of a regex match</span></span></th><th><span data-ttu-id="7ca7e-223">Clientes avançados do Outlook</span><span class="sxs-lookup"><span data-stu-id="7ca7e-223">Outlook rich clients</span></span></th><th><span data-ttu-id="7ca7e-224">Outlook no iOS e no Android</span><span class="sxs-lookup"><span data-stu-id="7ca7e-224">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="7ca7e-225">O corpo do item é texto sem formatação</span><span class="sxs-lookup"><span data-stu-id="7ca7e-225">Item body is plain text</span></span></td><td><span data-ttu-id="7ca7e-226">1,5 KB</span><span class="sxs-lookup"><span data-stu-id="7ca7e-226">1.5 KB</span></span></td><td><span data-ttu-id="7ca7e-227">3 KB</span><span class="sxs-lookup"><span data-stu-id="7ca7e-227">3 KB</span></span></td></tr><tr><td><span data-ttu-id="7ca7e-228">Corpo do item em HTML</span><span class="sxs-lookup"><span data-stu-id="7ca7e-228">Item body it HTML</span></span></td><td><span data-ttu-id="7ca7e-229">3 KB</span><span class="sxs-lookup"><span data-stu-id="7ca7e-229">3 KB</span></span></td><td><span data-ttu-id="7ca7e-230">3 KB</span><span class="sxs-lookup"><span data-stu-id="7ca7e-230">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="7ca7e-231">Confira também</span><span class="sxs-lookup"><span data-stu-id="7ca7e-231">See also</span></span>

- [<span data-ttu-id="7ca7e-232">Criar suplementos do Outlook para formulários de redação</span><span class="sxs-lookup"><span data-stu-id="7ca7e-232">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="7ca7e-233">Limites de ativação e da API do JavaScript API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="7ca7e-233">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="7ca7e-234">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="7ca7e-234">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="7ca7e-235">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="7ca7e-235">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
