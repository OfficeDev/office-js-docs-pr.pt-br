---
title: Elemento Rule no arquivo de manifesto
description: O elemento Rule especifica as regras de ativação que devem ser avaliadas para esse suplemento de email contextual.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 79b97f2e442e9d8ce59d17467161b5b9b7a7252d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641428"
---
# <a name="rule-element"></a><span data-ttu-id="c99fe-103">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="c99fe-103">Rule element</span></span>

<span data-ttu-id="c99fe-104">Especifica as regras de ativação que devem ser avaliadas para esse suplemento de email contextual.</span><span class="sxs-lookup"><span data-stu-id="c99fe-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="c99fe-105">**Tipo de suplemento:** Email (contextual)</span><span class="sxs-lookup"><span data-stu-id="c99fe-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="c99fe-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="c99fe-106">Contained in</span></span>

- [<span data-ttu-id="c99fe-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c99fe-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="c99fe-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (preterido)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span><span class="sxs-lookup"><span data-stu-id="c99fe-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="c99fe-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="c99fe-109">Attributes</span></span>

| <span data-ttu-id="c99fe-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="c99fe-110">Attribute</span></span> | <span data-ttu-id="c99fe-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c99fe-111">Required</span></span> | <span data-ttu-id="c99fe-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="c99fe-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c99fe-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c99fe-113">**xsi:type**</span></span> | <span data-ttu-id="c99fe-114">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-114">Yes</span></span> | <span data-ttu-id="c99fe-115">O tipo de regra que está sendo definida.</span><span class="sxs-lookup"><span data-stu-id="c99fe-115">The type of rule being defined.</span></span> |

<span data-ttu-id="c99fe-116">O tipo de regra pode ser um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="c99fe-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="c99fe-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="c99fe-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="c99fe-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c99fe-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="c99fe-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c99fe-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="c99fe-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c99fe-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="c99fe-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c99fe-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="c99fe-122">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="c99fe-122">ItemIs rule</span></span>

<span data-ttu-id="c99fe-123">Define uma regra que é avaliada como true se o item selecionado for do tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="c99fe-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="c99fe-124">Atributos</span><span class="sxs-lookup"><span data-stu-id="c99fe-124">Attributes</span></span>

| <span data-ttu-id="c99fe-125">Atributo</span><span class="sxs-lookup"><span data-stu-id="c99fe-125">Attribute</span></span> | <span data-ttu-id="c99fe-126">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c99fe-126">Required</span></span> | <span data-ttu-id="c99fe-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="c99fe-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c99fe-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="c99fe-128">**ItemType**</span></span> | <span data-ttu-id="c99fe-129">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-129">Yes</span></span> | <span data-ttu-id="c99fe-p101">Especifica o tipo de item para fazer a correspondência. Pode ser `Message` ou `Appointment`. O tipo de item `Message` inclui email, solicitações de reunião, respostas de reunião e cancelamentos de reunião.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="c99fe-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="c99fe-133">**FormType**</span></span> | <span data-ttu-id="c99fe-134">Não (dentro de [ExtensionPoint](extensionpoint.md)), Sim (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="c99fe-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="c99fe-p102">Especifica se o aplicativo deve aparecer no formulário de leitura ou edição do item. Pode ser um dos seguintes: `Read`, `Edit`, `ReadOrEdit`. Se não for especificado em um `Rule` dentro de um `ExtensionPoint`, esse valor DEVERÁ ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="c99fe-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="c99fe-138">**ItemClass**</span></span> | <span data-ttu-id="c99fe-139">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-139">No</span></span> | <span data-ttu-id="c99fe-p103">Especifica a classe de mensagens personalizada para fazer a correspondência. Para saber mais, confira o artigo [Ativar um suplemento de email no Outlook para uma classe de mensagens específica](../../outlook/activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="c99fe-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="c99fe-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="c99fe-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="c99fe-143">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-143">No</span></span> | <span data-ttu-id="c99fe-144">Especifica se a regra deve ser avaliada como true se o item pertencer a uma subclasse da classe de mensagens especificada. O padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="c99fe-145">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c99fe-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="c99fe-146">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c99fe-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="c99fe-147">Define uma regra que é avaliada como true se o item contiver um anexo.</span><span class="sxs-lookup"><span data-stu-id="c99fe-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="c99fe-148">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c99fe-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="c99fe-149">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c99fe-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="c99fe-150">Define uma regra que é avaliada como true se o item contiver texto do tipo de entidade especificada em seu assunto ou corpo.</span><span class="sxs-lookup"><span data-stu-id="c99fe-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="c99fe-151">Atributos</span><span class="sxs-lookup"><span data-stu-id="c99fe-151">Attributes</span></span>

| <span data-ttu-id="c99fe-152">Atributo</span><span class="sxs-lookup"><span data-stu-id="c99fe-152">Attribute</span></span> | <span data-ttu-id="c99fe-153">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c99fe-153">Required</span></span> | <span data-ttu-id="c99fe-154">Descrição</span><span class="sxs-lookup"><span data-stu-id="c99fe-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c99fe-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="c99fe-155">**EntityType**</span></span> | <span data-ttu-id="c99fe-156">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-156">Yes</span></span> | <span data-ttu-id="c99fe-p104">Especifica o tipo de entidade que deve ser encontrado para que a regra para que ela seja avaliada como true. Pode ser um dos seguintes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="c99fe-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="c99fe-159">**RegExFilter**</span></span> | <span data-ttu-id="c99fe-160">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-160">No</span></span> | <span data-ttu-id="c99fe-161">Especifica uma expressão regular para executar esta entidade para ativação.</span><span class="sxs-lookup"><span data-stu-id="c99fe-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="c99fe-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="c99fe-162">**FilterName**</span></span> | <span data-ttu-id="c99fe-163">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-163">No</span></span> | <span data-ttu-id="c99fe-164">Especifica o nome do filtro de expressões regulares para que seja possível consultá-lo posteriormente no código do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="c99fe-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="c99fe-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c99fe-165">**IgnoreCase**</span></span> | <span data-ttu-id="c99fe-166">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-166">No</span></span> | <span data-ttu-id="c99fe-167">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="c99fe-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="c99fe-168">**Realce**</span><span class="sxs-lookup"><span data-stu-id="c99fe-168">**Highlight**</span></span> | <span data-ttu-id="c99fe-169">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-169">No</span></span> | <span data-ttu-id="c99fe-p105">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar entidades correspondentes. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="c99fe-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c99fe-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="c99fe-175">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c99fe-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="c99fe-176">Define uma regra que é avaliada como true se uma correspondência para a expressão regular especificada pode ser encontrada na propriedade especificada do item.</span><span class="sxs-lookup"><span data-stu-id="c99fe-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="c99fe-177">Atributos</span><span class="sxs-lookup"><span data-stu-id="c99fe-177">Attributes</span></span>

| <span data-ttu-id="c99fe-178">Atributo</span><span class="sxs-lookup"><span data-stu-id="c99fe-178">Attribute</span></span> | <span data-ttu-id="c99fe-179">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c99fe-179">Required</span></span> | <span data-ttu-id="c99fe-180">Descrição</span><span class="sxs-lookup"><span data-stu-id="c99fe-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c99fe-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="c99fe-181">**RegExName**</span></span> | <span data-ttu-id="c99fe-182">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-182">Yes</span></span> | <span data-ttu-id="c99fe-183">Especifica o nome da expressão regular para que você possa fazer referência à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="c99fe-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="c99fe-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="c99fe-184">**RegExValue**</span></span> | <span data-ttu-id="c99fe-185">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-185">Yes</span></span> | <span data-ttu-id="c99fe-186">Especifica a expressão regular que será avaliada para determinar se o suplemento de email deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="c99fe-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="c99fe-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="c99fe-187">**PropertyName**</span></span> | <span data-ttu-id="c99fe-188">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-188">Yes</span></span> | <span data-ttu-id="c99fe-p106">Especifica o nome da propriedade em relação a qual expressão regular será avaliada. Pode ser um dos seguintes: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="c99fe-191">Se você especificar `BodyAsHTML`, o Outlook só aplicará a expressão regular se o corpo do item for HTML.</span><span class="sxs-lookup"><span data-stu-id="c99fe-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="c99fe-192">Caso contrário, o Outlook não retornará nenhuma correspondência para essa expressão regular.</span><span class="sxs-lookup"><span data-stu-id="c99fe-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="c99fe-193">Se você especificar `BodyAsPlaintext`, o Outlook sempre aplicará a expressão regular no corpo do item.</span><span class="sxs-lookup"><span data-stu-id="c99fe-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="c99fe-194">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="c99fe-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="c99fe-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c99fe-195">**IgnoreCase**</span></span> | <span data-ttu-id="c99fe-196">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-196">No</span></span> | <span data-ttu-id="c99fe-197">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada pelo atributo **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="c99fe-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="c99fe-198">**Realce**</span><span class="sxs-lookup"><span data-stu-id="c99fe-198">**Highlight**</span></span> | <span data-ttu-id="c99fe-199">Não</span><span class="sxs-lookup"><span data-stu-id="c99fe-199">No</span></span> | <span data-ttu-id="c99fe-200">Especifica como o cliente deve realçar texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="c99fe-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="c99fe-201">Esse atributo pode ser aplicado apenas à elementos **regra** dentro dos elementos **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="c99fe-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="c99fe-202">Pode ser um dos seguintes: `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="c99fe-203">Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="c99fe-204">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="c99fe-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="c99fe-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c99fe-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="c99fe-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c99fe-206">RuleCollection</span></span>

<span data-ttu-id="c99fe-207">Define uma coleção de regras e o operador lógico a ser usado ao avaliá-las.</span><span class="sxs-lookup"><span data-stu-id="c99fe-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="c99fe-208">Atributos</span><span class="sxs-lookup"><span data-stu-id="c99fe-208">Attributes</span></span>

| <span data-ttu-id="c99fe-209">Atributo</span><span class="sxs-lookup"><span data-stu-id="c99fe-209">Attribute</span></span> | <span data-ttu-id="c99fe-210">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c99fe-210">Required</span></span> | <span data-ttu-id="c99fe-211">Descrição</span><span class="sxs-lookup"><span data-stu-id="c99fe-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c99fe-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="c99fe-212">**Mode**</span></span> | <span data-ttu-id="c99fe-213">Sim</span><span class="sxs-lookup"><span data-stu-id="c99fe-213">Yes</span></span> | <span data-ttu-id="c99fe-p109">Especifica o operador lógico a ser usado quando estiver avaliando essa coleção de regras. Pode ser: `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="c99fe-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="c99fe-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c99fe-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="c99fe-217">Confira também</span><span class="sxs-lookup"><span data-stu-id="c99fe-217">See also</span></span>

- [<span data-ttu-id="c99fe-218">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="c99fe-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="c99fe-219">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="c99fe-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="c99fe-220">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="c99fe-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
