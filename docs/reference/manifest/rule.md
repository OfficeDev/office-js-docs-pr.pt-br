---
title: Elemento Rule no arquivo de manifesto
description: O elemento Rule especifica as regras de ativação que devem ser avaliadas para esse complemento de email contextual.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 60882a5e36a63832cf81eab9320b113a420b84a3
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348668"
---
# <a name="rule-element"></a><span data-ttu-id="cbc98-103">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="cbc98-103">Rule element</span></span>

<span data-ttu-id="cbc98-104">Especifica as regras de ativação que devem ser avaliadas para esse complemento de email contextual.</span><span class="sxs-lookup"><span data-stu-id="cbc98-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="cbc98-105">**Tipo de complemento:** Email (contextual)</span><span class="sxs-lookup"><span data-stu-id="cbc98-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="cbc98-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="cbc98-106">Contained in</span></span>

- [<span data-ttu-id="cbc98-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cbc98-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="cbc98-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (preterido)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span><span class="sxs-lookup"><span data-stu-id="cbc98-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="cbc98-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="cbc98-109">Attributes</span></span>

| <span data-ttu-id="cbc98-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="cbc98-110">Attribute</span></span> | <span data-ttu-id="cbc98-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cbc98-111">Required</span></span> | <span data-ttu-id="cbc98-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbc98-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="cbc98-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="cbc98-113">**xsi:type**</span></span> | <span data-ttu-id="cbc98-114">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-114">Yes</span></span> | <span data-ttu-id="cbc98-115">O tipo de regra que está sendo definida.</span><span class="sxs-lookup"><span data-stu-id="cbc98-115">The type of rule being defined.</span></span> |

<span data-ttu-id="cbc98-116">O tipo de regra pode ser um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="cbc98-116">The type of rule can be one of the following:</span></span>

- [<span data-ttu-id="cbc98-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="cbc98-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="cbc98-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="cbc98-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="cbc98-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="cbc98-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="cbc98-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="cbc98-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="cbc98-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="cbc98-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="cbc98-122">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="cbc98-122">ItemIs rule</span></span>

<span data-ttu-id="cbc98-123">Define uma regra que é avaliada como true se o item selecionado for do tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="cbc98-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="cbc98-124">Atributos</span><span class="sxs-lookup"><span data-stu-id="cbc98-124">Attributes</span></span>

| <span data-ttu-id="cbc98-125">Atributo</span><span class="sxs-lookup"><span data-stu-id="cbc98-125">Attribute</span></span> | <span data-ttu-id="cbc98-126">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cbc98-126">Required</span></span> | <span data-ttu-id="cbc98-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbc98-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="cbc98-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="cbc98-128">**ItemType**</span></span> | <span data-ttu-id="cbc98-129">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-129">Yes</span></span> | <span data-ttu-id="cbc98-p101">Especifica o tipo de item para fazer a correspondência. Pode ser `Message` ou `Appointment`. O tipo de item `Message` inclui email, solicitações de reunião, respostas de reunião e cancelamentos de reunião.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="cbc98-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="cbc98-133">**FormType**</span></span> | <span data-ttu-id="cbc98-134">Não (dentro de [ExtensionPoint](extensionpoint.md)), Sim (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="cbc98-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="cbc98-p102">Especifica se o aplicativo deve aparecer no formulário de leitura ou edição do item. Pode ser um dos seguintes: `Read`, `Edit`, `ReadOrEdit`. Se não for especificado em um `Rule` dentro de um `ExtensionPoint`, esse valor DEVERÁ ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="cbc98-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="cbc98-138">**ItemClass**</span></span> | <span data-ttu-id="cbc98-139">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-139">No</span></span> | <span data-ttu-id="cbc98-p103">Especifica a classe de mensagens personalizada para fazer a correspondência. Para saber mais, confira o artigo [Ativar um suplemento de email no Outlook para uma classe de mensagens específica](../../outlook/activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="cbc98-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="cbc98-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="cbc98-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="cbc98-143">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-143">No</span></span> | <span data-ttu-id="cbc98-144">Especifica se a regra deve ser avaliada como true se o item pertencer a uma subclasse da classe de mensagens especificada. O padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="cbc98-145">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cbc98-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="cbc98-146">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="cbc98-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="cbc98-147">Define uma regra que é avaliada como true se o item contiver um anexo.</span><span class="sxs-lookup"><span data-stu-id="cbc98-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="cbc98-148">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cbc98-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="cbc98-149">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="cbc98-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="cbc98-150">Define uma regra que é avaliada como true se o item contiver texto do tipo de entidade especificada em seu assunto ou corpo.</span><span class="sxs-lookup"><span data-stu-id="cbc98-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="cbc98-151">Atributos</span><span class="sxs-lookup"><span data-stu-id="cbc98-151">Attributes</span></span>

| <span data-ttu-id="cbc98-152">Atributo</span><span class="sxs-lookup"><span data-stu-id="cbc98-152">Attribute</span></span> | <span data-ttu-id="cbc98-153">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cbc98-153">Required</span></span> | <span data-ttu-id="cbc98-154">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbc98-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="cbc98-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="cbc98-155">**EntityType**</span></span> | <span data-ttu-id="cbc98-156">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-156">Yes</span></span> | <span data-ttu-id="cbc98-p104">Especifica o tipo de entidade que deve ser encontrado para que a regra para que ela seja avaliada como true. Pode ser um dos seguintes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="cbc98-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="cbc98-159">**RegExFilter**</span></span> | <span data-ttu-id="cbc98-160">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-160">No</span></span> | <span data-ttu-id="cbc98-161">Especifica uma expressão regular para executar esta entidade para ativação.</span><span class="sxs-lookup"><span data-stu-id="cbc98-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="cbc98-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="cbc98-162">**FilterName**</span></span> | <span data-ttu-id="cbc98-163">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-163">No</span></span> | <span data-ttu-id="cbc98-164">Especifica o nome do filtro de expressões regulares para que seja possível consultá-lo posteriormente no código do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cbc98-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="cbc98-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="cbc98-165">**IgnoreCase**</span></span> | <span data-ttu-id="cbc98-166">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-166">No</span></span> | <span data-ttu-id="cbc98-167">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="cbc98-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="cbc98-168">**Realce**</span><span class="sxs-lookup"><span data-stu-id="cbc98-168">**Highlight**</span></span> | <span data-ttu-id="cbc98-169">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-169">No</span></span> | <span data-ttu-id="cbc98-p105">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar entidades correspondentes. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="cbc98-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cbc98-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="cbc98-175">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="cbc98-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="cbc98-176">Define uma regra que é avaliada como true se uma correspondência para a expressão regular especificada pode ser encontrada na propriedade especificada do item.</span><span class="sxs-lookup"><span data-stu-id="cbc98-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="cbc98-177">Atributos</span><span class="sxs-lookup"><span data-stu-id="cbc98-177">Attributes</span></span>

| <span data-ttu-id="cbc98-178">Atributo</span><span class="sxs-lookup"><span data-stu-id="cbc98-178">Attribute</span></span> | <span data-ttu-id="cbc98-179">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cbc98-179">Required</span></span> | <span data-ttu-id="cbc98-180">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbc98-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="cbc98-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="cbc98-181">**RegExName**</span></span> | <span data-ttu-id="cbc98-182">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-182">Yes</span></span> | <span data-ttu-id="cbc98-183">Especifica o nome da expressão regular para que você possa fazer referência à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cbc98-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="cbc98-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="cbc98-184">**RegExValue**</span></span> | <span data-ttu-id="cbc98-185">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-185">Yes</span></span> | <span data-ttu-id="cbc98-186">Especifica a expressão regular que será avaliada para determinar se o suplemento de email deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="cbc98-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="cbc98-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="cbc98-187">**PropertyName**</span></span> | <span data-ttu-id="cbc98-188">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-188">Yes</span></span> | <span data-ttu-id="cbc98-p106">Especifica o nome da propriedade em relação a qual expressão regular será avaliada. Pode ser um dos seguintes: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="cbc98-191">Se você especificar `BodyAsHTML`, o Outlook só aplicará a expressão regular se o corpo do item for HTML.</span><span class="sxs-lookup"><span data-stu-id="cbc98-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="cbc98-192">Caso contrário, o Outlook não retornará nenhuma correspondência para essa expressão regular.</span><span class="sxs-lookup"><span data-stu-id="cbc98-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="cbc98-193">Se você especificar `BodyAsPlaintext`, o Outlook sempre aplicará a expressão regular no corpo do item.</span><span class="sxs-lookup"><span data-stu-id="cbc98-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="cbc98-194">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="cbc98-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="cbc98-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="cbc98-195">**IgnoreCase**</span></span> | <span data-ttu-id="cbc98-196">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-196">No</span></span> | <span data-ttu-id="cbc98-197">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada pelo atributo **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="cbc98-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="cbc98-198">**Realce**</span><span class="sxs-lookup"><span data-stu-id="cbc98-198">**Highlight**</span></span> | <span data-ttu-id="cbc98-199">Não</span><span class="sxs-lookup"><span data-stu-id="cbc98-199">No</span></span> | <span data-ttu-id="cbc98-200">Especifica como o cliente deve realçar texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="cbc98-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="cbc98-201">Esse atributo pode ser aplicado apenas à elementos **regra** dentro dos elementos **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="cbc98-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="cbc98-202">Pode ser um dos seguintes: `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="cbc98-203">Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="cbc98-204">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="cbc98-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="cbc98-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cbc98-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="cbc98-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="cbc98-206">RuleCollection</span></span>

<span data-ttu-id="cbc98-207">Define uma coleção de regras e o operador lógico a ser usado ao avaliá-las.</span><span class="sxs-lookup"><span data-stu-id="cbc98-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="cbc98-208">Atributos</span><span class="sxs-lookup"><span data-stu-id="cbc98-208">Attributes</span></span>

| <span data-ttu-id="cbc98-209">Atributo</span><span class="sxs-lookup"><span data-stu-id="cbc98-209">Attribute</span></span> | <span data-ttu-id="cbc98-210">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cbc98-210">Required</span></span> | <span data-ttu-id="cbc98-211">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbc98-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="cbc98-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="cbc98-212">**Mode**</span></span> | <span data-ttu-id="cbc98-213">Sim</span><span class="sxs-lookup"><span data-stu-id="cbc98-213">Yes</span></span> | <span data-ttu-id="cbc98-p109">Especifica o operador lógico a ser usado quando estiver avaliando essa coleção de regras. Pode ser: `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="cbc98-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="cbc98-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cbc98-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="cbc98-217">Confira também</span><span class="sxs-lookup"><span data-stu-id="cbc98-217">See also</span></span>

- [<span data-ttu-id="cbc98-218">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="cbc98-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="cbc98-219">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="cbc98-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="cbc98-220">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="cbc98-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
