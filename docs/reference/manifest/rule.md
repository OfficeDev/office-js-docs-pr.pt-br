---
title: Elemento Rule no arquivo de manifesto
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: ce7763ecb4ef81587ccacbd4090a6f412baf99b2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433108"
---
# <a name="rule-element"></a><span data-ttu-id="0eb75-102">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="0eb75-102">Rule element</span></span>

<span data-ttu-id="0eb75-103">Especifica a(s) regra(s) de ativação que deve(m) ser avaliada(s) para este suplemento contextual de email.</span><span class="sxs-lookup"><span data-stu-id="0eb75-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="0eb75-104">**Tipo de suplemento:** Suplemento contextual de email</span><span class="sxs-lookup"><span data-stu-id="0eb75-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="0eb75-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="0eb75-105">Contained in</span></span>

- [<span data-ttu-id="0eb75-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0eb75-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="0eb75-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0eb75-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="0eb75-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="0eb75-108">Attributes</span></span>

| <span data-ttu-id="0eb75-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="0eb75-109">Attribute</span></span> | <span data-ttu-id="0eb75-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0eb75-110">Required</span></span> | <span data-ttu-id="0eb75-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="0eb75-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="0eb75-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="0eb75-112">**xsi:type**</span></span> | <span data-ttu-id="0eb75-113">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-113">Yes</span></span> | <span data-ttu-id="0eb75-114">O tipo de regra que está sendo definida.</span><span class="sxs-lookup"><span data-stu-id="0eb75-114">The type of rule being defined.</span></span> |

<span data-ttu-id="0eb75-115">O tipo de regra pode ser um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="0eb75-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="0eb75-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="0eb75-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="0eb75-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="0eb75-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="0eb75-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="0eb75-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="0eb75-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="0eb75-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="0eb75-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="0eb75-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="0eb75-121">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="0eb75-121">ItemIs rule</span></span>

<span data-ttu-id="0eb75-122">Define uma regra que é avaliada como true se o item selecionado for do tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="0eb75-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="0eb75-123">Atributos</span><span class="sxs-lookup"><span data-stu-id="0eb75-123">Attributes</span></span>

| <span data-ttu-id="0eb75-124">Atributo</span><span class="sxs-lookup"><span data-stu-id="0eb75-124">Attribute</span></span> | <span data-ttu-id="0eb75-125">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0eb75-125">Required</span></span> | <span data-ttu-id="0eb75-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="0eb75-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="0eb75-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="0eb75-127">**ItemType**</span></span> | <span data-ttu-id="0eb75-128">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-128">Yes</span></span> | <span data-ttu-id="0eb75-p101">Especifica o tipo de item para fazer a correspondência. Pode ser `Message` ou `Appointment`. O tipo de item `Message` inclui email, solicitações de reunião, respostas a reuniões e cancelamentos de reuniões.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="0eb75-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="0eb75-132">**FormType**</span></span> | <span data-ttu-id="0eb75-133">Não (dentro de [ExtensionPoint](extensionpoint.md)), Sim (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="0eb75-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="0eb75-p102">Especifica se o aplicativo deve aparecer no formulário de leitura ou edição do item. Pode ser um dos seguintes: `Read`, `Edit`, `ReadOrEdit`. Se não for especificado em um `Rule` dentro de um `ExtensionPoint`, esse valor DEVERÁ ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="0eb75-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="0eb75-137">**ItemClass**</span></span> | <span data-ttu-id="0eb75-138">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-138">No</span></span> | <span data-ttu-id="0eb75-p103">Especifica a classe de mensagens personalizada para fazer a correspondência. Para saber mais, confira [Ativar um suplemento de email no Outlook para uma classe de mensagens específica](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="0eb75-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="0eb75-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="0eb75-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="0eb75-142">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-142">No</span></span> | <span data-ttu-id="0eb75-143">Especifica se a regra deve ser avaliada como true se o item pertencer a uma subclasse da classe de mensagens especificada. O padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="0eb75-144">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0eb75-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="0eb75-145">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="0eb75-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="0eb75-146">Define uma regra que é avaliada como true se o item contiver um anexo.</span><span class="sxs-lookup"><span data-stu-id="0eb75-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="0eb75-147">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0eb75-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="0eb75-148">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="0eb75-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="0eb75-149">Define uma regra que é avaliada como true se o item contiver texto do tipo de entidade especificada em seu assunto ou corpo.</span><span class="sxs-lookup"><span data-stu-id="0eb75-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="0eb75-150">Atributos</span><span class="sxs-lookup"><span data-stu-id="0eb75-150">Attributes</span></span>

| <span data-ttu-id="0eb75-151">Atributo</span><span class="sxs-lookup"><span data-stu-id="0eb75-151">Attribute</span></span> | <span data-ttu-id="0eb75-152">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0eb75-152">Required</span></span> | <span data-ttu-id="0eb75-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="0eb75-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="0eb75-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="0eb75-154">**EntityType**</span></span> | <span data-ttu-id="0eb75-155">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-155">Yes</span></span> | <span data-ttu-id="0eb75-p104">Especifica o tipo de entidade que deve ser encontrado para a regra para que ela seja avaliada como true. Pode ser um dos seguintes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="0eb75-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="0eb75-158">**RegExFilter**</span></span> | <span data-ttu-id="0eb75-159">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-159">No</span></span> | <span data-ttu-id="0eb75-160">Especifica uma expressão regular para executar esta entidade para ativação.</span><span class="sxs-lookup"><span data-stu-id="0eb75-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="0eb75-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="0eb75-161">**FilterName**</span></span> | <span data-ttu-id="0eb75-162">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-162">No</span></span> | <span data-ttu-id="0eb75-163">Especifica o nome do filtro de expressões regulares para que seja possível consultá-lo posteriormente no código do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="0eb75-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="0eb75-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="0eb75-164">**IgnoreCase**</span></span> | <span data-ttu-id="0eb75-165">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-165">No</span></span> | <span data-ttu-id="0eb75-166">Especifica para ignorar maiúsculas e minúsculas ao executar a expressão regular especificada pelo atributo **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="0eb75-166">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="0eb75-167">**Realce**</span><span class="sxs-lookup"><span data-stu-id="0eb75-167">**Highlight**</span></span> | <span data-ttu-id="0eb75-168">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-168">No</span></span> | <span data-ttu-id="0eb75-p105">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar entidades correspondentes. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="0eb75-173">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0eb75-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="0eb75-174">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="0eb75-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="0eb75-175">Define uma regra que é avaliada como true se uma correspondência para a expressão regular especificada pode ser encontrada na propriedade especificada do item.</span><span class="sxs-lookup"><span data-stu-id="0eb75-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="0eb75-176">Atributos</span><span class="sxs-lookup"><span data-stu-id="0eb75-176">Attributes</span></span>

| <span data-ttu-id="0eb75-177">Atributo</span><span class="sxs-lookup"><span data-stu-id="0eb75-177">Attribute</span></span> | <span data-ttu-id="0eb75-178">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0eb75-178">Required</span></span> | <span data-ttu-id="0eb75-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="0eb75-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="0eb75-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="0eb75-180">**RegExName**</span></span> | <span data-ttu-id="0eb75-181">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-181">Yes</span></span> | <span data-ttu-id="0eb75-182">Especifica o nome da expressão regular para que você possa fazer referência à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="0eb75-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="0eb75-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="0eb75-183">**RegExValue**</span></span> | <span data-ttu-id="0eb75-184">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-184">Yes</span></span> | <span data-ttu-id="0eb75-185">Especifica a expressão regular que será avaliada para determinar se o suplemento de email deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="0eb75-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="0eb75-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="0eb75-186">**PropertyName**</span></span> | <span data-ttu-id="0eb75-187">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-187">Yes</span></span> | <span data-ttu-id="0eb75-p106">Especifica o nome da propriedade em relação a qual expressão regular será avaliada. Pode ser um dos seguintes: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span> |
| <span data-ttu-id="0eb75-190">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="0eb75-190">**IgnoreCase**</span></span> | <span data-ttu-id="0eb75-191">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-191">No</span></span> | <span data-ttu-id="0eb75-192">Especifica que as maiúsculas e minúsculas devem ser ignoradas ao executar a expressão regular.</span><span class="sxs-lookup"><span data-stu-id="0eb75-192">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="0eb75-193">**Realce**</span><span class="sxs-lookup"><span data-stu-id="0eb75-193">**Highlight**</span></span> | <span data-ttu-id="0eb75-194">Não</span><span class="sxs-lookup"><span data-stu-id="0eb75-194">No</span></span> | <span data-ttu-id="0eb75-p107">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar texto correspondente. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="0eb75-199">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0eb75-199">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="0eb75-200">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="0eb75-200">RuleCollection</span></span>

<span data-ttu-id="0eb75-201">Define uma coleção de regras e o operador lógico a ser usado ao avaliá-las.</span><span class="sxs-lookup"><span data-stu-id="0eb75-201">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="0eb75-202">Atributos</span><span class="sxs-lookup"><span data-stu-id="0eb75-202">Attributes</span></span>

| <span data-ttu-id="0eb75-203">Atributo</span><span class="sxs-lookup"><span data-stu-id="0eb75-203">Attribute</span></span> | <span data-ttu-id="0eb75-204">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0eb75-204">Required</span></span> | <span data-ttu-id="0eb75-205">Descrição</span><span class="sxs-lookup"><span data-stu-id="0eb75-205">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="0eb75-206">**Mode**</span><span class="sxs-lookup"><span data-stu-id="0eb75-206">**Mode**</span></span> | <span data-ttu-id="0eb75-207">Sim</span><span class="sxs-lookup"><span data-stu-id="0eb75-207">Yes</span></span> | <span data-ttu-id="0eb75-p108">Especifica o operador lógico a ser usado quando estiver avaliando essa coleção de regras. Pode ser: `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="0eb75-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="0eb75-210">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0eb75-210">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="0eb75-211">Confira também</span><span class="sxs-lookup"><span data-stu-id="0eb75-211">See also</span></span>

- [<span data-ttu-id="0eb75-212">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="0eb75-212">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="0eb75-213">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="0eb75-213">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="0eb75-214">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="0eb75-214">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)