---
title: Elemento Rule no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 07037c43c111f735a7354a048066e4c4a88f7637
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871511"
---
# <a name="rule-element"></a><span data-ttu-id="57e3a-102">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="57e3a-102">Rule element</span></span>

<span data-ttu-id="57e3a-103">Especifica a(s) regra(s) de ativação que deve(m) ser avaliada(s) para este suplemento contextual de email.</span><span class="sxs-lookup"><span data-stu-id="57e3a-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="57e3a-104">**Tipo de suplemento:** Suplemento contextual de email</span><span class="sxs-lookup"><span data-stu-id="57e3a-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="57e3a-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="57e3a-105">Contained in</span></span>

- [<span data-ttu-id="57e3a-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="57e3a-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="57e3a-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="57e3a-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="57e3a-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="57e3a-108">Attributes</span></span>

| <span data-ttu-id="57e3a-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="57e3a-109">Attribute</span></span> | <span data-ttu-id="57e3a-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57e3a-110">Required</span></span> | <span data-ttu-id="57e3a-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="57e3a-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="57e3a-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="57e3a-112">**xsi:type**</span></span> | <span data-ttu-id="57e3a-113">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-113">Yes</span></span> | <span data-ttu-id="57e3a-114">O tipo de regra que está sendo definida.</span><span class="sxs-lookup"><span data-stu-id="57e3a-114">The type of rule being defined.</span></span> |

<span data-ttu-id="57e3a-115">O tipo de regra pode ser um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="57e3a-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="57e3a-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="57e3a-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="57e3a-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="57e3a-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="57e3a-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="57e3a-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="57e3a-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="57e3a-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="57e3a-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="57e3a-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="57e3a-121">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="57e3a-121">ItemIs rule</span></span>

<span data-ttu-id="57e3a-122">Define uma regra que é avaliada como true se o item selecionado for do tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="57e3a-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="57e3a-123">Atributos</span><span class="sxs-lookup"><span data-stu-id="57e3a-123">Attributes</span></span>

| <span data-ttu-id="57e3a-124">Atributo</span><span class="sxs-lookup"><span data-stu-id="57e3a-124">Attribute</span></span> | <span data-ttu-id="57e3a-125">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57e3a-125">Required</span></span> | <span data-ttu-id="57e3a-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="57e3a-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="57e3a-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="57e3a-127">**ItemType**</span></span> | <span data-ttu-id="57e3a-128">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-128">Yes</span></span> | <span data-ttu-id="57e3a-p101">Especifica o tipo de item para fazer a correspondência. Pode ser `Message` ou `Appointment`. O tipo de item `Message` inclui email, solicitações de reunião, respostas de reunião e cancelamentos de reunião.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="57e3a-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="57e3a-132">**FormType**</span></span> | <span data-ttu-id="57e3a-133">Não (dentro de [ExtensionPoint](extensionpoint.md)), Sim (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="57e3a-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="57e3a-p102">Especifica se o aplicativo deve aparecer no formulário de leitura ou edição do item. Pode ser um dos seguintes: `Read`, `Edit`, `ReadOrEdit`. Se não for especificado em um `Rule` dentro de um `ExtensionPoint`, esse valor DEVERÁ ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="57e3a-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="57e3a-137">**ItemClass**</span></span> | <span data-ttu-id="57e3a-138">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-138">No</span></span> | <span data-ttu-id="57e3a-p103">Especifica a classe de mensagens personalizada para fazer a correspondência. Para saber mais, confira o artigo [Ativar um suplemento de email no Outlook para uma classe de mensagens específica](/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="57e3a-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="57e3a-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="57e3a-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="57e3a-142">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-142">No</span></span> | <span data-ttu-id="57e3a-143">Especifica se a regra deve ser avaliada como true se o item pertencer a uma subclasse da classe de mensagens especificada. O padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="57e3a-144">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57e3a-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="57e3a-145">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="57e3a-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="57e3a-146">Define uma regra que é avaliada como true se o item contiver um anexo.</span><span class="sxs-lookup"><span data-stu-id="57e3a-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="57e3a-147">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57e3a-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="57e3a-148">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="57e3a-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="57e3a-149">Define uma regra que é avaliada como true se o item contiver texto do tipo de entidade especificada em seu assunto ou corpo.</span><span class="sxs-lookup"><span data-stu-id="57e3a-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="57e3a-150">Atributos</span><span class="sxs-lookup"><span data-stu-id="57e3a-150">Attributes</span></span>

| <span data-ttu-id="57e3a-151">Atributo</span><span class="sxs-lookup"><span data-stu-id="57e3a-151">Attribute</span></span> | <span data-ttu-id="57e3a-152">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57e3a-152">Required</span></span> | <span data-ttu-id="57e3a-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="57e3a-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="57e3a-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="57e3a-154">**EntityType**</span></span> | <span data-ttu-id="57e3a-155">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-155">Yes</span></span> | <span data-ttu-id="57e3a-p104">Especifica o tipo de entidade que deve ser encontrado para que a regra para que ela seja avaliada como true. Pode ser um dos seguintes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="57e3a-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="57e3a-158">**RegExFilter**</span></span> | <span data-ttu-id="57e3a-159">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-159">No</span></span> | <span data-ttu-id="57e3a-160">Especifica uma expressão regular para executar esta entidade para ativação.</span><span class="sxs-lookup"><span data-stu-id="57e3a-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="57e3a-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="57e3a-161">**FilterName**</span></span> | <span data-ttu-id="57e3a-162">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-162">No</span></span> | <span data-ttu-id="57e3a-163">Especifica o nome do filtro de expressões regulares para que seja possível consultá-lo posteriormente no código do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="57e3a-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="57e3a-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="57e3a-164">**IgnoreCase**</span></span> | <span data-ttu-id="57e3a-165">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-165">No</span></span> | <span data-ttu-id="57e3a-166">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada por **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="57e3a-166">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="57e3a-167">**Realce**</span><span class="sxs-lookup"><span data-stu-id="57e3a-167">**Highlight**</span></span> | <span data-ttu-id="57e3a-168">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-168">No</span></span> | <span data-ttu-id="57e3a-p105">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar entidades correspondentes. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="57e3a-173">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57e3a-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="57e3a-174">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="57e3a-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="57e3a-175">Define uma regra que é avaliada como true se uma correspondência para a expressão regular especificada pode ser encontrada na propriedade especificada do item.</span><span class="sxs-lookup"><span data-stu-id="57e3a-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="57e3a-176">Atributos</span><span class="sxs-lookup"><span data-stu-id="57e3a-176">Attributes</span></span>

| <span data-ttu-id="57e3a-177">Atributo</span><span class="sxs-lookup"><span data-stu-id="57e3a-177">Attribute</span></span> | <span data-ttu-id="57e3a-178">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57e3a-178">Required</span></span> | <span data-ttu-id="57e3a-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="57e3a-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="57e3a-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="57e3a-180">**RegExName**</span></span> | <span data-ttu-id="57e3a-181">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-181">Yes</span></span> | <span data-ttu-id="57e3a-182">Especifica o nome da expressão regular para que você possa fazer referência à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="57e3a-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="57e3a-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="57e3a-183">**RegExValue**</span></span> | <span data-ttu-id="57e3a-184">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-184">Yes</span></span> | <span data-ttu-id="57e3a-185">Especifica a expressão regular que será avaliada para determinar se o suplemento de email deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="57e3a-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="57e3a-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="57e3a-186">**PropertyName**</span></span> | <span data-ttu-id="57e3a-187">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-187">Yes</span></span> | <span data-ttu-id="57e3a-p106">Especifica o nome da propriedade em relação a qual expressão regular será avaliada. Pode ser um dos seguintes: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="57e3a-190">Se você especificar `BodyAsHTML`, o Outlook só aplicará a expressão regular se o corpo do item for HTML.</span><span class="sxs-lookup"><span data-stu-id="57e3a-190">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="57e3a-191">Caso contrário, o Outlook não retornará nenhuma correspondência para essa expressão regular.</span><span class="sxs-lookup"><span data-stu-id="57e3a-191">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="57e3a-192">Se você especificar `BodyAsPlaintext`, o Outlook sempre aplicará a expressão regular no corpo do item.</span><span class="sxs-lookup"><span data-stu-id="57e3a-192">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="57e3a-193">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="57e3a-193">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="57e3a-194">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="57e3a-194">**IgnoreCase**</span></span> | <span data-ttu-id="57e3a-195">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-195">No</span></span> | <span data-ttu-id="57e3a-196">Especifica se deve ignorar maiúsculas e minúsculas ao fazer a correspondência da expressão regular especificada pelo atributo **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="57e3a-196">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="57e3a-197">**Realce**</span><span class="sxs-lookup"><span data-stu-id="57e3a-197">**Highlight**</span></span> | <span data-ttu-id="57e3a-198">Não</span><span class="sxs-lookup"><span data-stu-id="57e3a-198">No</span></span> | <span data-ttu-id="57e3a-199">Especifica como o cliente deve realçar texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="57e3a-199">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="57e3a-200">Esse atributo pode ser aplicado apenas à elementos **regra** dentro dos elementos **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="57e3a-200">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="57e3a-201">Pode ser um dos seguintes: `all` ou `none`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-201">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="57e3a-202">Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-202">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="57e3a-203">**Observação:** você deve configurar o atributo **PropertyName** para `BodyAsPlaintext` se você especificar o atributo **realçar** para o elemento **regra**.</span><span class="sxs-lookup"><span data-stu-id="57e3a-203">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="57e3a-204">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57e3a-204">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="57e3a-205">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="57e3a-205">RuleCollection</span></span>

<span data-ttu-id="57e3a-206">Define uma coleção de regras e o operador lógico a ser usado ao avaliá-las.</span><span class="sxs-lookup"><span data-stu-id="57e3a-206">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="57e3a-207">Atributos</span><span class="sxs-lookup"><span data-stu-id="57e3a-207">Attributes</span></span>

| <span data-ttu-id="57e3a-208">Atributo</span><span class="sxs-lookup"><span data-stu-id="57e3a-208">Attribute</span></span> | <span data-ttu-id="57e3a-209">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57e3a-209">Required</span></span> | <span data-ttu-id="57e3a-210">Descrição</span><span class="sxs-lookup"><span data-stu-id="57e3a-210">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="57e3a-211">**Mode**</span><span class="sxs-lookup"><span data-stu-id="57e3a-211">**Mode**</span></span> | <span data-ttu-id="57e3a-212">Sim</span><span class="sxs-lookup"><span data-stu-id="57e3a-212">Yes</span></span> | <span data-ttu-id="57e3a-p109">Especifica o operador lógico a ser usado quando estiver avaliando essa coleção de regras. Pode ser: `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="57e3a-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="57e3a-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57e3a-215">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="57e3a-216">Confira também</span><span class="sxs-lookup"><span data-stu-id="57e3a-216">See also</span></span>

- [<span data-ttu-id="57e3a-217">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="57e3a-217">Activation rules for Outlook add-ins</span></span>](/outlook/add-ins/activation-rules)
- [<span data-ttu-id="57e3a-218">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="57e3a-218">Match strings in an Outlook item as well-known entities</span></span>](/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="57e3a-219">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="57e3a-219">Use regular expression activation rules to show an Outlook add-in</span></span>](/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)
