# <a name="rule-element"></a><span data-ttu-id="29cc5-101">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="29cc5-101">Rule element</span></span>

<span data-ttu-id="29cc5-102">Especifica a(s) regra(s) de ativação que deve(m) ser avaliada(s) para este suplemento contextual de email.</span><span class="sxs-lookup"><span data-stu-id="29cc5-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="29cc5-103">**Tipo de suplemento:** Suplemento contextual de email</span><span class="sxs-lookup"><span data-stu-id="29cc5-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="29cc5-104">Contido em</span><span class="sxs-lookup"><span data-stu-id="29cc5-104">Contained in</span></span>

- [<span data-ttu-id="29cc5-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="29cc5-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="29cc5-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="29cc5-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="29cc5-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="29cc5-107">Attributes</span></span>

| <span data-ttu-id="29cc5-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="29cc5-108">Attribute</span></span> | <span data-ttu-id="29cc5-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="29cc5-109">Required</span></span> | <span data-ttu-id="29cc5-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="29cc5-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="29cc5-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="29cc5-111">**xsi:type**</span></span> | <span data-ttu-id="29cc5-112">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-112">Yes</span></span> | <span data-ttu-id="29cc5-113">O tipo de regra que está sendo definida.</span><span class="sxs-lookup"><span data-stu-id="29cc5-113">The type of rule being defined.</span></span> |

<span data-ttu-id="29cc5-114">O tipo de regra pode ser um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="29cc5-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="29cc5-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="29cc5-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="29cc5-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="29cc5-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="29cc5-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="29cc5-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="29cc5-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="29cc5-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="29cc5-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="29cc5-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="29cc5-120">Regra ItemIs</span><span class="sxs-lookup"><span data-stu-id="29cc5-120">ItemIs rule</span></span>

<span data-ttu-id="29cc5-121">Define uma regra que é avaliada como true se o item selecionado for do tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="29cc5-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="29cc5-122">Atributos</span><span class="sxs-lookup"><span data-stu-id="29cc5-122">Attributes</span></span>

| <span data-ttu-id="29cc5-123">Atributo</span><span class="sxs-lookup"><span data-stu-id="29cc5-123">Attribute</span></span> | <span data-ttu-id="29cc5-124">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="29cc5-124">Required</span></span> | <span data-ttu-id="29cc5-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="29cc5-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="29cc5-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="29cc5-126">**ItemType**</span></span> | <span data-ttu-id="29cc5-127">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-127">Yes</span></span> | <span data-ttu-id="29cc5-p101">Especifica o tipo de item para fazer a correspondência. Pode ser `Message` ou `Appointment`. O tipo de item `Message` inclui email, solicitações de reunião, respostas a reuniões e cancelamentos de reuniões.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="29cc5-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="29cc5-131">**FormType**</span></span> | <span data-ttu-id="29cc5-132">Não (dentro de [ExtensionPoint](extensionpoint.md)), Sim (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="29cc5-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="29cc5-p102">Especifica se o aplicativo deve aparecer no formulário de leitura ou edição do item. Pode ser um dos seguintes: `Read`, `Edit`, `ReadOrEdit`. Se não for especificado em um `Rule` dentro de um `ExtensionPoint`, esse valor DEVERÁ ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="29cc5-136">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="29cc5-136">**ItemClass**</span></span> | <span data-ttu-id="29cc5-137">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-137">No</span></span> | <span data-ttu-id="29cc5-p103">Especifica a classe de mensagens personalizada para fazer a correspondência. Para saber mais, confira [Ativar um suplemento de email no Outlook para uma classe de mensagens específica](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="29cc5-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="29cc5-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="29cc5-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="29cc5-141">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-141">No</span></span> | <span data-ttu-id="29cc5-142">Especifica se a regra deve ser avaliada como true se o item pertencer a uma subclasse da classe de mensagens especificada. O padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="29cc5-143">Exemplo</span><span class="sxs-lookup"><span data-stu-id="29cc5-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="29cc5-144">Regra ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="29cc5-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="29cc5-145">Define uma regra que é avaliada como true se o item contiver um anexo.</span><span class="sxs-lookup"><span data-stu-id="29cc5-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="29cc5-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="29cc5-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="29cc5-147">Regra ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="29cc5-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="29cc5-148">Define uma regra que é avaliada como true se o item contiver texto do tipo de entidade especificada em seu assunto ou corpo.</span><span class="sxs-lookup"><span data-stu-id="29cc5-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="29cc5-149">Atributos</span><span class="sxs-lookup"><span data-stu-id="29cc5-149">Attributes</span></span>

| <span data-ttu-id="29cc5-150">Atributo</span><span class="sxs-lookup"><span data-stu-id="29cc5-150">Attribute</span></span> | <span data-ttu-id="29cc5-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="29cc5-151">Required</span></span> | <span data-ttu-id="29cc5-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="29cc5-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="29cc5-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="29cc5-153">**EntityType**</span></span> | <span data-ttu-id="29cc5-154">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-154">Yes</span></span> | <span data-ttu-id="29cc5-p104">Especifica o tipo de entidade que deve ser encontrado para a regra para que ela seja avaliada como true. Pode ser um dos seguintes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` ou `Contact`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="29cc5-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="29cc5-157">**RegExFilter**</span></span> | <span data-ttu-id="29cc5-158">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-158">No</span></span> | <span data-ttu-id="29cc5-159">Especifica uma expressão regular para executar esta entidade para ativação.</span><span class="sxs-lookup"><span data-stu-id="29cc5-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="29cc5-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="29cc5-160">**FilterName**</span></span> | <span data-ttu-id="29cc5-161">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-161">No</span></span> | <span data-ttu-id="29cc5-162">Especifica o nome do filtro de expressões regulares para que seja possível consultá-lo posteriormente no código do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="29cc5-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="29cc5-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="29cc5-163">**IgnoreCase**</span></span> | <span data-ttu-id="29cc5-164">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-164">No</span></span> | <span data-ttu-id="29cc5-165">Especifica para ignorar maiúsculas e minúsculas ao executar a expressão regular especificada pelo atributo **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="29cc5-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="29cc5-166">**Realce**</span><span class="sxs-lookup"><span data-stu-id="29cc5-166">**Highlight**</span></span> | <span data-ttu-id="29cc5-167">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-167">No</span></span> | <span data-ttu-id="29cc5-p105">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar entidades correspondentes. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="29cc5-172">Exemplo</span><span class="sxs-lookup"><span data-stu-id="29cc5-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="29cc5-173">Regra ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="29cc5-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="29cc5-174">Define uma regra que é avaliada como true se uma correspondência para a expressão regular especificada pode ser encontrada na propriedade especificada do item.</span><span class="sxs-lookup"><span data-stu-id="29cc5-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="29cc5-175">Atributos</span><span class="sxs-lookup"><span data-stu-id="29cc5-175">Attributes</span></span>

| <span data-ttu-id="29cc5-176">Atributo</span><span class="sxs-lookup"><span data-stu-id="29cc5-176">Attribute</span></span> | <span data-ttu-id="29cc5-177">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="29cc5-177">Required</span></span> | <span data-ttu-id="29cc5-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="29cc5-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="29cc5-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="29cc5-179">**RegExName**</span></span> | <span data-ttu-id="29cc5-180">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-180">Yes</span></span> | <span data-ttu-id="29cc5-181">Especifica o nome da expressão regular para que você possa fazer referência à expressão no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="29cc5-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="29cc5-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="29cc5-182">**RegExValue**</span></span> | <span data-ttu-id="29cc5-183">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-183">Yes</span></span> | <span data-ttu-id="29cc5-184">Especifica a expressão regular que será avaliada para determinar se o suplemento de email deve ser mostrado.</span><span class="sxs-lookup"><span data-stu-id="29cc5-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="29cc5-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="29cc5-185">**PropertyName**</span></span> | <span data-ttu-id="29cc5-186">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-186">Yes</span></span> | <span data-ttu-id="29cc5-p106">Especifica o nome da propriedade em relação a qual expressão regular será avaliada. Pode ser um dos seguintes: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` ou `SenderSTMPAddress`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="29cc5-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="29cc5-189">**IgnoreCase**</span></span> | <span data-ttu-id="29cc5-190">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-190">No</span></span> | <span data-ttu-id="29cc5-191">Especifica que as maiúsculas e minúsculas devem ser ignoradas ao executar a expressão regular.</span><span class="sxs-lookup"><span data-stu-id="29cc5-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="29cc5-192">**Realce**</span><span class="sxs-lookup"><span data-stu-id="29cc5-192">**Highlight**</span></span> | <span data-ttu-id="29cc5-193">Não</span><span class="sxs-lookup"><span data-stu-id="29cc5-193">No</span></span> | <span data-ttu-id="29cc5-p107">**Observação:** isso se aplica somente aos elementos **Rule** dentro dos elementos **ExtensionPoint**. Especifica como o cliente deve realçar texto correspondente. Pode ser um dos seguintes: `all` ou `none`. Se não for especificado, o valor padrão será `all`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="29cc5-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="29cc5-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="29cc5-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="29cc5-199">RuleCollection</span></span>

<span data-ttu-id="29cc5-200">Define uma coleção de regras e o operador lógico a ser usado ao avaliá-las.</span><span class="sxs-lookup"><span data-stu-id="29cc5-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="29cc5-201">Atributos</span><span class="sxs-lookup"><span data-stu-id="29cc5-201">Attributes</span></span>

| <span data-ttu-id="29cc5-202">Atributo</span><span class="sxs-lookup"><span data-stu-id="29cc5-202">Attribute</span></span> | <span data-ttu-id="29cc5-203">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="29cc5-203">Required</span></span> | <span data-ttu-id="29cc5-204">Descrição</span><span class="sxs-lookup"><span data-stu-id="29cc5-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="29cc5-205">**Mode**</span><span class="sxs-lookup"><span data-stu-id="29cc5-205">**Mode**</span></span> | <span data-ttu-id="29cc5-206">Sim</span><span class="sxs-lookup"><span data-stu-id="29cc5-206">Yes</span></span> | <span data-ttu-id="29cc5-p108">Especifica o operador lógico a ser usado quando estiver avaliando essa coleção de regras. Pode ser: `And` ou `Or`.</span><span class="sxs-lookup"><span data-stu-id="29cc5-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="29cc5-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="29cc5-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="29cc5-210">Confira também</span><span class="sxs-lookup"><span data-stu-id="29cc5-210">See also</span></span>

- [<span data-ttu-id="29cc5-211">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="29cc5-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="29cc5-212">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="29cc5-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="29cc5-213">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="29cc5-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)