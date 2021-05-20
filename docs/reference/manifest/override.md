---
title: Elemento Override no arquivo de manifesto
description: O elemento Substituir permite especificar o valor de uma configuração dependendo de uma condição especificada.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555154"
---
# <a name="override-element"></a><span data-ttu-id="7ed95-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="7ed95-103">Override element</span></span>

<span data-ttu-id="7ed95-104">Fornece uma maneira de substituir o valor de uma configuração manifesto, dependendo de uma condição especificada.</span><span class="sxs-lookup"><span data-stu-id="7ed95-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="7ed95-105">Existem três tipos de condições:</span><span class="sxs-lookup"><span data-stu-id="7ed95-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="7ed95-106">Uma Office local que é diferente do `LocaleToken` padrão, chamado **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="7ed95-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="7ed95-107">Um padrão de suporte de conjunto de requisitos diferente do `RequirementToken` padrão padrão, chamado **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="7ed95-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="7ed95-108">A fonte é diferente do padrão `Runtime` , chamado **RuntimeOverride** (atualmente em pré-visualização).</span><span class="sxs-lookup"><span data-stu-id="7ed95-108">The source is different from the default `Runtime`, called **RuntimeOverride** (currently in preview).</span></span>

<span data-ttu-id="7ed95-109">Um `<Override>` elemento que está dentro de um elemento deve ser do tipo `<Runtime>` **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="7ed95-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="7ed95-110">Não há `overrideType` atributo para o `<Override>` elemento.</span><span class="sxs-lookup"><span data-stu-id="7ed95-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="7ed95-111">A diferença é determinada pelo elemento pai e pelo tipo do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="7ed95-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="7ed95-112">Um `<Override>` elemento que está dentro de um elemento cujo é , deve ser do tipo `<Token>` `xsi:type` `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="7ed95-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="7ed95-113">Um `<Override>` elemento dentro de qualquer outro elemento pai, ou dentro de um elemento do `<Override>` `LocaleToken` tipo, deve ser do tipo **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="7ed95-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="7ed95-114">Para obter mais informações sobre o uso desse elemento quando for filho de um `<Token>` elemento, consulte [Trabalho com substituições estendidas do manifesto](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="7ed95-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="7ed95-115">Cada tipo é descrito em seções separadas mais tarde neste artigo.</span><span class="sxs-lookup"><span data-stu-id="7ed95-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="7ed95-116">Elemento de substituição para `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="7ed95-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="7ed95-117">Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração.</span><span class="sxs-lookup"><span data-stu-id="7ed95-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="7ed95-118">Se o `<Override>` elemento é do tipo **LocaleTokenOverride,** então o `Locale` atributo é a condição, e o atributo é `Value` o consequente.</span><span class="sxs-lookup"><span data-stu-id="7ed95-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="7ed95-119">Por exemplo, o seguinte é "Se o Office configuração local for fr-fr, então o nome de exibição é 'Lecteur vidéo'."</span><span class="sxs-lookup"><span data-stu-id="7ed95-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="7ed95-120">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="7ed95-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="7ed95-121">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="7ed95-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="7ed95-122">Contido em</span><span class="sxs-lookup"><span data-stu-id="7ed95-122">Contained in</span></span>

|<span data-ttu-id="7ed95-123">Elemento</span><span class="sxs-lookup"><span data-stu-id="7ed95-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="7ed95-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="7ed95-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="7ed95-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="7ed95-125">Description</span></span>](description.md)|
|[<span data-ttu-id="7ed95-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="7ed95-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="7ed95-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="7ed95-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="7ed95-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="7ed95-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="7ed95-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="7ed95-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="7ed95-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="7ed95-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="7ed95-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="7ed95-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="7ed95-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7ed95-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="7ed95-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="7ed95-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="7ed95-134">Token</span><span class="sxs-lookup"><span data-stu-id="7ed95-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="7ed95-135">Atributos</span><span class="sxs-lookup"><span data-stu-id="7ed95-135">Attributes</span></span>

|<span data-ttu-id="7ed95-136">Atributo</span><span class="sxs-lookup"><span data-stu-id="7ed95-136">Attribute</span></span>|<span data-ttu-id="7ed95-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="7ed95-137">Type</span></span>|<span data-ttu-id="7ed95-138">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-138">Required</span></span>|<span data-ttu-id="7ed95-139">Descrição</span><span class="sxs-lookup"><span data-stu-id="7ed95-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7ed95-140">Locale</span><span class="sxs-lookup"><span data-stu-id="7ed95-140">Locale</span></span>|<span data-ttu-id="7ed95-141">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7ed95-141">string</span></span>|<span data-ttu-id="7ed95-142">obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-142">required</span></span>|<span data-ttu-id="7ed95-143">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="7ed95-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="7ed95-144">Valor</span><span class="sxs-lookup"><span data-stu-id="7ed95-144">Value</span></span>|<span data-ttu-id="7ed95-145">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7ed95-145">string</span></span>|<span data-ttu-id="7ed95-146">obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-146">required</span></span>|<span data-ttu-id="7ed95-147">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="7ed95-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="7ed95-148">Exemplos</span><span class="sxs-lookup"><span data-stu-id="7ed95-148">Examples</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="7ed95-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="7ed95-149">See also</span></span>

- [<span data-ttu-id="7ed95-150">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7ed95-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="7ed95-151">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="7ed95-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="7ed95-152">Elemento de substituição para `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="7ed95-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="7ed95-153">Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração.</span><span class="sxs-lookup"><span data-stu-id="7ed95-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="7ed95-154">Se o `<Override>` elemento é do tipo **RequirementTokenOverride**, então o elemento criança `<Requirements>` expressa a condição, e o atributo é o `Value` consequente.</span><span class="sxs-lookup"><span data-stu-id="7ed95-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="7ed95-155">Por exemplo, o primeiro `<Override>` a seguir é lido "Se a plataforma atual suporta o FeatureOne versão 1.7, em seguida, use string 'oldAddinVersion' no lugar do token na URL do avô `${token.requirements}` `<ExtendedOverrides>` (em vez da 'atualização' de string padrão)."</span><span class="sxs-lookup"><span data-stu-id="7ed95-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

<span data-ttu-id="7ed95-156">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7ed95-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="7ed95-157">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="7ed95-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="7ed95-158">Contido em</span><span class="sxs-lookup"><span data-stu-id="7ed95-158">Contained in</span></span>

|<span data-ttu-id="7ed95-159">Elemento</span><span class="sxs-lookup"><span data-stu-id="7ed95-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="7ed95-160">Token</span><span class="sxs-lookup"><span data-stu-id="7ed95-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="7ed95-161">Deve conter</span><span class="sxs-lookup"><span data-stu-id="7ed95-161">Must contain</span></span>

|<span data-ttu-id="7ed95-162">Elemento</span><span class="sxs-lookup"><span data-stu-id="7ed95-162">Element</span></span>|<span data-ttu-id="7ed95-163">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="7ed95-163">Content</span></span>|<span data-ttu-id="7ed95-164">Correio</span><span class="sxs-lookup"><span data-stu-id="7ed95-164">Mail</span></span>|<span data-ttu-id="7ed95-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="7ed95-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="7ed95-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7ed95-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="7ed95-167">x</span><span class="sxs-lookup"><span data-stu-id="7ed95-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="7ed95-168">Atributos</span><span class="sxs-lookup"><span data-stu-id="7ed95-168">Attributes</span></span>

|<span data-ttu-id="7ed95-169">Atributo</span><span class="sxs-lookup"><span data-stu-id="7ed95-169">Attribute</span></span>|<span data-ttu-id="7ed95-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="7ed95-170">Type</span></span>|<span data-ttu-id="7ed95-171">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-171">Required</span></span>|<span data-ttu-id="7ed95-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="7ed95-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7ed95-173">Valor</span><span class="sxs-lookup"><span data-stu-id="7ed95-173">Value</span></span>|<span data-ttu-id="7ed95-174">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7ed95-174">string</span></span>|<span data-ttu-id="7ed95-175">obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-175">required</span></span>|<span data-ttu-id="7ed95-176">Valor do token avó quando a condição está satisfeita.</span><span class="sxs-lookup"><span data-stu-id="7ed95-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="7ed95-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7ed95-177">Example</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="7ed95-178">Confira também</span><span class="sxs-lookup"><span data-stu-id="7ed95-178">See also</span></span>

- [<span data-ttu-id="7ed95-179">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="7ed95-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="7ed95-180">Definir o elemento Requirements no manifesto</span><span class="sxs-lookup"><span data-stu-id="7ed95-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="7ed95-181">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="7ed95-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime-preview"></a><span data-ttu-id="7ed95-182">Elemento de substituição para `Runtime` (visualização)</span><span class="sxs-lookup"><span data-stu-id="7ed95-182">Override element for `Runtime` (preview)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7ed95-183">Este recurso só é suportado para [visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) em Outlook na web e em Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="7ed95-183">This feature is only supported for [preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="7ed95-184">Para obter mais detalhes, consulte [Configurar seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="7ed95-184">For more details, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>
>
> <span data-ttu-id="7ed95-185">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.</span><span class="sxs-lookup"><span data-stu-id="7ed95-185">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

<span data-ttu-id="7ed95-186">Um `<Override>` elemento expressa um condicional e pode ser lido como um "Se ... então..." declaração.</span><span class="sxs-lookup"><span data-stu-id="7ed95-186">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="7ed95-187">Se o `<Override>` elemento é do tipo **RuntimeOverride,** então o `type` atributo é a condição, e o atributo é `resid` o consequente.</span><span class="sxs-lookup"><span data-stu-id="7ed95-187">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="7ed95-188">Por exemplo, o seguinte é "Se o tipo é 'javascript', então o `resid` é 'JSRuntime.Url'." Outlook A área de trabalho requer esse elemento para manipuladores [de pontos de extensão LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="7ed95-188">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="7ed95-189">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="7ed95-189">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="7ed95-190">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="7ed95-190">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="7ed95-191">Contido em</span><span class="sxs-lookup"><span data-stu-id="7ed95-191">Contained in</span></span>

- [<span data-ttu-id="7ed95-192">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="7ed95-192">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="7ed95-193">Atributos</span><span class="sxs-lookup"><span data-stu-id="7ed95-193">Attributes</span></span>

|<span data-ttu-id="7ed95-194">Atributo</span><span class="sxs-lookup"><span data-stu-id="7ed95-194">Attribute</span></span>|<span data-ttu-id="7ed95-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="7ed95-195">Type</span></span>|<span data-ttu-id="7ed95-196">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7ed95-196">Required</span></span>|<span data-ttu-id="7ed95-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="7ed95-197">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7ed95-198">**type**</span><span class="sxs-lookup"><span data-stu-id="7ed95-198">**type**</span></span>|<span data-ttu-id="7ed95-199">string</span><span class="sxs-lookup"><span data-stu-id="7ed95-199">string</span></span>|<span data-ttu-id="7ed95-200">Sim</span><span class="sxs-lookup"><span data-stu-id="7ed95-200">Yes</span></span>|<span data-ttu-id="7ed95-201">Especifica o idioma para esta substituição.</span><span class="sxs-lookup"><span data-stu-id="7ed95-201">Specifies the language for this override.</span></span> <span data-ttu-id="7ed95-202">No momento, `"javascript"` é a única opção suportada.</span><span class="sxs-lookup"><span data-stu-id="7ed95-202">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="7ed95-203">**resid**</span><span class="sxs-lookup"><span data-stu-id="7ed95-203">**resid**</span></span>|<span data-ttu-id="7ed95-204">string</span><span class="sxs-lookup"><span data-stu-id="7ed95-204">string</span></span>|<span data-ttu-id="7ed95-205">Sim</span><span class="sxs-lookup"><span data-stu-id="7ed95-205">Yes</span></span>|<span data-ttu-id="7ed95-206">Especifica a localização do URL do arquivo JavaScript que deve substituir a localização url do HTML padrão definido no elemento [Runtime](runtime.md) do pai `resid` .</span><span class="sxs-lookup"><span data-stu-id="7ed95-206">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="7ed95-207">O `resid` pode não ter mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="7ed95-207">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="7ed95-208">Exemplos</span><span class="sxs-lookup"><span data-stu-id="7ed95-208">Examples</span></span>

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a><span data-ttu-id="7ed95-209">Confira também</span><span class="sxs-lookup"><span data-stu-id="7ed95-209">See also</span></span>

- [<span data-ttu-id="7ed95-210">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="7ed95-210">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="7ed95-211">Configure seu Outlook complemento para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="7ed95-211">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
