---
title: Elemento Override no arquivo de manifesto
description: O elemento Override permite que você especifique o valor de uma configuração dependendo de uma condição especificada.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd270fa19750810238b42c26c2abc35a61c1bac8
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590901"
---
# <a name="override-element"></a><span data-ttu-id="ea060-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="ea060-103">Override element</span></span>

<span data-ttu-id="ea060-104">Fornece uma maneira de substituir o valor de uma configuração de manifesto, dependendo de uma condição especificada.</span><span class="sxs-lookup"><span data-stu-id="ea060-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="ea060-105">Há três tipos de condições:</span><span class="sxs-lookup"><span data-stu-id="ea060-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="ea060-106">Uma Office local diferente do padrão `LocaleToken` , chamado **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="ea060-107">Um padrão de suporte ao conjunto de requisitos diferente do `RequirementToken` padrão, chamado **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="ea060-108">A origem é diferente do `Runtime` padrão , chamado **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-108">The source is different from the default `Runtime`, called **RuntimeOverride**.</span></span>

<span data-ttu-id="ea060-109">Um `<Override>` elemento que está dentro de um elemento deve ser do tipo `<Runtime>` **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="ea060-110">Não há `overrideType` atributo para o `<Override>` elemento.</span><span class="sxs-lookup"><span data-stu-id="ea060-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="ea060-111">A diferença é determinada pelo elemento pai e pelo tipo do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="ea060-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="ea060-112">Um elemento que está dentro de um elemento cujo é , deve ser do `<Override>` `<Token>` tipo `xsi:type` `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="ea060-113">Um elemento dentro de qualquer outro elemento pai ou dentro de um elemento de tipo deve ser do `<Override>` `<Override>` tipo `LocaleToken` **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="ea060-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="ea060-114">Para obter mais informações sobre o uso desse elemento quando ele é filho de um elemento, consulte Trabalhar com substituições `<Token>` [estendidas do manifesto](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="ea060-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="ea060-115">Cada tipo é descrito em seções separadas posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="ea060-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="ea060-116">Elemento Override para `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="ea060-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="ea060-117">Um `<Override>` elemento expressa uma condição e pode ser lido como um "Se ... then ..." instrução.</span><span class="sxs-lookup"><span data-stu-id="ea060-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ea060-118">Se o `<Override>` elemento for do tipo **LocaleTokenOverride**, o atributo será a condição e o `Locale` atributo será o `Value` conseqüente.</span><span class="sxs-lookup"><span data-stu-id="ea060-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ea060-119">Por exemplo, o seguinte é lido "Se a configuração de Office local for fr-fr, o nome para exibição será 'Lecteur vidéo'."</span><span class="sxs-lookup"><span data-stu-id="ea060-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="ea060-120">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="ea060-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="ea060-121">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ea060-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="ea060-122">Contido em</span><span class="sxs-lookup"><span data-stu-id="ea060-122">Contained in</span></span>

|<span data-ttu-id="ea060-123">Elemento</span><span class="sxs-lookup"><span data-stu-id="ea060-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="ea060-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="ea060-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ea060-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea060-125">Description</span></span>](description.md)|
|[<span data-ttu-id="ea060-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ea060-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ea060-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ea060-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ea060-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ea060-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ea060-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ea060-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ea060-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ea060-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ea060-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ea060-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ea060-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ea060-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ea060-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ea060-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="ea060-134">Token</span><span class="sxs-lookup"><span data-stu-id="ea060-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="ea060-135">Atributos</span><span class="sxs-lookup"><span data-stu-id="ea060-135">Attributes</span></span>

|<span data-ttu-id="ea060-136">Atributo</span><span class="sxs-lookup"><span data-stu-id="ea060-136">Attribute</span></span>|<span data-ttu-id="ea060-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="ea060-137">Type</span></span>|<span data-ttu-id="ea060-138">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-138">Required</span></span>|<span data-ttu-id="ea060-139">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea060-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ea060-140">Locale</span><span class="sxs-lookup"><span data-stu-id="ea060-140">Locale</span></span>|<span data-ttu-id="ea060-141">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ea060-141">string</span></span>|<span data-ttu-id="ea060-142">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-142">required</span></span>|<span data-ttu-id="ea060-143">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ea060-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ea060-144">Valor</span><span class="sxs-lookup"><span data-stu-id="ea060-144">Value</span></span>|<span data-ttu-id="ea060-145">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ea060-145">string</span></span>|<span data-ttu-id="ea060-146">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-146">required</span></span>|<span data-ttu-id="ea060-147">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="ea060-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="ea060-148">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ea060-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ea060-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="ea060-149">See also</span></span>

- [<span data-ttu-id="ea060-150">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ea060-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="ea060-151">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="ea060-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="ea060-152">Elemento Override para `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="ea060-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="ea060-153">Um `<Override>` elemento expressa uma condição e pode ser lido como um "Se ... then ..." instrução.</span><span class="sxs-lookup"><span data-stu-id="ea060-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ea060-154">Se o `<Override>` elemento for do tipo **RequirementTokenOverride**, o elemento filho expressará a condição `<Requirements>` e o atributo será o `Value` conseqüente.</span><span class="sxs-lookup"><span data-stu-id="ea060-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="ea060-155">Por exemplo, o primeiro na seguinte leitura é "Se a plataforma atual dá suporte ao FeatureOne versão 1.7, use a cadeia de caracteres 'oldAddinVersion' no lugar do token na URL do vô-vô (em vez da cadeia de caracteres `<Override>` `${token.requirements}` padrão `<ExtendedOverrides>` 'upgrade')."</span><span class="sxs-lookup"><span data-stu-id="ea060-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="ea060-156">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ea060-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="ea060-157">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ea060-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="ea060-158">Contido em</span><span class="sxs-lookup"><span data-stu-id="ea060-158">Contained in</span></span>

|<span data-ttu-id="ea060-159">Elemento</span><span class="sxs-lookup"><span data-stu-id="ea060-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="ea060-160">Token</span><span class="sxs-lookup"><span data-stu-id="ea060-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="ea060-161">Deve conter</span><span class="sxs-lookup"><span data-stu-id="ea060-161">Must contain</span></span>

|<span data-ttu-id="ea060-162">Elemento</span><span class="sxs-lookup"><span data-stu-id="ea060-162">Element</span></span>|<span data-ttu-id="ea060-163">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ea060-163">Content</span></span>|<span data-ttu-id="ea060-164">Email</span><span class="sxs-lookup"><span data-stu-id="ea060-164">Mail</span></span>|<span data-ttu-id="ea060-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="ea060-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="ea060-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ea060-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="ea060-167">x</span><span class="sxs-lookup"><span data-stu-id="ea060-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="ea060-168">Atributos</span><span class="sxs-lookup"><span data-stu-id="ea060-168">Attributes</span></span>

|<span data-ttu-id="ea060-169">Atributo</span><span class="sxs-lookup"><span data-stu-id="ea060-169">Attribute</span></span>|<span data-ttu-id="ea060-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="ea060-170">Type</span></span>|<span data-ttu-id="ea060-171">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-171">Required</span></span>|<span data-ttu-id="ea060-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea060-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ea060-173">Valor</span><span class="sxs-lookup"><span data-stu-id="ea060-173">Value</span></span>|<span data-ttu-id="ea060-174">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ea060-174">string</span></span>|<span data-ttu-id="ea060-175">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-175">required</span></span>|<span data-ttu-id="ea060-176">Valor do token de vôvão quando a condição for atendida.</span><span class="sxs-lookup"><span data-stu-id="ea060-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="ea060-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ea060-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ea060-178">Confira também</span><span class="sxs-lookup"><span data-stu-id="ea060-178">See also</span></span>

- [<span data-ttu-id="ea060-179">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="ea060-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ea060-180">Definir o elemento Requirements no manifesto</span><span class="sxs-lookup"><span data-stu-id="ea060-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="ea060-181">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="ea060-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a><span data-ttu-id="ea060-182">Elemento Override para `Runtime`</span><span class="sxs-lookup"><span data-stu-id="ea060-182">Override element for `Runtime`</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ea060-183">O suporte a esse elemento foi introduzido no conjunto de requisitos de Caixa de [Correio 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) com o [recurso de ativação baseada em evento.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="ea060-183">Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md).</span></span> <span data-ttu-id="ea060-184">Confira, [clientes e plataformas](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="ea060-184">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="ea060-185">Um `<Override>` elemento expressa uma condição e pode ser lido como um "Se ... then ..." instrução.</span><span class="sxs-lookup"><span data-stu-id="ea060-185">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="ea060-186">Se o `<Override>` elemento for do tipo **RuntimeOverride**, o atributo será a condição e o `type` atributo será o `resid` conseqüente.</span><span class="sxs-lookup"><span data-stu-id="ea060-186">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="ea060-187">Por exemplo, o seguinte é ler "Se o tipo for 'javascript', será `resid` 'JSRuntime.Url'." Outlook A área de trabalho requer esse elemento [para manipuladores de pontos de extensão LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="ea060-187">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="ea060-188">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="ea060-188">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="ea060-189">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ea060-189">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="ea060-190">Contido em</span><span class="sxs-lookup"><span data-stu-id="ea060-190">Contained in</span></span>

- [<span data-ttu-id="ea060-191">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="ea060-191">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="ea060-192">Atributos</span><span class="sxs-lookup"><span data-stu-id="ea060-192">Attributes</span></span>

|<span data-ttu-id="ea060-193">Atributo</span><span class="sxs-lookup"><span data-stu-id="ea060-193">Attribute</span></span>|<span data-ttu-id="ea060-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="ea060-194">Type</span></span>|<span data-ttu-id="ea060-195">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea060-195">Required</span></span>|<span data-ttu-id="ea060-196">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea060-196">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ea060-197">**type**</span><span class="sxs-lookup"><span data-stu-id="ea060-197">**type**</span></span>|<span data-ttu-id="ea060-198">string</span><span class="sxs-lookup"><span data-stu-id="ea060-198">string</span></span>|<span data-ttu-id="ea060-199">Sim</span><span class="sxs-lookup"><span data-stu-id="ea060-199">Yes</span></span>|<span data-ttu-id="ea060-200">Especifica o idioma para essa substituição.</span><span class="sxs-lookup"><span data-stu-id="ea060-200">Specifies the language for this override.</span></span> <span data-ttu-id="ea060-201">No momento, `"javascript"` é a única opção com suporte.</span><span class="sxs-lookup"><span data-stu-id="ea060-201">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="ea060-202">**resid**</span><span class="sxs-lookup"><span data-stu-id="ea060-202">**resid**</span></span>|<span data-ttu-id="ea060-203">string</span><span class="sxs-lookup"><span data-stu-id="ea060-203">string</span></span>|<span data-ttu-id="ea060-204">Sim</span><span class="sxs-lookup"><span data-stu-id="ea060-204">Yes</span></span>|<span data-ttu-id="ea060-205">Especifica o local da URL do arquivo JavaScript que deve substituir o local da URL do HTML padrão definido no elemento [Runtime](runtime.md) `resid` pai.</span><span class="sxs-lookup"><span data-stu-id="ea060-205">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="ea060-206">O `resid` pode ter não mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento.</span><span class="sxs-lookup"><span data-stu-id="ea060-206">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="ea060-207">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ea060-207">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="ea060-208">Confira também</span><span class="sxs-lookup"><span data-stu-id="ea060-208">See also</span></span>

- [<span data-ttu-id="ea060-209">Tempo de execução</span><span class="sxs-lookup"><span data-stu-id="ea060-209">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="ea060-210">Configurar seu Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="ea060-210">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
