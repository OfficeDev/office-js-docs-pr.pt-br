---
title: Elemento Override no arquivo de manifesto
description: O elemento Override permite que você especifique o valor de uma configuração dependendo de uma condição especificada.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: d2146cc1f44e829bc78076c8093b2ebf791dc722
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505336"
---
# <a name="override-element"></a><span data-ttu-id="1ea61-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="1ea61-103">Override element</span></span>

<span data-ttu-id="1ea61-104">Fornece uma maneira de substituir o valor de uma configuração de manifesto, dependendo de uma condição especificada.</span><span class="sxs-lookup"><span data-stu-id="1ea61-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="1ea61-105">Há dois tipos de condições:</span><span class="sxs-lookup"><span data-stu-id="1ea61-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="1ea61-106">Uma localidade do Office diferente do padrão.</span><span class="sxs-lookup"><span data-stu-id="1ea61-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="1ea61-107">Um padrão de suporte ao conjunto de requisitos que é diferente do padrão padrão.</span><span class="sxs-lookup"><span data-stu-id="1ea61-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="1ea61-108">Há dois tipos de elementos, um é para substituições de localidade, chamado `<Override>` **LocaleTokenOverride** e o outro para substituições de conjunto de requisitos, chamado **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="1ea61-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride**, and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="1ea61-109">Mas não há parâmetro `type` para o `<Override>` elemento.</span><span class="sxs-lookup"><span data-stu-id="1ea61-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="1ea61-110">A diferença é determinada pelo elemento pai e pelo tipo do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="1ea61-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="1ea61-111">Um elemento que está dentro de um elemento cujo é , deve ser do `<Override>` `<Token>` tipo `xsi:type` `RequirementToken` **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="1ea61-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="1ea61-112">Um elemento dentro de qualquer outro elemento pai ou dentro de um elemento de tipo deve ser do `<Override>` `<Override>` tipo `LocaleToken` **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="1ea61-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="1ea61-113">Cada tipo é descrito em seções separadas abaixo.</span><span class="sxs-lookup"><span data-stu-id="1ea61-113">Each type is described in separate sections below.</span></span> <span data-ttu-id="1ea61-114">Para obter mais informações sobre o uso desse elemento quando ele é filho de um elemento, consulte Trabalhar com substituições `<Token>` [estendidas do manifesto](../../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="1ea61-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="1ea61-115">Substituir elemento do tipo LocaleTokenOverride</span><span class="sxs-lookup"><span data-stu-id="1ea61-115">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="1ea61-116">Um `<Override>` elemento expressa uma condição e pode ser lido como um "Se ... then ..." instrução.</span><span class="sxs-lookup"><span data-stu-id="1ea61-116">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="1ea61-117">Se o `<Override>` elemento for do tipo **LocaleTokenOverride**, o atributo será a condição e o `Locale` atributo será o `Value` conseqüente.</span><span class="sxs-lookup"><span data-stu-id="1ea61-117">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="1ea61-118">Por exemplo, o seguinte é lido "Se a configuração de localidade do Office for fr-fr, o nome de exibição será 'Lecteur vidéo'".</span><span class="sxs-lookup"><span data-stu-id="1ea61-118">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="1ea61-119">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="1ea61-119">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="1ea61-120">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1ea61-120">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="1ea61-121">Contido em</span><span class="sxs-lookup"><span data-stu-id="1ea61-121">Contained in</span></span>

|<span data-ttu-id="1ea61-122">Elemento</span><span class="sxs-lookup"><span data-stu-id="1ea61-122">Element</span></span>|
|:-----|
|[<span data-ttu-id="1ea61-123">CitationText</span><span class="sxs-lookup"><span data-stu-id="1ea61-123">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="1ea61-124">Descrição</span><span class="sxs-lookup"><span data-stu-id="1ea61-124">Description</span></span>](description.md)|
|[<span data-ttu-id="1ea61-125">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="1ea61-125">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="1ea61-126">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="1ea61-126">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="1ea61-127">DisplayName</span><span class="sxs-lookup"><span data-stu-id="1ea61-127">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="1ea61-128">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="1ea61-128">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="1ea61-129">IconUrl</span><span class="sxs-lookup"><span data-stu-id="1ea61-129">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="1ea61-130">QueryUri</span><span class="sxs-lookup"><span data-stu-id="1ea61-130">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="1ea61-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1ea61-131">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="1ea61-132">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="1ea61-132">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="1ea61-133">Token</span><span class="sxs-lookup"><span data-stu-id="1ea61-133">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="1ea61-134">Atributos</span><span class="sxs-lookup"><span data-stu-id="1ea61-134">Attributes</span></span>

|<span data-ttu-id="1ea61-135">Atributo</span><span class="sxs-lookup"><span data-stu-id="1ea61-135">Attribute</span></span>|<span data-ttu-id="1ea61-136">Tipo</span><span class="sxs-lookup"><span data-stu-id="1ea61-136">Type</span></span>|<span data-ttu-id="1ea61-137">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1ea61-137">Required</span></span>|<span data-ttu-id="1ea61-138">Descrição</span><span class="sxs-lookup"><span data-stu-id="1ea61-138">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1ea61-139">Locale</span><span class="sxs-lookup"><span data-stu-id="1ea61-139">Locale</span></span>|<span data-ttu-id="1ea61-140">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1ea61-140">string</span></span>|<span data-ttu-id="1ea61-141">obrigatório</span><span class="sxs-lookup"><span data-stu-id="1ea61-141">required</span></span>|<span data-ttu-id="1ea61-142">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="1ea61-142">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="1ea61-143">Valor</span><span class="sxs-lookup"><span data-stu-id="1ea61-143">Value</span></span>|<span data-ttu-id="1ea61-144">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1ea61-144">string</span></span>|<span data-ttu-id="1ea61-145">obrigatório</span><span class="sxs-lookup"><span data-stu-id="1ea61-145">required</span></span>|<span data-ttu-id="1ea61-146">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="1ea61-146">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="1ea61-147">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1ea61-147">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="1ea61-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="1ea61-148">See also</span></span>

- [<span data-ttu-id="1ea61-149">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1ea61-149">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="1ea61-150">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="1ea61-150">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="1ea61-151">Substituir elemento do tipo RequirementTokenOverride</span><span class="sxs-lookup"><span data-stu-id="1ea61-151">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="1ea61-152">Um `<Override>` elemento expressa uma condição e pode ser lido como um "Se ... then ..." instrução.</span><span class="sxs-lookup"><span data-stu-id="1ea61-152">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="1ea61-153">Se o `<Override>` elemento for do tipo **RequirementTokenOverride**, o elemento filho expressará a condição `<Requirements>` e o atributo será o `Value` conseqüente.</span><span class="sxs-lookup"><span data-stu-id="1ea61-153">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="1ea61-154">Por exemplo, o primeiro na seguinte leitura é "Se a plataforma atual dá suporte ao FeatureOne versão 1.7, use a cadeia de caracteres 'oldAddinVersion' no lugar do token na URL do vô-vô (em vez da cadeia de caracteres `<Override>` `${token.requirements}` padrão `<ExtendedOverrides>` 'upgrade')."</span><span class="sxs-lookup"><span data-stu-id="1ea61-154">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="1ea61-155">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="1ea61-155">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="1ea61-156">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1ea61-156">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="1ea61-157">Contido em</span><span class="sxs-lookup"><span data-stu-id="1ea61-157">Contained in</span></span>

|<span data-ttu-id="1ea61-158">Elemento</span><span class="sxs-lookup"><span data-stu-id="1ea61-158">Element</span></span>|
|:-----|
|[<span data-ttu-id="1ea61-159">Token</span><span class="sxs-lookup"><span data-stu-id="1ea61-159">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="1ea61-160">Deve conter</span><span class="sxs-lookup"><span data-stu-id="1ea61-160">Must contain</span></span>

|<span data-ttu-id="1ea61-161">Elemento</span><span class="sxs-lookup"><span data-stu-id="1ea61-161">Element</span></span>|<span data-ttu-id="1ea61-162">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1ea61-162">Content</span></span>|<span data-ttu-id="1ea61-163">Email</span><span class="sxs-lookup"><span data-stu-id="1ea61-163">Mail</span></span>|<span data-ttu-id="1ea61-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="1ea61-164">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="1ea61-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1ea61-165">Requirements</span></span>](requirements.md)|||<span data-ttu-id="1ea61-166">x</span><span class="sxs-lookup"><span data-stu-id="1ea61-166">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="1ea61-167">Atributos</span><span class="sxs-lookup"><span data-stu-id="1ea61-167">Attributes</span></span>

|<span data-ttu-id="1ea61-168">Atributo</span><span class="sxs-lookup"><span data-stu-id="1ea61-168">Attribute</span></span>|<span data-ttu-id="1ea61-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="1ea61-169">Type</span></span>|<span data-ttu-id="1ea61-170">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1ea61-170">Required</span></span>|<span data-ttu-id="1ea61-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="1ea61-171">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1ea61-172">Valor</span><span class="sxs-lookup"><span data-stu-id="1ea61-172">Value</span></span>|<span data-ttu-id="1ea61-173">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1ea61-173">string</span></span>|<span data-ttu-id="1ea61-174">obrigatório</span><span class="sxs-lookup"><span data-stu-id="1ea61-174">required</span></span>|<span data-ttu-id="1ea61-175">Valor do token de vôvão quando a condição for atendida.</span><span class="sxs-lookup"><span data-stu-id="1ea61-175">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="1ea61-176">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1ea61-176">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="1ea61-177">Confira também</span><span class="sxs-lookup"><span data-stu-id="1ea61-177">See also</span></span>

- [<span data-ttu-id="1ea61-178">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="1ea61-178">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="1ea61-179">Definir o elemento Requirements no manifesto</span><span class="sxs-lookup"><span data-stu-id="1ea61-179">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="1ea61-180">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="1ea61-180">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
