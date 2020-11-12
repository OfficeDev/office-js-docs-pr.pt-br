---
title: Elemento Override no arquivo de manifesto
description: O elemento override permite que você especifique o valor de uma configuração, dependendo de uma condição especificada.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996309"
---
# <a name="override-element"></a><span data-ttu-id="2c8e6-103">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="2c8e6-103">Override element</span></span>

<span data-ttu-id="2c8e6-104">Fornece uma maneira de substituir o valor de uma configuração de manifesto, dependendo de uma condição especificada.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="2c8e6-105">Há dois tipos de condições:</span><span class="sxs-lookup"><span data-stu-id="2c8e6-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="2c8e6-106">Uma localidade do Office diferente do padrão.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="2c8e6-107">Um padrão de suporte ao conjunto de requisitos diferente do padrão padrão.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="2c8e6-108">Há dois tipos de `<Override>` elementos, um é para substituições de localidade, chamado **LocaleTokenOverride** , e o outro para o conjunto de requisitos substitui, chamado **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride** , and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="2c8e6-109">Mas não há nenhum `type` parâmetro para o `<Override>` elemento.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="2c8e6-110">A diferença é determinada pelo elemento pai e o tipo do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="2c8e6-111">Um `<Override>` elemento que está dentro de um `<Token>` elemento `xsi:type` , cujo é `RequirementToken` , deve ser do tipo **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="2c8e6-112">Um `<Override>` elemento dentro de qualquer outro elemento pai ou dentro `<Override>` de um elemento do tipo `LocaleToken` deve ser do tipo **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="2c8e6-113">Cada tipo é descrito em seções separadas abaixo.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-113">Each type is described in separate sections below.</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="2c8e6-114">Elemento override do tipo LocaleTokenOverride</span><span class="sxs-lookup"><span data-stu-id="2c8e6-114">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="2c8e6-115">Um `<Override>` elemento expressa um condicional e pode ser lido como um "If... Then... " instrução.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-115">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="2c8e6-116">Se o `<Override>` elemento for do tipo **LocaleTokenOverride** , o `Locale` atributo será a condição e o `Value` atributo será o consequent.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-116">If the `<Override>` element is of type **LocaleTokenOverride** , then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="2c8e6-117">Por exemplo, o seguinte é lido "se a configuração de localidade do Office for fr-fr, o nome para exibição será" Lecteur Vidéo "."</span><span class="sxs-lookup"><span data-stu-id="2c8e6-117">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="2c8e6-118">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="2c8e6-118">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="2c8e6-119">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2c8e6-119">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="2c8e6-120">Contido em</span><span class="sxs-lookup"><span data-stu-id="2c8e6-120">Contained in</span></span>

|<span data-ttu-id="2c8e6-121">Elemento</span><span class="sxs-lookup"><span data-stu-id="2c8e6-121">Element</span></span>|
|:-----|
|[<span data-ttu-id="2c8e6-122">CitationText</span><span class="sxs-lookup"><span data-stu-id="2c8e6-122">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="2c8e6-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="2c8e6-123">Description</span></span>](description.md)|
|[<span data-ttu-id="2c8e6-124">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="2c8e6-124">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="2c8e6-125">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="2c8e6-125">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="2c8e6-126">DisplayName</span><span class="sxs-lookup"><span data-stu-id="2c8e6-126">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="2c8e6-127">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="2c8e6-127">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="2c8e6-128">IconUrl</span><span class="sxs-lookup"><span data-stu-id="2c8e6-128">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="2c8e6-129">QueryUri</span><span class="sxs-lookup"><span data-stu-id="2c8e6-129">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="2c8e6-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="2c8e6-130">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="2c8e6-131">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="2c8e6-131">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="2c8e6-132">Token</span><span class="sxs-lookup"><span data-stu-id="2c8e6-132">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="2c8e6-133">Atributos</span><span class="sxs-lookup"><span data-stu-id="2c8e6-133">Attributes</span></span>

|<span data-ttu-id="2c8e6-134">Atributo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-134">Attribute</span></span>|<span data-ttu-id="2c8e6-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-135">Type</span></span>|<span data-ttu-id="2c8e6-136">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="2c8e6-136">Required</span></span>|<span data-ttu-id="2c8e6-137">Descrição</span><span class="sxs-lookup"><span data-stu-id="2c8e6-137">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2c8e6-138">Locale</span><span class="sxs-lookup"><span data-stu-id="2c8e6-138">Locale</span></span>|<span data-ttu-id="2c8e6-139">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2c8e6-139">string</span></span>|<span data-ttu-id="2c8e6-140">obrigatório</span><span class="sxs-lookup"><span data-stu-id="2c8e6-140">required</span></span>|<span data-ttu-id="2c8e6-141">Especifica o nome da cultura da localidade para essa substituição no formato de marca do idioma BCP 47, como `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-141">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="2c8e6-142">Valor</span><span class="sxs-lookup"><span data-stu-id="2c8e6-142">Value</span></span>|<span data-ttu-id="2c8e6-143">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2c8e6-143">string</span></span>|<span data-ttu-id="2c8e6-144">obrigatório</span><span class="sxs-lookup"><span data-stu-id="2c8e6-144">required</span></span>|<span data-ttu-id="2c8e6-145">Especifica o valor da configuração expressa para a localidade especificada.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-145">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="2c8e6-146">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2c8e6-146">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="2c8e6-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="2c8e6-147">See also</span></span>

- [<span data-ttu-id="2c8e6-148">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2c8e6-148">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="2c8e6-149">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="2c8e6-149">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="2c8e6-150">Elemento override do tipo RequirementTokenOverride</span><span class="sxs-lookup"><span data-stu-id="2c8e6-150">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="2c8e6-151">Um `<Override>` elemento expressa um condicional e pode ser lido como um "If... Then... " instrução.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-151">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="2c8e6-152">Se o `<Override>` elemento for do tipo **RequirementTokenOverride** , o elemento filho `<Requirements>` expressará a condição e o `Value` atributo será o consequent.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-152">If the `<Override>` element is of type **RequirementTokenOverride** , then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="2c8e6-153">Por exemplo, o primeiro `<Override>` no seguinte é lido "se a plataforma atual suportar o FeatureOne versão 1,7, use a cadeia de caracteres ' oldAddinVersion ' no lugar do `${token.requirements}` token na URL do avô `<ExtendedOverrides>` (em vez da cadeia de caracteres padrão" upgrade ")."</span><span class="sxs-lookup"><span data-stu-id="2c8e6-153">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="2c8e6-154">**Tipo de suplemento:** Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="2c8e6-154">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="2c8e6-155">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2c8e6-155">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="2c8e6-156">Contido em</span><span class="sxs-lookup"><span data-stu-id="2c8e6-156">Contained in</span></span>

|<span data-ttu-id="2c8e6-157">Elemento</span><span class="sxs-lookup"><span data-stu-id="2c8e6-157">Element</span></span>|
|:-----|
|[<span data-ttu-id="2c8e6-158">Token</span><span class="sxs-lookup"><span data-stu-id="2c8e6-158">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="2c8e6-159">Deve conter</span><span class="sxs-lookup"><span data-stu-id="2c8e6-159">Must contain</span></span>

|<span data-ttu-id="2c8e6-160">Elemento</span><span class="sxs-lookup"><span data-stu-id="2c8e6-160">Element</span></span>|<span data-ttu-id="2c8e6-161">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-161">Content</span></span>|<span data-ttu-id="2c8e6-162">Email</span><span class="sxs-lookup"><span data-stu-id="2c8e6-162">Mail</span></span>|<span data-ttu-id="2c8e6-163">TaskPane</span><span class="sxs-lookup"><span data-stu-id="2c8e6-163">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2c8e6-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2c8e6-164">Requirements</span></span>](requirements.md)|||<span data-ttu-id="2c8e6-165">x</span><span class="sxs-lookup"><span data-stu-id="2c8e6-165">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="2c8e6-166">Atributos</span><span class="sxs-lookup"><span data-stu-id="2c8e6-166">Attributes</span></span>

|<span data-ttu-id="2c8e6-167">Atributo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-167">Attribute</span></span>|<span data-ttu-id="2c8e6-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-168">Type</span></span>|<span data-ttu-id="2c8e6-169">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="2c8e6-169">Required</span></span>|<span data-ttu-id="2c8e6-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="2c8e6-170">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2c8e6-171">Valor</span><span class="sxs-lookup"><span data-stu-id="2c8e6-171">Value</span></span>|<span data-ttu-id="2c8e6-172">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2c8e6-172">string</span></span>|<span data-ttu-id="2c8e6-173">obrigatório</span><span class="sxs-lookup"><span data-stu-id="2c8e6-173">required</span></span>|<span data-ttu-id="2c8e6-174">Valor do token avô quando a condição for atendida.</span><span class="sxs-lookup"><span data-stu-id="2c8e6-174">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="2c8e6-175">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2c8e6-175">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="2c8e6-176">Confira também</span><span class="sxs-lookup"><span data-stu-id="2c8e6-176">See also</span></span>

- [<span data-ttu-id="2c8e6-177">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="2c8e6-177">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="2c8e6-178">Definir o elemento Requirements no manifesto</span><span class="sxs-lookup"><span data-stu-id="2c8e6-178">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="2c8e6-179">Atalhos de teclado para o SharePoint</span><span class="sxs-lookup"><span data-stu-id="2c8e6-179">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
