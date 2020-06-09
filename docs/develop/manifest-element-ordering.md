---
title: Como encontrar a ordem correta dos elementos do manifesto
description: Saiba como encontrar a ordem correta na qual colocar elementos filho em um elemento pai.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 1c9f6323e4bc904fd254f776e1f9c132b0050e91
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611986"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="5cc02-103">Como encontrar a ordem correta dos elementos do manifesto</span><span class="sxs-lookup"><span data-stu-id="5cc02-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="5cc02-104">Os elementos XML do manifesto de um Suplemento do Office devem estar no elemento pai apropriado *e* em uma ordem específica em relação uns aos outros.</span><span class="sxs-lookup"><span data-stu-id="5cc02-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="5cc02-105">A ordem exigida é especificada nos arquivos XSD, na pasta [Esquemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="5cc02-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="5cc02-106">Os arquivos XSD são categorizados em subpastas para suplementos de painel de tarefas, conteúdo e email.</span><span class="sxs-lookup"><span data-stu-id="5cc02-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="5cc02-107">Por exemplo, no elemento `<OfficeApp>`, os elementos `<Id>`, `<Version>` e `<ProviderName>` devem aparecer nessa ordem.</span><span class="sxs-lookup"><span data-stu-id="5cc02-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="5cc02-108">Se adicionar um elemento `<AlternateId>`, deverá colocá-lo entre os elementos `<Id>` e `<Version>`.</span><span class="sxs-lookup"><span data-stu-id="5cc02-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="5cc02-109">Se algum dos elementos estiver na posição incorreta, o manifesto não será válido e o suplemento não será carregado.</span><span class="sxs-lookup"><span data-stu-id="5cc02-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="5cc02-110">O [validador no Office-AddIn-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) usa a mesma mensagem de erro quando um elemento está fora de ordem, como ocorre quando um elemento está sob o pai errado.</span><span class="sxs-lookup"><span data-stu-id="5cc02-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="5cc02-111">A mensagem de erro informa que o elemento não é um elemento filho válido do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="5cc02-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="5cc02-112">Caso receba este erro, mas a documentação de referência do elemento filho indique que ele *está* válido para o pai, talvez o problema seja o filho ter sido colocado na ordem incorreta.</span><span class="sxs-lookup"><span data-stu-id="5cc02-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="5cc02-113">As seções a seguir mostram os elementos manifest na ordem em que devem ser exibidos.</span><span class="sxs-lookup"><span data-stu-id="5cc02-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="5cc02-114">Há diferenças dependendo se o `type` atributo do `<OfficeApp>` elemento é `TaskPaneApp` , `ContentApp` ou `MailApp` .</span><span class="sxs-lookup"><span data-stu-id="5cc02-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="5cc02-115">Para evitar que essas seções fiquem muito difíceis, o elemento altamente complexo `<VersionOverrides>` é dividido em seções separadas.</span><span class="sxs-lookup"><span data-stu-id="5cc02-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="5cc02-116">Nem todos os elementos mostrados são obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="5cc02-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="5cc02-117">Se o `minOccurs` valor de um elemento for **0** no [esquema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), o elemento será opcional.</span><span class="sxs-lookup"><span data-stu-id="5cc02-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="5cc02-118">Ordenação básica de elemento de suplemento do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5cc02-118">Basic task pane add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="TaskPaneApp">
    <Id>
    <AlternateID>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
        <Sets>
            <Set>
        <Methods>
            <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <Permissions>
    <Dictionary>
        <TargetDialects>
        <QueryUri>
        <CitationText>
        <DictionaryName>
        <DictionaryHomePage>
    <VersionOverrides>*
```

<span data-ttu-id="5cc02-119">\*Confira [ordenação de elemento do suplemento do painel de tarefas no VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) para a ordenação dos elementos filhos de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="5cc02-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="5cc02-120">Ordenação básica de elementos de suplemento de email</span><span class="sxs-lookup"><span data-stu-id="5cc02-120">Basic mail add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="MailApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <FormSettings>
        <Form>
        <DesktopSettings>
            <SourceLocation>
            <RequestedHeight>
        <TabletSettings>
            <SourceLocation>
            <RequestedHeight>
        <PhoneSettings>
            <SourceLocation>
    <Permissions>
    <Rule>
    <DisableEntityHighlighting>
    <VersionOverrides>*
```

<span data-ttu-id="5cc02-121">\*Veja [ordenação de elemento de suplemento de email dentro do VersionOverrides ver. 1,0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) e [email ordenação do elemento de suplemento em VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para a ordenação dos elementos filhos de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="5cc02-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="5cc02-122">Ordenação de elemento de suplemento de conteúdo básico</span><span class="sxs-lookup"><span data-stu-id="5cc02-122">Basic content add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="ContentApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl >
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <Methods>
        <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <RequestedWidth>
    <RequestedHeight>
    <Permissions>
    <AllowSnapshot>
    <VersionOverrides>*
```

<span data-ttu-id="5cc02-123">\*Veja [ordenação de elemento do suplemento de conteúdo dentro do VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) para a ordenação dos elementos filhos de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="5cc02-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="5cc02-124">Ordenação do elemento do suplemento do painel de tarefas no VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5cc02-124">Task pane add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <Runtimes>
                <Runtime>
            <AllFormFactors>
                <ExtensionPoint>
                    <Script>
                        <SourceLocation>
                    <Page>
                        <SourceLocation>
                    <Metadata>
                        <SourceLocation>
                    <Namespace>
            <DesktopFormFactor>
                <GetStarted>
                    <Title>
                    <Description>
                    <LearnMoreUrl>
                <FunctionFile>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
        <Resources>
            <Images>
                <Image>
                    <Override>
            <Urls>
                <Url>
                    <Override>
            <ShortStrings>
                <String>
                    <Override>
            <LongStrings>
                <String>
                    <Override>
        <WebApplicationInfo>
            <Id>
            <MsaId>
            <Resource>
            <Scopes>
                <Scope>
            <Authorizations>
                <Authorization>
                    <Resource>
                    <Scopes>
                        <Scope>
        <EquivalentAddins>
            <EquivalentAddin>
                <ProgId>
                <DisplayName>
                <FileName>
                <Type>
```

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="5cc02-125">Ordenação do elemento de suplemento de email no VersionOverrides ver.</span><span class="sxs-lookup"><span data-stu-id="5cc02-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="5cc02-126">1.0</span><span class="sxs-lookup"><span data-stu-id="5cc02-126">1.0</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <DesktopFormFactor>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>
                                <Action>
                                    <SourceLocation>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Items>
                                    <Item>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <VersionOverrides>*
```

<span data-ttu-id="5cc02-127">\*Um VersionOverrides com `type` valor `VersionOverridesV1_1` , em vez de `VersionOverridesV1_0` , pode ser aninhado no final da VersionOverrides externa.</span><span class="sxs-lookup"><span data-stu-id="5cc02-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="5cc02-128">Veja [ordenação de elemento de suplemento de email em VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para a ordenação dos elementos no `VersionOverridesV1_1` .</span><span class="sxs-lookup"><span data-stu-id="5cc02-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="5cc02-129">Ordenação do elemento de suplemento de email no VersionOverrides ver.</span><span class="sxs-lookup"><span data-stu-id="5cc02-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="5cc02-130">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc02-130">1.1</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
    <Sets>
        <Set>
    <Hosts>
    <Host>
        <DesktopFormFactor>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <SourceLocation>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                    <SourceLocation>
                <SourceLocation>
                <Label>
                <CommandSurface>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="5cc02-131">Ordenação de elemento de suplemento de conteúdo dentro de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="5cc02-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="5cc02-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="5cc02-132">See also</span></span>

- [<span data-ttu-id="5cc02-133">Referência de esquema para manifestos de Suplementos do Office (versão 1.1)</span><span class="sxs-lookup"><span data-stu-id="5cc02-133">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
