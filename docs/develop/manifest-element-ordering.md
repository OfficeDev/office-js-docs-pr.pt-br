---
title: Como encontrar a ordem correta dos elementos do manifesto
description: Saiba como encontrar a ordem correta na qual colocar elementos filho em um elemento pai.
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8c460c970c0288389097f64e5de09f74744da892
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660106"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Como encontrar a ordem correta dos elementos do manifesto

Os elementos XML do manifesto de um Suplemento do Office devem estar no elemento pai apropriado *e* em uma ordem específica em relação uns aos outros.

A ordem exigida é especificada nos arquivos XSD, na pasta [Esquemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Os arquivos XSD são categorizados em subpastas para suplementos de painel de tarefas, conteúdo e email.

Por exemplo, no elemento **\<OfficeApp\>** , o **\<Id\>**, **\<Version\>** deve **\<ProviderName\>** aparecer nessa ordem. Se um **\<AlternateId\>** elemento for adicionado, ele deverá estar entre o elemento **\<Id\>** e **\<Version\>** o elemento. Se algum dos elementos estiver na posição incorreta, o manifesto não será válido e o suplemento não será carregado.

> [!NOTE]
> O [validador no office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) usa a mesma mensagem de erro quando um elemento está fora de ordem como faz quando um elemento está sob o pai errado. A mensagem de erro informa que o elemento não é um elemento filho válido do elemento pai. Caso receba este erro, mas a documentação de referência do elemento filho indique que ele *está* válido para o pai, talvez o problema seja o filho ter sido colocado na ordem incorreta.

As seções a seguir mostram os elementos de manifesto na ordem em que eles devem aparecer. Há diferenças dependendo se o `type` atributo do elemento **\<OfficeApp\>** é `TaskPaneApp`, `ContentApp`ou `MailApp`. Para impedir que essas seções se torne muito complicadas, **\<VersionOverrides\>** o elemento altamente complexo é dividido em seções separadas.

> [!Note]
> Nem todos os elementos mostrados são obrigatórios. Se o `minOccurs` valor de um elemento for **0** no [esquema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), o elemento será opcional.

## <a name="basic-task-pane-add-in-element-ordering"></a>Ordenação básica de elementos do suplemento do painel de tarefas

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
    <ExtendedOverrides>
```

\*Consulte a ordenação de elementos do suplemento do painel de tarefas [em VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) para obter a ordenação de elementos filhos de VersionOverrides.

## <a name="basic-mail-add-in-element-ordering"></a>Ordenação básica de elementos do suplemento de email

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

\*Consulte a ordenação de elementos do suplemento Mail no [VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) e na ordenação de elementos do suplemento [Mail no VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para a ordenação de elementos filhos de VersionOverrides.

## <a name="basic-content-add-in-element-ordering"></a>Ordenação de elementos do suplemento de conteúdo básico

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

\*Consulte [a ordenação de elementos de suplemento de conteúdo em VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) para obter a ordenação de elementos filhos de VersionOverrides.

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>Ordenação de elementos do suplemento do painel de tarefas em VersionOverrides

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
                        <Group> (can be below <ControlGroup>)
                            <OverriddenByRibbonApi>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <OverriddenByRibbonApi>
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
                                        <OverriddenByRibbonApi>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <ControlGroup> (can be above <Group>)
                        <Label>
                        <InsertAfter> (or <InsertBefore>)
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>Ordenação de elemento de suplemento de email no VersionOverrides Ver. 1.0

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

\* Um VersionOverrides com `type` valor `VersionOverridesV1_1`, em vez de `VersionOverridesV1_0`, pode ser aninhado no final dos VersionOverrides externos. Consulte [a ordenação de elementos do suplemento Mail no VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para obter a ordenação de elementos em `VersionOverridesV1_1`.

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>Ordenação de elemento de suplemento de email no VersionOverrides Ver. 1.1

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

## <a name="content-add-in-element-ordering-within-versionoverrides"></a>Ordenação de elementos do suplemento de conteúdo em VersionOverrides

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a>Confira também

- [Referência para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md)
- [Definições oficiais de esquema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
