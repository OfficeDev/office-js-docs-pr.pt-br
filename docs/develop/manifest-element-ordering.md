---
title: Como encontrar a ordem correta dos elementos do manifesto
description: Saiba como encontrar a ordem correta na qual colocar elementos filho em um elemento pai.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ee80167a76861209e814dc6c272720feb3a9cf1
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173910"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Como encontrar a ordem correta dos elementos do manifesto

Os elementos XML do manifesto de um Suplemento do Office devem estar no elemento pai apropriado *e* em uma ordem específica em relação uns aos outros.

A ordem exigida é especificada nos arquivos XSD, na pasta [Esquemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Os arquivos XSD são categorizados em subpastas para suplementos de painel de tarefas, conteúdo e email.

Por exemplo, no elemento `<OfficeApp>`, os elementos `<Id>`, `<Version>` e `<ProviderName>` devem aparecer nessa ordem. Se adicionar um elemento `<AlternateId>`, deverá colocá-lo entre os elementos `<Id>` e `<Version>`. Se algum dos elementos estiver na posição incorreta, o manifesto não será válido e o suplemento não será carregado.

> [!NOTE]
> O [validador no office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) usa a mesma mensagem de erro quando um elemento está fora de ordem que quando um elemento está sob o pai errado. A mensagem de erro informa que o elemento não é um elemento filho válido do elemento pai. Caso receba este erro, mas a documentação de referência do elemento filho indique que ele *está* válido para o pai, talvez o problema seja o filho ter sido colocado na ordem incorreta.

As seções a seguir mostram os elementos de manifesto na ordem em que devem aparecer. Há diferenças dependendo se o `type` atributo do elemento é , ou `<OfficeApp>` `TaskPaneApp` `ContentApp` `MailApp` . Para evitar que essas seções se tornam muito complicadas, o elemento altamente complexo é `<VersionOverrides>` dividido em seções separadas.

> [!Note]
> Nem todos os elementos mostrados são obrigatórios. Se o `minOccurs` valor de um elemento for **0** no [esquema,](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3)o elemento será opcional.

## <a name="basic-task-pane-add-in-element-ordering"></a>Ordenação básica de elementos do painel de tarefas

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

\*Confira a ordenação de elementos do complemento do painel de tarefas [em VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) para a ordenação de elementos filhos de VersionOverrides.

## <a name="basic-mail-add-in-element-ordering"></a>Ordenação de elementos básicos do complemento de email

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

\*Confira a ordenação de elementos de add-in de email em [VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) e ordenação de elementos do complemento Mail em [VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para a ordenação de elementos filhos de VersionOverrides.

## <a name="basic-content-add-in-element-ordering"></a>Ordenação básica de elementos do complemento de conteúdo

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

\*Confira a ordenação de elementos do complemento de conteúdo [em VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) para a ordenação de elementos filhos de VersionOverrides.

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>Ordenação de elementos do complemento do painel de tarefas em VersionOverrides

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
                        <OverriddenByRibbonApi>
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>Ordenação de elementos de add-in de email em VersionOverrides Ver. 1.0

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

\* Um VersionOverrides com valor , em vez de , pode ser `type` `VersionOverridesV1_1` `VersionOverridesV1_0` aninhado no final dos VersionOverrides externos. Confira a ordenação de elementos de complemento de email [em VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) para a ordenação de elementos em `VersionOverridesV1_1` .

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>Ordenação de elementos de add-in de email em VersionOverrides Ver. 1.1

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

## <a name="content-add-in-element-ordering-within-versionoverrides"></a>Ordenação de elementos do complemento de conteúdo em VersionOverrides

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a>Confira também

- [Referência para manifestos de Complementos do Office (v1.1)](../develop/add-in-manifests.md)
- [Definições oficiais de esquema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
