---
title: Manifesto XML dos Suplementos do Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: 24c212335fa50feb4d13b6069a24cacbd9849715
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-ins-xml-manifest"></a>Manifesto XML dos Suplementos do Office

O arquivo de manifesto XML de um Suplemento do Office descreve como seu suplemento deve ser ativado quando um usu?rio final o instala e usa com os aplicativos e documentos do Office.

Um arquivo de manifesto XML com base nesse esquema permite que um Suplemento do Office fa?a o seguinte:

* Descreva a si mesmo fornecendo ID, vers?o, descri??o, nome para exibi??o e local padr?o.

* Especifique as imagens usadas para identidade visual do suplemento e a iconografia usada para os [Comandos do suplemento][] na Faixa de Op??es do Office.

* Especifique como o suplemento se integra ao Office, incluindo qualquer interface do usu?rio personalizada, como bot?es da faixa de op??es criados pelo suplemento.

* Especifique as dimens?es padr?o solicitadas para suplementos de conte?do e a altura solicitada para Suplementos do Outlook.

* Declare permiss?es exigidas pelo Suplemento do Office, como ler ou gravar no documento.

* Para os suplementos do Outlook, defina a regra ou as regras que especificam o contexto no qual ser?o ativados e interagir?o com uma mensagem, compromisso ou item de solicita??o da reuni?o.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).

## <a name="required-elements"></a>Elementos exigidos

A tabela a seguir especifica os elementos exigidos para os tr?s tipos de Suplementos do Office.

### <a name="required-elements-by-office-add-in-type"></a>Elementos obrigat?rios de acordo com o tipo de Suplemento do Office

| Elemento                                                                                      | Conte?do | Painel de tarefas | Outlook |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| [OfficeApp][]                                                                                |    X    |     X     |    X    |
| [Id][]                                                                                       |    X    |     X     |    X    |
| [Vers?o][]                                                                                  |    X    |     X     |    X    |
| [ProviderName][]                                                                             |    X    |     X     |    X    |
| [DefaultLocale][]                                                                            |    X    |     X     |    X    |
| [DisplayName][]                                                                              |    X    |     X     |    X    |
| [Descri??o][]                                                                              |    X    |     X     |    X    |
| [IconUrl][]                                                                                  |    X    |     X     |    X    |
| [HighResolutionIconUrl][]                                                                    |    X    |     X     |    X    |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       |    X    |     X     |         |
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]                         |    X    |     X     |         |
| [DesktopSettings][]                                                                          |         |           |    X    |
| [SourceLocation (MailApp)][]                                                                 |         |           |    X    |
| [Permiss?es (ContentApp)][]<br/>[Permiss?es (TaskPaneApp)][]<br/>[Permiss?es (MailApp)][] |    X    |     X     |    X    |
| [Regra (RuleCollection)][]<br/>[Regra (MailApp)][]                                             |         |           |    X    |
| [Requisitos (MailApp)*][]                                                                  |         |           |    X    |
| [Conjunto*][]<br/>[Conjuntos (MailAppRequirements)*][]                                                 |         |           |    X    |
| [Formul?rio*][]<br/>[Formsettings*][]                                                              |         |           |    X    |
| [Conjuntos (Requisitos)*][]                                                                     |    X    |     X     |         |
| [Hosts*][]                                                                                   |    X    |     X     |         |

_\*Adicionados no esquema de manifesto de suplementos da vers?o 1.1 do Office._

<!-- Links for above table -->

[officeapp]: http://msdn.microsoft.com/en-us/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx
[id]: http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx
[vers?o]: http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx
[providername]: http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx
[defaultlocale]: http://msdn.microsoft.com/en-us/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx
[displayname]: http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx
[descri??o]: http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx
[iconurl]: http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx
[highresolutioniconurl]: http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx
[defaultsettings (contentapp)]: http://msdn.microsoft.com/en-us/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx
[defaultsettings (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx
[sourcelocation (contentapp)]: http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx
[sourcelocation (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx
[desktopsettings]: http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/en-us/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permiss?es (contentapp)]: http://msdn.microsoft.com/en-us/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx
[permiss?es (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx
[permiss?es (mailapp)]: http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx
[regra (rulecollection)]: http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx
[regra (mailapp)]: http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx
[requisitos (mailapp)]: http://msdn.microsoft.com/en-us/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx
[conjunto*]: http://msdn.microsoft.com/en-us/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx
[conjuntos (mailapprequirements)*]: http://msdn.microsoft.com/en-us/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx
[formul?rio*]: http://msdn.microsoft.com/en-us/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx
[formsettings*]: http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx
[conjuntos (requisitos)*]: http://msdn.microsoft.com/en-us/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx
[hosts*]: http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx

## <a name="hosting-requirements"></a>Requisitos de hospedagem

Todas as imagem URIs, como as usadas para os [Comandos do suplemento][], devem ser compat?veis com armazenamento em cache. O servidor que hospeda a imagem n?o deve retornar um cabe?alho `Cache-Control` especificando `no-cache`, `no-store` ou op??es semelhantes na resposta HTTP.

Todas as URLs, como os locais dos arquivos de origem especificados no elemento [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation), devem estar **protegidos por SSL (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>Pr?ticas recomendadas de envio ao AppSource

Verifique se a identifica??o do suplemento ? um GUID v?lido e exclusivo. Diversas ferramentas de gerador de GUID est?o dispon?veis na Web e podem ser usadas para criar um GUID exclusivo.

Os suplementos enviados ao AppSource tamb?m devem conter o elemento [SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl). Saiba mais em [Pol?ticas de valida??o para aplicativos e suplementos enviados ao AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).

Use apenas o elemento [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) para especificar dom?nios diferentes daqueles especificados no elemento [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) para cen?rios de autentica??o.

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>Especificar os dom?nios que voc? deseja abrir na janela do suplemento

Por padr?o, se o suplemento tentar acessar uma URL em um dom?nio diferente do dom?nio que hospeda a p?gina inicial (conforme especificado no elemento [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) do arquivo de manifesto), essa URL abrir? em uma nova janela de navegador fora do painel de suplementos do aplicativo host do Office. Esse comportamento padr?o protege o usu?rio contra a navega??o de p?gina inesperada dentro do painel de suplemento de elementos **iFrame**.

Para substituir esse comportamento, especifique cada dom?nio que voc? deseja abrir na janela do suplemento na lista de dom?nios especificados no elemento [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) do arquivo de manifesto. Se o suplemento tentar acessar uma URL em um dom?nio que n?o est? na lista, essa URL abre em uma nova janela do navegador (fora do painel de suplementos).

O exemplo de manifesto XML a seguir hospeda sua p?gina de suplemento principal no dom?nio `https://www.contoso.com`, conforme especificado no elemento **SourceLocation**. Ele tamb?m especifica o dom?nio `https://www.northwindtraders.com` em um elemento [AppDomain](http://msdn.microsoft.com/en-us/library/2a0353ec-5e09-6fbf-1636-4bb5dcebb9bf%28Office.15%29.aspx), dentro da lista de elementos **AppDomains**. Se o suplemento acessar uma p?gina no dom?nio www.northwindtraders.com, essa p?gina abrir? no painel do suplemento.

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>Exemplos e esquemas do arquivo XML de manifesto v1.1
As se??es a seguir mostram exemplos de arquivos XML de manifesto v1.1 para suplementos de conte?do, de painel de tarefas e do Outlook.

# <a name="task-panetabtabid-1"></a>[Painel de tarefas](#tab/tabid-1)

[Esquema de manifesto do aplicativo do painel de tarefas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
   <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
            <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
                <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                 <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                     <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                     <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
                  <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                     <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
            <!-- Menu example -->
            <Control xsi:type="Menu" id="Contoso.Menu">
              <Label resid="Contoso.Dropdown.Label" />
              <Supertip>
                <Title resid="Contoso.Dropdown.Label" />
                <Description resid="Contoso.Dropdown.Tooltip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
              </Icon>
              <Items>
                <Item id="Contoso.Menu.Item1">
                  <Label resid="Contoso.Item1.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item1.Label" />
                    <Description resid="Contoso.Item1.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Item>

                <Item id="Contoso.Menu.Item2">
                  <Label resid="Contoso.Item2.Label"/>
                  <Supertip>
                    <Title resid="Contoso.Item2.Label" />
                    <Description resid="Contoso.Item2.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane2.Url" />
                  </Action>
                </Item>

              </Items>
            </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="contenttabtabid-2"></a>[Conte?do](#tab/tabid-2)

[Esquema de manifesto do aplicativo de conte?do](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mailtabtabid-3"></a>[Email](#tab/tabid-3)

[Esquema de manifesto do aplicativo de email](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Validar e solucionar problemas com seu manifesto

Para solucionar problemas com seu manifesto, confira [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). L?, voc? encontrar? informa??es sobre como validar o manifesto em rela??o ? [Defini??o de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) e tamb?m como usar o log de tempo de execu??o para depurar o manifesto.

## <a name="see-also"></a>Veja tamb?m

* [Criar comandos de suplementos em seu manifesto][comandos de suplementos]
* [Especificar requisitos da API e de hosts do Office](specify-office-hosts-and-api-requirements.md)
* [Localiza??o para suplementos do Office](localization.md)
* [Refer?ncia de esquema para manifestos de suplementos do Office](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md)

[Comandos de suplemento]: create-addin-commands.md