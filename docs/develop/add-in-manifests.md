---
title: Manifesto XML dos Suplementos do Office
description: Obtenha uma visão geral do manifesto de suplemento do Office e seus usos.
ms.date: 05/24/2022
ms.localizationpriority: high
ms.openlocfilehash: 60368d74cad0d1b8c0562888613d960f52b21a74
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810222"
---
# <a name="office-add-ins-xml-manifest"></a>Manifesto XML dos Suplementos do Office

O arquivo de manifesto XML de um Suplemento do Office descreve como seu suplemento deve ser ativado quando um usuário final o instala e usa com os aplicativos e documentos do Office.

> [!TIP]
> Este artigo descreve o manifesto atual formatado em XML. Também há um manifesto de Teams formatado em JSON que está disponível na visualização. Para obter mais informações, consulte [Manifesto do Teams para Suplementos do Office (visualização)](json-manifest-overview.md).

Um arquivo de manifesto XML permite que um Suplemento do Office faça o seguinte:

- Descreva a si mesmo fornecendo ID, versão, descrição, nome para exibição e local padrão.

- Especifique as imagens usadas para identidade visual do suplemento e a iconografia usada para os [comandos do suplemento](create-addin-commands.md) na faixa de opções do Aplicativo do Office.

- Especifique como o suplemento se integra ao Office, incluindo qualquer interface do usuário personalizada, como botões da faixa de opções criados pelo suplemento.

- Especifique as dimensões padrão solicitadas para suplementos de conteúdo e a altura solicitada para Suplementos do Outlook.

- Declare permissões exigidas pelo Suplemento do Office, como ler ou gravar no documento.

- Para os suplementos do Outlook, defina a regra ou as regras que especificam o contexto no qual serão ativados e interagirão com uma mensagem, compromisso ou item de solicitação da reunião.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a>Elementos exigidos

A tabela a seguir especifica os elementos exigidos para os três tipos de Suplementos do Office.

> [!NOTE]
> Também há uma ordem obrigatória na qual os elementos devem aparecer dentro de seu elemento-pai. Confira mais informações em [Como encontrar a ordem adequada de elementos de manifesto](manifest-element-ordering.md).

### <a name="required-elements-by-office-add-in-type"></a>Elementos obrigatórios de acordo com o tipo de Suplemento do Office

| Elemento                                                                                      | Conteúdo    | Painel de tarefas    | Email<br>(Outlook)      |
| :------------------------------------------------------------------------------------------- | :--------: | :----------: | :--------:   |
| [OfficeApp][]                                                                                | Obrigatório   | Obrigatório     | Obrigatório     |
| [Id][]                                                                                       | Obrigatório   | Obrigatório     | Obrigatório     |
| [Versão][]                                                                                  | Obrigatório   | Obrigatório     | Obrigatório     |
| [ProviderName][]                                                                             | Obrigatório   | Obrigatório     | Obrigatório     |
| [DefaultLocale][]                                                                            | Obrigatório   | Obrigatório     | Obrigatório     |
| [DisplayName][]                                                                              | Obrigatório   | Obrigatório     | Obrigatório     |
| [Description][]                                                                              | Obrigatório   | Obrigatório     | Obrigatório     |
| [IconUrl][]                                                                                  | Obrigatório   | Obrigatório     | Obrigatório     |
| [SupportUrl][]\*\*                                                                           | Obrigatório   | Obrigatório     | Obrigatório     |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       | Obrigatório   | Obrigatório     | Não disponível|
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]<br/>[SourceLocation (MailApp)][]| Obrigatório | Obrigatório | Obrigatório   |
| [DesktopSettings][]                                                                          | Não disponível | Não disponível | Obrigatório |
| [Permissões (ContentApp)][]<br/>[Permissões (TaskPaneApp)][]<br/>[Permissões (MailApp)][] | Obrigatório   | Obrigatório     | Obrigatório     |
| [Regra (RuleCollection)][]<br/>[Regra (MailApp)][]                                             | Não disponível | Não disponível | Obrigatório |
| [Requisitos (MailApp)][]\*                                                                 | Não aplicável| Não disponível | Obrigatório |
| [Conjuntos][]\*<br/>[Conjuntos (Requisitos)][]\*<br/>[Conjuntos (MailAppRequirements)][]\*                 | Obrigatório   | Obrigatório     | Obrigatório     |
| [Formulário][]\*<br/>[FormSettings][]\*                                                            | Não disponível | Não disponível | Obrigatório |
| [Hosts][]\*                                                                                  | Obrigatório   | Obrigatório     | Opcional     |

_\*Adicionados no esquema de manifesto de suplementos da versão 1.1 do Office._

_\*\* SupportUrl só é necessário para suplementos distribuídos pelo AppSource._

<!-- Links for above table -->

[officeapp]: /javascript/api/manifest/officeapp
[id]: /javascript/api/manifest/id
[version]: /javascript/api/manifest/version
[providername]: /javascript/api/manifest/providername
[defaultlocale]: /javascript/api/manifest/defaultlocale
[displayname]: /javascript/api/manifest/displayname
[description]: /javascript/api/manifest/description
[iconurl]: /javascript/api/manifest/iconurl
[supporturl]: /javascript/api/manifest/supporturl
[defaultsettings (contentapp)]: /javascript/api/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /javascript/api/manifest/defaultsettings
[sourcelocation (contentapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (mailapp)]: /javascript/api/manifest/sourcelocation
[desktopsettings]: /javascript/api/manifest/desktopsettings
[permissões (contentapp)]: /javascript/api/manifest/permissions
[permissões (taskpaneapp)]: /javascript/api/manifest/permissions
[permissões (mailapp)]: /javascript/api/manifest/permissions
[regra (rulecollection)]: /javascript/api/manifest/rule
[regra (mailapp)]: /javascript/api/manifest/rule
[requisitos (mailapp)]: /javascript/api/manifest/requirements
[set]: /javascript/api/manifest/set
[conjuntos (mailapprequirements)]: /javascript/api/manifest/sets
[formulário]: /javascript/api/manifest/form
[formsettings]: /javascript/api/manifest/formsettings
[conjuntos (requisitos)]: /javascript/api/manifest/sets
[hosts]: /javascript/api/manifest/hosts

## <a name="hosting-requirements"></a>Requisitos de hospedagem

Todas as imagem URIs, como as usadas para os [comandos do suplemento](create-addin-commands.md), devem ser compatíveis com armazenamento em cache. O servidor que hospeda a imagem não deve retornar um cabeçalho `Cache-Control` especificando `no-cache`, `no-store` ou opções semelhantes na resposta HTTP.

Todas as URLs, como os locais dos arquivos de origem especificados no elemento [SourceLocation](/javascript/api/manifest/sourcelocation), devem estar **protegidos por SSL (HTTPS)**. [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>Práticas recomendadas de envio ao AppSource

Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.

Os suplementos enviados ao AppSource também devem conter o elemento [SupportUrl](/javascript/api/manifest/supporturl). Saiba mais em [Políticas de validação para aplicativos e suplementos enviados ao AppSource](/legal/marketplace/certification-policies).

Use apenas o elemento [AppDomains](/javascript/api/manifest/appdomains) para especificar domínios diferentes daqueles especificados no elemento [SourceLocation](/javascript/api/manifest/sourcelocation) para cenários de autenticação.

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>Especificar os domínios que você deseja abrir na janela do suplemento

When running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/javascript/api/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.

Para substituir esse comportamento (Office para desktop), especifique cada domínio que você deseja abrir na janela do suplemento na lista de domínios especificados no elemento [AppDomains](/javascript/api/manifest/appdomains) do arquivo de manifesto. Se o suplemento tentar ir para uma URL em um domínio que está na lista, ela então abre no painel de tarefas do Office para desktop e no Office Online. Se ele tentar acessar uma URL que não está na lista, no Office para desktop, essa URL abre em uma nova janela do navegador (fora do painel de suplementos).

> [!NOTE]
> Há duas exceções para esse comportamento.
>
> - Isso se aplica somente ao painel raiz do suplemento. Se houver um iframe incorporado à página do suplemento, o iframe poderá ser direcionado para qualquer URL, independentemente de estar listado em **\<AppDomains\>**, mesmo no Office para desktop.
> - Quando uma caixa de diálogo é aberta com a API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#office-office-ui-displaydialogasync-member(1)), a URL que é passada para o método deve estar no mesmo domínio que o suplemento, mas a caixa de diálogo pode ser direcionada para qualquer URL, independentemente de estar listada em **\<AppDomains\>**, mesmo no Office para desktop.

O exemplo de manifesto XML a seguir hospeda sua página principal do suplemento no domínio `https://www.contoso.com`, conforme especificado no elemento **\<SourceLocation\>**. Ele também especifica o domínio `https://www.northwindtraders.com` em um elemento [AppDomain](/javascript/api/manifest/appdomain) dentro da lista de elementos **\<AppDomains\>**. Se o suplemento acessar uma página no `www.northwindtraders.com`domínio, essa página abre no painel do suplemento, mesmo na área de trabalho do Office.

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="version-overrides-in-the-manifest"></a>Substituições de versão no manifesto

O elemento opcional [VersionOverrides](/javascript/api/manifest/versionoverrides) merece uma menção especial. Ele contém marcação infantil que habilita a recursos de suplemento adicionais. Alguns deles são:

- Personalizando a faixa de opções do Office e os menus.
- Personalizando como o Office funciona com os runtimes inseridos nos quais os suplementos são executados.
- Configurando como o suplemento interage com o Azure Active Directory e o Microsoft Graph para o Logon único.

Alguns elementos descendentes de `VersionOverrides` têm valores que substituem os valores do elemento pai `OfficeApp`. Por exemplo, o elemento `Hosts` em `VersionOverrides` substitui o elemento `Hosts` em `OfficeApp`.

The `VersionOverrides` element has its own schema, actually four of them, depending on the type of add-in and the features it uses. The schemas are:

- [Painel de tarefas 1.0](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)
- [Conteúdo 1.0](/openspecs/office_file_formats/ms-owemxml/c9cb8dca-e9e7-45a7-86b7-f1f0833ce2c7)
- [Email 1.0](/openspecs/office_file_formats/ms-owemxml/578d8214-2657-4e6a-8485-25899e772fac)
- [Email 1.1](/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)

Quando um elemento `VersionOverrides` é utilizado, então o elemento `OfficeApp` deve ter um atributo `xmlns` que identifica o esquema apropriado. Os possíveis valores do atributo são os seguintes:

- `http://schemas.microsoft.com/office/taskpaneappversionoverrides`
- `http://schemas.microsoft.com/office/contentappversionoverrides`
- `http://schemas.microsoft.com/office/mailappversionoverrides`

O elemento `VersionOverrides` em si também deve ter um atributo `xmlns` especificando o esquema. Os possíveis valores são os três acima e os seguintes:

- `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`

O elemento `VersionOverrides` também deve ter um atributo `xsi:type` que especifica a versão do esquema. Os valores possíveis são os seguintes:

- `VersionOverridesV1_0`
- `VersionOverridesV1_1`

A seguir estão alguns exemplos de `VersionOverrides`utilizados, respectivamente, em um suplemento do painel de tarefas e em um suplemento de email. Observe que quando um email `VersionOverrides` com a versão 1.1 é usado, ele deve ser o último filho de um pai `VersionOverrides` do tipo 1.0. Os valores dos elementos filho no interior de `VersionOverrides` substituem os valores dos elementos do mesmo nome no elemento pai `VersionOverrides` e do elemento avô `OfficeApp`.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- child elements omitted -->
</VersionOverrides>
```

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <!-- other child elements omitted -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- child elements omitted -->
  </VersionOverrides>
</VersionOverrides>
```

Para obter um exemplo de um manifesto que inclui um elemento `VersionOverrides`, consulte [Exemplos e esquemas do arquivo XML de manifesto v1.1](#manifest-v11-xml-file-examples-and-schemas).

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a>Especificar domínios a partir dos quais as chamadas da API do Office.js são feitas

Seu suplemento pode fazer chamadas API do Office.js a partir do domínio referenciado no elemento[SourceLocation](/javascript/api/manifest/sourcelocation) do arquivo de manifesto. Se você tiver outros iFrames dentro de seu suplemento que precisem acessar APIs do Office.js, adicione o domínio dessa URL de origem à lista especificada no elemento [AppDomains](/javascript/api/manifest/appdomains) do arquivo de manifesto. Se um iFrame com uma fonte não incluída na lista `AppDomains` tentar fazer uma chamada de API do Office. js, o suplemento receberá um[ erro de permissão negada](../reference/javascript-api-for-office-error-codes.md).

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>Exemplos e esquemas do arquivo XML de manifesto v1.1

As seções a seguir mostram exemplos de arquivos XML v1.1 de manifesto para conteúdo, painel de tarefas e suplementos de email (Outlook).

# <a name="task-pane"></a>[Painel de tarefas](#tab/tabid-1)

[Esquemas de manifesto de suplemento](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
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
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
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

# <a name="content"></a>[Conteúdo](#tab/tabid-2)

[Esquemas de manifesto de suplemento](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

# <a name="mail"></a>[Email](#tab/tabid-3)

[Esquemas de manifesto de suplemento](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

## <a name="validate-an-office-add-ins-manifest"></a>Validar o manifesto de suplemento do Office

Para saber mais sobre como validar um manifesto em relação à [Definição do Esquema XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), confira [Validar o manifesto de suplemento do Office](../testing/troubleshoot-manifest.md).

## <a name="see-also"></a>Confira também

- [Como identificar a ordem correta dos elementos do manifesto](manifest-element-ordering.md)
- [Criar comandos de suplementos em seu manifesto](create-addin-commands.md)
- [Especificar requisitos da API e de aplicativos do Office](specify-office-hosts-and-api-requirements.md)
- [Localização para suplementos do Office](localization.md)
- [Referência de esquema para manifestos de suplementos do Office](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
- [Atualizar a versão da API e do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md)
- [Identificar um suplemento COM equivalente](make-office-add-in-compatible-with-existing-com-add-in.md)
- [Solicitar permissões para uso da API em suplementos ](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [Validar o manifesto de suplemento do Office](../testing/troubleshoot-manifest.md)
