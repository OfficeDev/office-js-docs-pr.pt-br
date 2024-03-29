---
title: Manifestos do suplemento do Outlook
description: Obtenha uma visão geral dos dois tipos de manifestos disponíveis para suplementos do Outlook.
ms.date: 10/18/2022
ms.localizationpriority: high
ms.openlocfilehash: a22b5180fee6b4f9f0663eff54b57510016202a2
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607552"
---
# <a name="outlook-add-in-manifests"></a>Manifestos do suplemento do Outlook

Um suplemento do Outlook consiste em dois componentes: o manifesto do suplemento e um aplicativo Web compatível com a biblioteca JavaScript para suplementos do Office (office.js). O manifesto descreve como o suplemento integra-se a clientes do Outlook.

Há dois formatos possíveis para o manifesto: XML e JSON. Você pode saber mais sobre o manifesto JSON no manifesto [do Teams para Suplementos do Office (versão prévia)](../develop/json-manifest-overview.md). Este artigo é sobre o manifesto XML.

A seguir está um exemplo do manifesto XML.

 > [!NOTE]
 > All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## <a name="schema-versions"></a>Versões de esquema

Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.

O elemento **\<VersionOverrides\>** no manifesto é um exemplo disso. Todos os elementos definidos dentro de **\<VersionOverrides\>** substituirão o mesmo elemento na outra parte do manifesto. Isso significa que, sempre que possível, o Outlook usará o que está na seção **\<VersionOverrides\>** para configurar o suplemento. No entanto, se a versão do Outlook não oferecer suporte a uma determinada versão de **\<VersionOverrides\>**, o Outlook a ignorará e dependerá das informações no restante do manifesto. 

Com essa abordagem, os desenvolvedores não precisam criar vários manifestos individuais e podem, em vez disso, manter tudo definido em um único arquivo.

As versões atuais do esquema são:


|Versão|Descrição|
|:-----|:-----|
|v1.0|Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form. |
|v1.1|Dá suporte à versão 1.1 da API JavaScript do Office e **\<VersionOverrides\>**. Para suplementos do Outlook, isso acrescenta o suporte ao formulário de composição.|
|**\<VersionOverrides\>** 1.0|Dá suporte a versões posteriores da API JavaScript do Office. Oferece suporte aos comandos de suplemento.|
|**\<VersionOverrides\>** 1.1|Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.|

Este artigo abordará os requisitos de um manifesto da versão 1.1. Mesmo que seu manifesto de suplemento use o elemento **\<VersionOverrides\>**, ainda é importante incluir os elementos do manifesto da versão 1.1 para permitir que seu suplemento funcione com clientes mais antigos que não oferecem suporte a **\<VersionOverrides\>**.

> [!NOTE]
> Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.

## <a name="root-element"></a>Elemento root

O elemento root para o manifesto de suplementos do Outlook é **\<OfficeApp\>**. Esse elemento também declara o namespace padrão, a versão do esquema e o tipo de suplemento. Coloque todos os outros elementos no manifesto dentro de suas marcas de abertura e fechamento. A seguir está um exemplo do elemento root.


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## <a name="version"></a>Versão

This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.

## <a name="versionoverrides"></a>VersionOverrides

O elemento **\<VersionOverrides\>** é a localização das informações para os [comandos de suplemento](add-in-commands-for-outlook.md).

Este elemento também é onde os suplementos definem o suporte para [suplementos móveis](add-mobile-support.md).

Para uma discussão sobre este elemento, consulte [Criar comandos de suplementos em seu manifesto para Excel, PowerPoint e Word ](../develop/create-addin-commands.md).

## <a name="localization"></a>Localização

Alguns aspectos do suplemento precisam ser localizados para localidades diferentes, como nome, descrição e a URL que é carregada. Esses elementos podem ser localizados facilmente especificando o valor padrão e, em seguida, a localidade substitui no elemento **\<Resources\>** dentro do elemento **\<VersionOverrides\>**. O exemplo a seguir mostra como substituir uma imagem, uma URL e uma cadeia de caracteres.


```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

A referência de esquema contém informações completas sobre quais elementos podem ser localizados.

## <a name="hosts"></a>Hosts

Os suplementos do Outlook especificam o elemento **\<Hosts\>** como o seguinte:

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Isso é separado do elemento **\<Hosts\>** dentro do elemento **\<VersionOverrides\>**, que é discutido em [Criar comandos de suplemento em seu manifesto para Excel, PowerPoint e Word](../develop/create-addin-commands.md).

## <a name="requirements"></a>Requisitos

O elemento **\<Requirements\>** especifica o conjunto de APIs disponível para o suplemento. Para um suplemento do Outlook, o conjunto de requisitos deve ser uma Caixa de Correio e um valor de 1.1 ou superior. Confira a referência à API para obter a versão mais recente do conjunto de requisitos. Confira as [APIs de suplemento do Outlook](apis.md) para saber mais sobre os conjuntos de requisitos.

O elemento **\<Requirements\>** também pode aparecer no elemento **\<VersionOverrides\>**, permitindo que o suplemento especifique um requisito diferente quando carregado em clientes que dão suporte a **\<VersionOverrides\>**.

O exemplo a seguir usa o atributo **DefaultMinVersion** do elemento **\<Sets\>** para exigir o office.js versão 1.1 ou superior e o atributo **MinVersion** do elemento **\<Set\>** para exigir o conjunto de requisitos de Caixa de Correio versão 1.1.

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## <a name="form-settings"></a>Configurações de formulário

O elemento **\<FormSettings\>** é usado por clientes mais antigos do Outlook, que só dão suporte ao esquema 1.1 e não **\<VersionOverrides\>**. Com esse elemento, os desenvolvedores podem definir como o suplemento será exibido nesses clientes. Há duas partes: **ItemRead** e **ItemEdit**. **ItemRead** é usado para especificar como o suplemento aparece quando o usuário lê mensagens e compromissos. **ItemEdit** descreve como o suplemento será exibido enquanto o usuário estiver redigindo uma resposta, uma nova mensagem, um novo compromisso ou editando um compromisso do qual ele é o organizador.

Essas configurações estão diretamente relacionadas às regras de ativação no elemento **\<Rule\>**. Por exemplo, se um suplemento especificar que ele deve aparecer em uma mensagem no modo de composição, será necessário especificar um formulário **ItemEdit**.

Para saber mais, confira a Referência de esquema para manifestos de Suplementos do Office (v1.1).

## <a name="app-domains"></a>Domínios de aplicativo

O domínio da página inicial do suplemento que você especifica no elemento **\<SourceLocation\>** é o domínio padrão do suplemento. Sem usar os elementos **\<AppDomains\>** e **\<AppDomain\>**, se o suplemento tentar navegar para outro domínio, o navegador abrirá uma nova janela fora do painel do suplemento. Para permitir que o suplemento navegue até outro domínio dentro do painel do suplemento, adicione um elemento **\<AppDomains\>** e inclua cada domínio adicional em seu próprio subelemento **\<AppDomain\>** no manifesto do suplemento.

O exemplo a seguir especifica um domínio  `https://www.contoso2.com` como um segundo domínio para o qual o suplemento pode navegar dentro do painel do suplemento.

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Domínios de aplicativo também são necessários para habilitar o compartilhamento de cookies entre a janela pop-out e o suplemento em execução no cliente avançado.

A tabela a seguir descreve o comportamento do navegador quando o seu suplemento tenta navegar para uma URL fora do domínio padrão do suplemento.

|Cliente Outlook|Domínio definido<br>em AppDomains?|Comportamento do navegador|
|---|---|---|
|Todos os clientes|Sim|O link é aberto no painel de tarefas do suplemento.|
|– Outlook 2016 no Windows (perpétuo licenciado por volume)<br>- Outlook 2013 no Windows (perpétuo)|Não|O link é aberto no Internet Explorer 11.|
|Outros clientes|Não|O link é aberto no navegador padrão do usuário.|

Para mais detalhes, confira [Especificar os domínios que você deseja abrir na janela do suplemento](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).

## <a name="permissions"></a>Permissões

O elemento **\<Permissions\>** contém as permissões necessárias para o suplemento. Em geral, você deve especificar a permissão mínima exigida por seu suplemento, dependendo dos métodos exatos que você planeja usar. Por exemplo, um suplemento de email ativado em formulários de composição que apenas lê, mas não grava nas propriedades do item, por exemplo, [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), e não chama [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) para acessar quaisquer operações dos Serviços Web do Exchange, deve especificar a permissão **ReadItem**. Para obter detalhes sobre as permissões disponíveis, confira [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md).

**Modelo de permissões de quatro camadas para suplementos de email**

![Modelo de permissões de 4 camadas para o esquema de aplicativos de email v1.1.](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a>Regras de ativação

As regras de ativação são especificadas no elemento **\<Rule\>**. O elemento **\<Rule\>** pode aparecer como um filho do elemento **\<OfficeApp\>** em manifestos 1.1.

As regras de ativação podem ser usadas para ativar um suplemento com base em uma ou mais das seguintes condições no item selecionado no momento.

> [!NOTE]
> As regras de ativação somente se aplicam aos clientes que não dão suporte ao elemento **\<VersionOverrides\>**.

- O tipo de item e/ou a classe da mensagem

- A presença de um tipo específico de entidade conhecida, como um endereço ou número de telefone

- Uma correspondência de expressão regular no corpo, assunto ou endereço de email do remetente

- A presença de um anexo

Para obter detalhes e exemplos das regras de ativação, confira [Regras de ativação para suplementos do Outlook](activation-rules.md).

## <a name="next-steps-add-in-commands"></a>Próximas etapas: Comandos de suplemento

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

Para obter um exemplo de suplemento que defina comandos de suplementos, confira [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## <a name="next-steps-add-mobile-support"></a>Próximas etapas: Adicionar suporte móvel

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).

## <a name="see-also"></a>Confira também

- [Localização para suplementos do Office](../develop/localization.md)
- [Privacidade, permissões e segurança de suplementos do Outlook](privacy-and-security.md)
- [APIs de suplemento do Outlook](apis.md)
- [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)
- [Referência de esquema para manifestos de Suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)
- [Criar o design dos seus suplementos do Office](../design/add-in-design.md)
- [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md)
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md)