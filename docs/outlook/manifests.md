---
title: Manifestos do suplemento do Outlook
description: O manifesto descreve como um suplemento do Outlook se integra a clientes do Outlook; inclui um exemplo.
ms.date: 10/31/2019
localization_priority: Priority
ms.openlocfilehash: f4d60919db15c4f470ecccac634abee94973bb6c
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324937"
---
# <a name="outlook-add-in-manifests"></a><span data-ttu-id="09304-103">Manifestos do suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-103">Outlook add-in manifests</span></span>

<span data-ttu-id="09304-p101">Um suplemento do Outlook é composto por dois componentes: o manifesto de suplemento XML e uma página da Web, compatível com a biblioteca JavaScript para Suplementos do Office (office.js). O manifesto descreve como o suplemento integra-se a clientes do Outlook. Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="09304-p101">An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.</span></span>

 > [!NOTE]
 > <span data-ttu-id="09304-p102">Todos os valores da URL no exemplo a seguir começam com "https://appdemo.contoso.com". Esse valor é um espaço reservado. Em um manifesto válido real, esses valores contêm URLs https da Web válidas.</span><span class="sxs-lookup"><span data-stu-id="09304-p102">All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.</span></span>

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

## <a name="schema-versions"></a><span data-ttu-id="09304-110">Versões de esquema</span><span class="sxs-lookup"><span data-stu-id="09304-110">Schema versions</span></span>

<span data-ttu-id="09304-p103">Nem todos os clientes do Outlook oferecem suporte aos recursos mais recentes e alguns usuários terão uma versão mais antiga do Outlook. Ter versões do esquema permite que os desenvolvedores compilem suplementos compatíveis com versões anteriores, usando os recursos mais recentes quando estiverem disponíveis, mas ainda funcionando em versões mais antigas.</span><span class="sxs-lookup"><span data-stu-id="09304-p103">Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.</span></span>

<span data-ttu-id="09304-p104">O elemento **VersionOverrides** no manifesto é um exemplo disso. Todos os elementos definidos dentro de **VersionOverrides** substituirão o mesmo elemento na outra parte do manifesto. Isso significa que, sempre que possível, o Outlook usará o que está na seção **VersionOverrides** para configurar o suplemento. No entanto, se a versão do Outlook não oferecer suporte a uma determinada versão de **VersionOverrides**, o Outlook a ignorará e dependerá das informações no restante do manifesto.</span><span class="sxs-lookup"><span data-stu-id="09304-p104">The **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest.</span></span> 

<span data-ttu-id="09304-117">Com essa abordagem, os desenvolvedores não precisam criar vários manifestos individuais e podem, em vez disso, manter tudo definido em um único arquivo.</span><span class="sxs-lookup"><span data-stu-id="09304-117">This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.</span></span>

<span data-ttu-id="09304-118">As versões atuais do esquema são:</span><span class="sxs-lookup"><span data-stu-id="09304-118">The current versions of the schema are:</span></span>


|<span data-ttu-id="09304-119">Versão</span><span class="sxs-lookup"><span data-stu-id="09304-119">Version</span></span>|<span data-ttu-id="09304-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="09304-120">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="09304-121">v1.0</span><span class="sxs-lookup"><span data-stu-id="09304-121">v1.0</span></span>|<span data-ttu-id="09304-p105">Oferece suporte à versão 1.0 da API JavaScript do Office. Para suplementos do Outlook, isso oferece suporte ao formulário de leitura.</span><span class="sxs-lookup"><span data-stu-id="09304-p105">Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form.</span></span> |
|<span data-ttu-id="09304-124">v1.1</span><span class="sxs-lookup"><span data-stu-id="09304-124">v1.1</span></span>|<span data-ttu-id="09304-p106">Oferece suporte à versão 1.1 da API do Office JavaScript e **VersionOverrides**. Para suplementos do Outlook, isso adiciona suporte ao formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="09304-p106">Supports version 1.1 of the Office JavaScript API and **VersionOverrides**. For Outlook add-ins, this adds support for compose form.</span></span>|
|<span data-ttu-id="09304-127">**VersionOverrides** 1.0</span><span class="sxs-lookup"><span data-stu-id="09304-127">**VersionOverrides** 1.0</span></span>|<span data-ttu-id="09304-p107">Oferece suporte a versões posteriores da API JavaScript do Office. Isso oferece suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="09304-p107">Supports later versions of the Office JavaScript API. This supports add-in commands.</span></span>|
|<span data-ttu-id="09304-130">**VersionOverrides** 1.1</span><span class="sxs-lookup"><span data-stu-id="09304-130">**VersionOverrides** 1.1</span></span>|<span data-ttu-id="09304-p108">Oferece suporte a versões posteriores da API JavaScript do Office. Isso oferece suporte a comandos de suplemento e adiciona suporte para recursos mais recentes, como [painéis de tarefas fixáveis](pinnable-taskpane.md) e suplementos para dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="09304-p108">Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.</span></span>|

<span data-ttu-id="09304-p109">Este artigo abordará os requisitos de um manifesto da versão 1.1. Mesmo que seu manifesto de suplemento use o elemento **VersionOverrides**, ainda é importante incluir os elementos do manifesto da versão 1.1 para permitir que seu suplemento funcione com clientes mais antigos que não oferecem suporte a **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="09304-p109">This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **VersionOverrides** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.</span></span>

> [!NOTE]
> <span data-ttu-id="09304-p110">O Outlook usa um esquema para validar manifestos. O esquema requer que os elementos no manifesto apareçam em uma ordem específica. Se você incluir elementos fora do pedido exigido, poderá ocorrer erros ao fazer o carregamento lateral do seu suplemento. Você pode carregar o [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) para ajudar a criar seu manifesto com elementos na ordem necessária.</span><span class="sxs-lookup"><span data-stu-id="09304-p110">Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.</span></span>

## <a name="root-element"></a><span data-ttu-id="09304-139">Elemento root</span><span class="sxs-lookup"><span data-stu-id="09304-139">Root element</span></span>

<span data-ttu-id="09304-p111">O elemento raiz do manifesto de suplementos do Outlook é **OfficeApp**. Esse elemento também declara o namespace padrão, a versão do esquema e o tipo de suplemento. Coloque todos os outros elementos no manifesto dentro de suas marcas de abertura e fechamento. Veja a seguir um exemplo do elemento Root:</span><span class="sxs-lookup"><span data-stu-id="09304-p111">The root element for the Outlook add-in manifest is **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:</span></span>


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

## <a name="version"></a><span data-ttu-id="09304-144">Version</span><span class="sxs-lookup"><span data-stu-id="09304-144">Version</span></span>

<span data-ttu-id="09304-p112">Esta é a versão do suplemento específico. Se um desenvolvedor atualiza algo no manifesto, a versão deve ser incrementada também. Dessa forma, quando o novo manifesto é instalado, ele substitui o existente, e o usuário recebe a nova funcionalidade. Se esse suplemento foi enviado para o repositório, o novo manifesto precisa ser re-enviado e revalidado. Em seguida, usuários desse suplemento obterão o novo manifesto atualizado automaticamente em algumas horas, depois da aprovação.</span><span class="sxs-lookup"><span data-stu-id="09304-p112">This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.</span></span>

<span data-ttu-id="09304-p113">Se as permissões solicitadas do suplemento mudarem, os usuários serão solicitados a atualizar e novamente concordar com o suplemento. Se o administrador tiver instalado esse suplemento para toda a organização, o administrador precisará primeiro concordar novamente. Os usuários continuarão a ver a funcionalidade antiga enquanto isso.</span><span class="sxs-lookup"><span data-stu-id="09304-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.</span></span>

## <a name="versionoverrides"></a><span data-ttu-id="09304-153">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="09304-153">VersionOverrides</span></span>

<span data-ttu-id="09304-p114">O elemento **VersionOverrides** é o local das informações de comandos do suplemento. Saiba mais sobre esse elemento em [Definir comandos de suplemento no manifesto](../develop/define-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="09304-p114">The **VersionOverrides** element is the location of information for add-in commands. For more information about this element, see [Define add-in commands in your manifest](../develop/define-add-in-commands.md).</span></span>

<span data-ttu-id="09304-156">Este elemento também é onde os suplementos definem o suporte para [suplementos móveis](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="09304-156">This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).</span></span>

## <a name="localization"></a><span data-ttu-id="09304-157">Localização</span><span class="sxs-lookup"><span data-stu-id="09304-157">Localization</span></span>

<span data-ttu-id="09304-p115">Alguns aspectos do suplemento precisam ser localizados para localidades diferentes, como nome, descrição e a URL que é carregada. Esses elementos podem ser localizados facilmente especificando o valor padrão e, em seguida, a localidade substitui no elemento **Resources** dentro do elemento **VersionOverrides**. Veja a seguir como substituir uma imagem, uma URL e uma cadeia de caracteres:</span><span class="sxs-lookup"><span data-stu-id="09304-p115">Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:</span></span>


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

<span data-ttu-id="09304-161">A referência de esquema contém informações completas sobre quais elementos podem ser localizados.</span><span class="sxs-lookup"><span data-stu-id="09304-161">The schema reference contains full information on which elements can be localized.</span></span>

## <a name="hosts"></a><span data-ttu-id="09304-162">Hosts</span><span class="sxs-lookup"><span data-stu-id="09304-162">Hosts</span></span>

<span data-ttu-id="09304-163">Os suplementos do Outlook especificam o elemento **Hosts** da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="09304-163">Outlook add-ins specify the **Hosts** element like the following.</span></span>

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

<span data-ttu-id="09304-164">Isso é separado do elemento **Hosts** dentro do elemento **VersionOverrides**, que é discutido em [Definir comandos de suplemento em seu manifesto](../develop/define-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="09304-164">This is separate from the **Hosts** element inside the **VersionOverrides** element, which is discussed in [Define add-in commands in your manifest](../develop/define-add-in-commands.md).</span></span>

## <a name="requirements"></a><span data-ttu-id="09304-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="09304-165">Requirements</span></span>

<span data-ttu-id="09304-p116">O elemento **Requirements** especifica o conjunto de APIs disponível para o suplemento. Para um suplemento do Outlook, o conjunto de requisitos deve ser uma caixa de correio e um valor de 1.1 ou superior. Confira a referência à API para obter a versão mais recente do conjunto de requisitos. Saiba mais sobre os conjuntos de requisitos em [APIs de suplementos do Outlook](apis.md).</span><span class="sxs-lookup"><span data-stu-id="09304-p116">The **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.</span></span>

<span data-ttu-id="09304-170">O elemento **Requirements** também pode aparecer no elemento **VersionOverrides**, permitindo que o suplemento especifique um requisito diferente quando for carregado em clientes que oferecem suporte a **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="09304-170">The **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.</span></span>

<span data-ttu-id="09304-171">O exemplo a seguir usa o atributo **DefaultMinVersion** do elemento **Sets** a fim de exigir a versão 1.1 ou superior do office.js, e o atributo **MinVersion** do elemento **Set** a fim de exigir a versão 1.1 do conjunto de requisitos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="09304-171">The following example uses the **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.</span></span>

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

## <a name="form-settings"></a><span data-ttu-id="09304-172">Configurações de formulário</span><span class="sxs-lookup"><span data-stu-id="09304-172">Form settings</span></span>

<span data-ttu-id="09304-p117">O elemento **FormSettings** é usado por clientes mais antigos do Outlook, que oferecem suporte apenas ao esquema 1.1 e não a **VersionOverrides**. Com esse elemento, os desenvolvedores podem definir como o suplemento será exibido nesses clientes. Há duas partes: **ItemRead** e **ItemEdit**. **ItemRead** é usado para especificar como o suplemento será exibido quando o usuário ler mensagens e compromissos. **ItemEdit** descreve como o suplemento será exibido enquanto o usuário estiver redigindo uma resposta, uma nova mensagem, um novo compromisso ou editando um compromisso do qual seja organizador.</span><span class="sxs-lookup"><span data-stu-id="09304-p117">The **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.</span></span>

<span data-ttu-id="09304-p118">Essas configurações estão diretamente relacionadas às regras de ativação no elemento **Rule**. Por exemplo, se um suplemento especificar que ele deve aparecer em uma mensagem no modo de redação, será necessário especificar um formulário **ItemEdit**.</span><span class="sxs-lookup"><span data-stu-id="09304-p118">These settings are directly related to the activation rules in the **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.</span></span>

<span data-ttu-id="09304-180">Para saber mais, confira a Referência de esquema para manifestos de Suplementos do Office (v1.1).</span><span class="sxs-lookup"><span data-stu-id="09304-180">For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../overview/add-in-manifests.md).</span></span>

## <a name="app-domains"></a><span data-ttu-id="09304-181">Domínios de aplicativo</span><span class="sxs-lookup"><span data-stu-id="09304-181">App domains</span></span>

<span data-ttu-id="09304-p119">O domínio da página inicial do suplemento que você especifica no elemento **SourceLocation** é o domínio padrão do suplemento. Sem usar os elementos **AppDomains** e **AppDomain**, se o suplemento tentar navegar para outro domínio, o navegador abrirá uma nova janela fora do painel do suplemento. Para permitir que o suplemento navegue para outro domínio dentro do painel do suplemento, adicione um elemento **AppDomains** e inclua cada domínio adicional em seu próprio subelemento **AppDomain** no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09304-p119">The domain of the add-in start page that you specify in the **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.</span></span>

<span data-ttu-id="09304-185">O exemplo a seguir especifica um domínio `https://www.contoso2.com` como um segundo domínio ao qual o suplemento pode navegar dentro do painel do suplemento:</span><span class="sxs-lookup"><span data-stu-id="09304-185">The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:</span></span>

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

<span data-ttu-id="09304-186">Domínios de aplicativo também são necessários para habilitar o compartilhamento de cookies entre a janela pop-out e o suplemento em execução no cliente avançado.</span><span class="sxs-lookup"><span data-stu-id="09304-186">App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.</span></span>

<span data-ttu-id="09304-187">A tabela a seguir descreve o comportamento do navegador quando o seu suplemento tenta navegar para uma URL fora do domínio padrão do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09304-187">The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.</span></span>

|<span data-ttu-id="09304-188">Cliente Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-188">Outlook client</span></span>|<span data-ttu-id="09304-189">Domínio definido</span><span class="sxs-lookup"><span data-stu-id="09304-189">Domain defined</span></span><br><span data-ttu-id="09304-190">em AppDomains?</span><span class="sxs-lookup"><span data-stu-id="09304-190">in AppDomains?</span></span>|<span data-ttu-id="09304-191">Comportamento do navegador</span><span class="sxs-lookup"><span data-stu-id="09304-191">Browser behavior</span></span>|
|---|---|---|
|<span data-ttu-id="09304-192">Todos os clientes</span><span class="sxs-lookup"><span data-stu-id="09304-192">All clients</span></span>|<span data-ttu-id="09304-193">Sim</span><span class="sxs-lookup"><span data-stu-id="09304-193">Yes</span></span>|<span data-ttu-id="09304-194">O link é aberto no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09304-194">Link opens in add-in task pane.</span></span>|
|<span data-ttu-id="09304-195">Outlook 2016 para Windows (compra única)</span><span class="sxs-lookup"><span data-stu-id="09304-195">Outlook 2016 on Windows (one-time purchase)</span></span><br><span data-ttu-id="09304-196">Outlook 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="09304-196">Outlook 2013 on Windows</span></span>|<span data-ttu-id="09304-197">Não</span><span class="sxs-lookup"><span data-stu-id="09304-197">No</span></span>|<span data-ttu-id="09304-198">O link é aberto no Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="09304-198">Link opens in Internet Explorer 11.</span></span>|
|<span data-ttu-id="09304-199">Outros clientes</span><span class="sxs-lookup"><span data-stu-id="09304-199">Other clients</span></span>|<span data-ttu-id="09304-200">Não</span><span class="sxs-lookup"><span data-stu-id="09304-200">No</span></span>|<span data-ttu-id="09304-201">O link é aberto no navegador padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="09304-201">Link opens in user's default browser.</span></span>|

<span data-ttu-id="09304-202">Para mais detalhes, confira [Especificar os domínios que você deseja abrir na janela do suplemento](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span><span class="sxs-lookup"><span data-stu-id="09304-202">For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span></span>

## <a name="permissions"></a><span data-ttu-id="09304-203">Permissões</span><span class="sxs-lookup"><span data-stu-id="09304-203">Permissions</span></span>

<span data-ttu-id="09304-p120">O elemento **Permissions** contém as permissões necessárias para o suplemento. Em geral, você deve especificar a permissão mínima exigida por seu suplemento, dependendo dos métodos exatos que você planeja usar. Por exemplo, um suplemento de email ativado em formulários de redação que apenas lê, mas não grava nas propriedades do item, como [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), e não chama [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para acessar quaisquer operações dos Serviços Web do Exchange, deve especificar a permissão **ReadItem**. Confira detalhes sobre as permissões disponíveis em [Noções básicas sobre permissões de suplementos do Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="09304-p120">The **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and does not call [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="09304-208">**Modelo de permissões de quatro camadas para suplementos de email**</span><span class="sxs-lookup"><span data-stu-id="09304-208">**Four-tier permissions model for mail add-ins**</span></span>

![Modelo de permissões de quatro camadas para o esquema de aplicativos de correio v1.1](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a><span data-ttu-id="09304-210">Regras de ativação</span><span class="sxs-lookup"><span data-stu-id="09304-210">Activation rules</span></span>

<span data-ttu-id="09304-p121">As regras de ativação são especificadas no elemento **Rule**. O elemento **Rule** pode aparecer como um filho do elemento **OfficeApp** em manifestos 1.1.</span><span class="sxs-lookup"><span data-stu-id="09304-p121">Activation rules are specified in the **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests.</span></span>

<span data-ttu-id="09304-213">As regras de ativação podem ser usadas para ativar um suplemento com base em uma ou mais das seguintes condições no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="09304-213">Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="09304-214">As regras de ativação somente se aplicam aos clientes que não dão suporte ao elemento **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="09304-214">Activation rules only apply to clients that do not support the **VersionOverrides** element.</span></span>

- <span data-ttu-id="09304-215">O tipo de item e/ou a classe da mensagem</span><span class="sxs-lookup"><span data-stu-id="09304-215">The item type and/or message class</span></span>

- <span data-ttu-id="09304-216">A presença de um tipo específico de entidade conhecida, como um endereço ou número de telefone</span><span class="sxs-lookup"><span data-stu-id="09304-216">The presence of a specific type of known entity, such as an address or phone number</span></span>

- <span data-ttu-id="09304-217">Uma correspondência de expressão regular no corpo, assunto ou endereço de email do remetente</span><span class="sxs-lookup"><span data-stu-id="09304-217">A regular expression match in the body, subject, or sender email address</span></span>

- <span data-ttu-id="09304-218">A presença de um anexo</span><span class="sxs-lookup"><span data-stu-id="09304-218">The presence of an attachment</span></span>

<span data-ttu-id="09304-219">Para obter detalhes e exemplos das regras de ativação, confira [Regras de ativação para suplementos do Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="09304-219">For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>


## <a name="next-steps-add-in-commands"></a><span data-ttu-id="09304-220">Próximas etapas: Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="09304-220">Next steps: Add-in commands</span></span>

<span data-ttu-id="09304-p122">Após definir um manifesto básico, [defina os comandos de suplemento para seu suplemento](../develop/define-add-in-commands.md). Os comandos de suplemento apresentam um botão na faixa de opções para que os usuários possam ativar o suplemento de uma maneira simples e intuitiva. Para saber mais, confira [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="09304-p122">After defining a basic manifest, [define add-in commands for your add-in](../develop/define-add-in-commands.md). Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="09304-224">Para obter um exemplo de suplemento que defina comandos de suplementos, confira [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span><span class="sxs-lookup"><span data-stu-id="09304-224">For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span></span>

## <a name="next-steps-add-mobile-support"></a><span data-ttu-id="09304-225">Próximas etapas: Adicionar suporte móvel</span><span class="sxs-lookup"><span data-stu-id="09304-225">Next steps: Add mobile support</span></span>

<span data-ttu-id="09304-p123">Os suplementos podem, opcionalmente, adicionar suporte para o Outlook mobile. O Outlook Mobile dá suporte a comandos de suplemento de maneira semelhante ao Outlook no Windows e no Mac. Para saber mais, veja [Adicionar suporte para comandos de suplementos no Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="09304-p123">Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="09304-229">Confira também</span><span class="sxs-lookup"><span data-stu-id="09304-229">See also</span></span>

- [<span data-ttu-id="09304-230">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="09304-230">Localization for Office Add-ins</span></span>](../develop/localization.md)
- [<span data-ttu-id="09304-231">Privacidade, permissões e segurança de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-231">Privacy, permissions, and security for Outlook add-ins</span></span>](privacy-and-security.md)
- [<span data-ttu-id="09304-232">APIs de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-232">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="09304-233">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="09304-233">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="09304-234">Referência de esquema para manifestos de Suplementos do Office (versão 1.1)</span><span class="sxs-lookup"><span data-stu-id="09304-234">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="09304-235">Criar o design dos seus suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="09304-235">Design your Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="09304-236">Noções básicas sobre permissões de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-236">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="09304-237">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="09304-237">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="09304-238">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="09304-238">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
