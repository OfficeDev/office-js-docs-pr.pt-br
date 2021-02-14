---
title: Implementar append-on-send no seu complemento do Outlook
description: Saiba como implementar o recurso append-on-send no seu complemento do Outlook.
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 8b69fbbaef1d0f060f0675fe5c4948a70d935b7a
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234286"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="de14f-103">Implementar append-on-send no seu complemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="de14f-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="de14f-104">No final deste passo a passo, você terá um complemento do Outlook que pode inserir um aviso de isenção de responsabilidade quando uma mensagem é enviada.</span><span class="sxs-lookup"><span data-stu-id="de14f-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="de14f-105">O suporte para esse recurso foi introduzido no conjunto de requisitos 1.9.</span><span class="sxs-lookup"><span data-stu-id="de14f-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="de14f-106">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="de14f-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="de14f-107">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="de14f-107">Set up your environment</span></span>

<span data-ttu-id="de14f-108">Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.</span><span class="sxs-lookup"><span data-stu-id="de14f-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="de14f-109">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="de14f-109">Configure the manifest</span></span>

<span data-ttu-id="de14f-110">Para habilitar o recurso append-on-send no seu complemento, você deve incluir a permissão na coleção `AppendOnSend` de [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="de14f-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="de14f-111">Para esse cenário, em vez de executar a função ao escolher o botão Executar uma `action` ação, você executará a  `appendOnSend` função.</span><span class="sxs-lookup"><span data-stu-id="de14f-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="de14f-112">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="de14f-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="de14f-113">Abra o **manifest.xml** arquivo localizado na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="de14f-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="de14f-114">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.</span><span class="sxs-lookup"><span data-stu-id="de14f-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.3">
            <bt:Set Name="Mailbox" />
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                      <Label resid="TaskpaneButton.Label" />
                      <Supertip>
                        <Title resid="TaskpaneButton.Label" />
                        <Description resid="TaskpaneButton.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16" />
                        <bt:Image size="32" resid="Icon.32x32" />
                        <bt:Image size="80" resid="Icon.80x80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <SourceLocation resid="Taskpane.Url" />
                      </Action>
                    </Control>
                    <Control xsi:type="Button" id="ActionButton">
                      <Label resid="ActionButton.Label"/>
                      <Supertip>
                        <Title resid="ActionButton.Label"/>
                        <Description resid="ActionButton.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16"/>
                        <bt:Image size="32" resid="Icon.32x32"/>
                        <bt:Image size="80" resid="Icon.80x80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

            </DesktopFormFactor>
          </Host>
        </Hosts>
        <Resources>
          <bt:Images>
            <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
            <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
            <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
            <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
            <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
            <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
            <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
            <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
          </bt:LongStrings>
        </Resources>
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> <span data-ttu-id="de14f-115">Para saber mais sobre manifestos para os complementos do Outlook, confira [manifestos de complementos do Outlook.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="de14f-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="de14f-116">Implementar a manipulação append-on-send</span><span class="sxs-lookup"><span data-stu-id="de14f-116">Implement append-on-send handling</span></span>

<span data-ttu-id="de14f-117">Em seguida, implemente a aplicação de acordo com o evento de envio.</span><span class="sxs-lookup"><span data-stu-id="de14f-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="de14f-118">Se o seu complemento [ `ItemSend` ](outlook-on-send-addins.md)também implementa a manipulação de eventos ao enviar usando , chamar o manipulador Ao enviar retornará um erro, pois não há suporte para `AppendOnSendAsync` esse cenário.</span><span class="sxs-lookup"><span data-stu-id="de14f-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="de14f-119">Para esse cenário, você implementará a aplicação de um aviso de isenção de responsabilidade ao item quando o usuário enviar.</span><span class="sxs-lookup"><span data-stu-id="de14f-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="de14f-120">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="de14f-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="de14f-121">Após a `action` função, insira a seguinte função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="de14f-121">After the `action` function, insert the following JavaScript function.</span></span>

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. <span data-ttu-id="de14f-122">No final do arquivo, adicione a instrução a seguir.</span><span class="sxs-lookup"><span data-stu-id="de14f-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="de14f-123">Experimente</span><span class="sxs-lookup"><span data-stu-id="de14f-123">Try it out</span></span>

1. <span data-ttu-id="de14f-124">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="de14f-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="de14f-125">Quando você executar esse comando, o servidor Web local será lançado se ele ainda não estiver em execução e o seu complemento será sideloaded.</span><span class="sxs-lookup"><span data-stu-id="de14f-125">When you run this command, the local web server will start if it's not already running and your add-in will be sideloaded.</span></span> 

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="de14f-126">Crie uma nova mensagem e adicione-se à **linha** Para.</span><span class="sxs-lookup"><span data-stu-id="de14f-126">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="de14f-127">No menu faixa de opções ou estouro, escolha **Executar uma ação.**</span><span class="sxs-lookup"><span data-stu-id="de14f-127">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="de14f-128">Envie a mensagem e abra-a  na pasta Itens Enviados ou na Caixa de Entrada para exibir o aviso de isenção de responsabilidade. </span><span class="sxs-lookup"><span data-stu-id="de14f-128">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Uma captura de tela de uma mensagem de exemplo com o aviso de isenção de responsabilidade anexado ao enviar no Outlook na Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="de14f-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="de14f-130">See also</span></span>

[<span data-ttu-id="de14f-131">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="de14f-131">Outlook add-in manifests</span></span>](manifests.md)
