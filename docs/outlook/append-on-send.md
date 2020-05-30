---
title: Implementar Append no Send no suplemento do Outlook (visualização)
description: Saiba como implementar o recurso Append-on-Send em seu suplemento do Outlook.
ms.topic: article
ms.date: 05/26/2020
localization_priority: Normal
ms.openlocfilehash: ecbc1e043b6af2a0e0a6cd89de8cf4bfcec03943
ms.sourcegitcommit: 3a72d13c82b3d627691f4712d0d24b9e71bae9dc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/29/2020
ms.locfileid: "44415889"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a><span data-ttu-id="4289c-103">Implementar Append no Send no suplemento do Outlook (visualização)</span><span class="sxs-lookup"><span data-stu-id="4289c-103">Implement append on send in your Outlook add-in (preview)</span></span>

<span data-ttu-id="4289c-104">Ao final deste passo a passo, você terá um suplemento do Outlook que pode inserir um aviso de isenção de responsabilidade quando uma mensagem for enviada.</span><span class="sxs-lookup"><span data-stu-id="4289c-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4289c-105">No momento, esse recurso tem suporte para [Visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web e no Windows com uma assinatura do Office 365.</span><span class="sxs-lookup"><span data-stu-id="4289c-105">This feature is currently supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with an Office 365 subscription.</span></span> <span data-ttu-id="4289c-106">Veja [como visualizar o recurso Append-on-Send](#how-to-preview-the-append-on-send-feature) neste artigo para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="4289c-106">See [How to preview the append-on-send feature](#how-to-preview-the-append-on-send-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="4289c-107">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em suplementos de produção.</span><span class="sxs-lookup"><span data-stu-id="4289c-107">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-append-on-send-feature"></a><span data-ttu-id="4289c-108">Como visualizar o recurso Append-on-Send</span><span class="sxs-lookup"><span data-stu-id="4289c-108">How to preview the append-on-send feature</span></span>

<span data-ttu-id="4289c-109">Convidamos você a experimentar o recurso Append-on-Send!</span><span class="sxs-lookup"><span data-stu-id="4289c-109">We invite you to try out the append-on-send feature!</span></span> <span data-ttu-id="4289c-110">Informe-nos seus cenários e saiba como podemos melhorar enviando seus comentários por meio do GitHub (consulte a seção **comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="4289c-110">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="4289c-111">Para visualizar esse recurso:</span><span class="sxs-lookup"><span data-stu-id="4289c-111">To preview this feature:</span></span>

- <span data-ttu-id="4289c-112">Faça referência à biblioteca **beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="4289c-112">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="4289c-113">O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) para compilação TypeScript e IntelliSense é encontrado em CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="4289c-113">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="4289c-114">Você pode instalar esses tipos com o `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="4289c-114">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="4289c-115">Para o Windows, talvez seja necessário participar do [programa Office Insider](https://insider.office.com) para acessar versões mais recentes do Office.</span><span class="sxs-lookup"><span data-stu-id="4289c-115">For Windows, you may need to join the [Office Insider program](https://insider.office.com) to access more recent Office builds.</span></span>
- <span data-ttu-id="4289c-116">Para o Outlook na Web, [Configure o lançamento direcionado no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="4289c-116">For Outlook on the web, [configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="4289c-117">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="4289c-117">Set up your environment</span></span>

<span data-ttu-id="4289c-118">Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de suplemento com o gerador Yeoman para suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="4289c-118">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="4289c-119">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="4289c-119">Configure the manifest</span></span>

<span data-ttu-id="4289c-120">Para habilitar o recurso Append-on-Send no suplemento, você deve incluir a `AppendOnSend` permissão na coleção de [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="4289c-120">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="4289c-121">Para esse cenário, em vez de executar a `action` função ao escolher o botão **executar uma ação** , você executará a `appendOnSend` função.</span><span class="sxs-lookup"><span data-stu-id="4289c-121">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="4289c-122">Em seu editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="4289c-122">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="4289c-123">Abra o arquivo **manifest. xml** localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4289c-123">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="4289c-124">Selecione o `<VersionOverrides>` nó inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo seguinte XML.</span><span class="sxs-lookup"><span data-stu-id="4289c-124">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="4289c-125">Para saber mais sobre manifestos para suplementos do Outlook, confira [manifestos de suplementos do Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="4289c-125">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="4289c-126">Implementar a manipulação de Append-on-Send</span><span class="sxs-lookup"><span data-stu-id="4289c-126">Implement append-on-send handling</span></span>

<span data-ttu-id="4289c-127">Em seguida, implemente Append no evento Send.</span><span class="sxs-lookup"><span data-stu-id="4289c-127">Next, implement appending on the send event.</span></span>

<span data-ttu-id="4289c-128">Para este cenário, você implementará o acréscimo de um aviso de isenção de responsabilidade ao item quando o usuário enviar.</span><span class="sxs-lookup"><span data-stu-id="4289c-128">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="4289c-129">No mesmo projeto de início rápido, abra o arquivo **./src/Commands/Commands.js** no editor de código.</span><span class="sxs-lookup"><span data-stu-id="4289c-129">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="4289c-130">Após a `action` função, insira a seguinte função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4289c-130">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="4289c-131">No final do arquivo, adicione a instrução a seguir.</span><span class="sxs-lookup"><span data-stu-id="4289c-131">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="4289c-132">Experimente</span><span class="sxs-lookup"><span data-stu-id="4289c-132">Try it out</span></span>

1. <span data-ttu-id="4289c-133">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4289c-133">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="4289c-134">Quando você executar este comando, o servidor Web local será iniciado se ainda não estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="4289c-134">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="4289c-135">Siga as instruções em [Sideload suplementos do Outlook para teste](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="4289c-135">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="4289c-136">Crie uma nova mensagem e adicione-a à linha **para** .</span><span class="sxs-lookup"><span data-stu-id="4289c-136">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="4289c-137">No menu faixa de opções ou estouro, escolha **executar uma ação**.</span><span class="sxs-lookup"><span data-stu-id="4289c-137">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="4289c-138">Envie a mensagem e, em seguida, abra-a na pasta **caixa de entrada** ou **itens enviados** para exibir o aviso de isenção de responsabilidade anexado.</span><span class="sxs-lookup"><span data-stu-id="4289c-138">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Uma captura de tela de uma mensagem de exemplo com a isenção de responsabilidade anexada em enviar no Outlook na Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="4289c-140">Confira também</span><span class="sxs-lookup"><span data-stu-id="4289c-140">See also</span></span>

[<span data-ttu-id="4289c-141">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="4289c-141">Outlook add-in manifests</span></span>](manifests.md)
