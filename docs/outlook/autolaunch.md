---
title: Configurar seu complemento do Outlook para ativação baseada em eventos (visualização)
description: Saiba como configurar seu complemento do Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 02/03/2021
localization_priority: Normal
ms.openlocfilehash: d9108b4debea5e59503f3c935a537e5fafde00c8
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234272"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="32082-103">Configurar seu complemento do Outlook para ativação baseada em eventos (visualização)</span><span class="sxs-lookup"><span data-stu-id="32082-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="32082-104">Sem o recurso de ativação baseada em eventos, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas.</span><span class="sxs-lookup"><span data-stu-id="32082-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="32082-105">Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item.</span><span class="sxs-lookup"><span data-stu-id="32082-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="32082-106">Você também pode integrar com o painel de tarefas e a funcionalidade sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="32082-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="32082-107">Atualmente, os seguintes eventos são suportados.</span><span class="sxs-lookup"><span data-stu-id="32082-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="32082-108">`OnNewMessageCompose`: Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar)</span><span class="sxs-lookup"><span data-stu-id="32082-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="32082-109">`OnNewAppointmentOrganizer`: Ao criar um novo compromisso</span><span class="sxs-lookup"><span data-stu-id="32082-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="32082-110">Esse recurso não **é** ativado na edição de um item, por exemplo, um rascunho ou um compromisso existente.</span><span class="sxs-lookup"><span data-stu-id="32082-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="32082-111">No final deste passo a passo, você terá um complemento que é executado sempre que uma nova mensagem é criada.</span><span class="sxs-lookup"><span data-stu-id="32082-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="32082-112">Esse recurso só tem suporte para [visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web e no Windows com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="32082-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="32082-113">Veja [como visualizar o recurso de ativação baseada em eventos](#how-to-preview-the-event-based-activation-feature) neste artigo para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="32082-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="32082-114">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.</span><span class="sxs-lookup"><span data-stu-id="32082-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="32082-115">Como visualizar o recurso de ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="32082-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="32082-116">Convidamos você a experimentar o recurso de ativação baseada em eventos!</span><span class="sxs-lookup"><span data-stu-id="32082-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="32082-117">Conheça seus cenários e como podemos melhorar nos fazendo comentários por meio do GitHub (confira a seção **Comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="32082-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="32082-118">Para visualizar esse recurso:</span><span class="sxs-lookup"><span data-stu-id="32082-118">To preview this feature:</span></span>

- <span data-ttu-id="32082-119">Para o Outlook na Web:</span><span class="sxs-lookup"><span data-stu-id="32082-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="32082-120">[Configure o lançamento direcionado no locatário do Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="32082-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="32082-121">Fazer referência **à biblioteca beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="32082-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="32082-122">O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para compilação de TypeScript e IntelliSense é encontrado na CDN e [definitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="32082-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="32082-123">Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="32082-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="32082-124">Para o Outlook no Windows: o build mínimo necessário é 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="32082-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="32082-125">Participe do [programa Office Insider](https://insider.office.com) para acessar as versões beta do Office.</span><span class="sxs-lookup"><span data-stu-id="32082-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="32082-126">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="32082-126">Set up your environment</span></span>

<span data-ttu-id="32082-127">Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.</span><span class="sxs-lookup"><span data-stu-id="32082-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="32082-128">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="32082-128">Configure the manifest</span></span>

<span data-ttu-id="32082-129">Para habilitar a ativação baseada em eventos do seu complemento, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) no `VersionOverridesV1_1` nó do manifesto.</span><span class="sxs-lookup"><span data-stu-id="32082-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="32082-130">Por enquanto, `DesktopFormFactor` é o único fator forma com suporte.</span><span class="sxs-lookup"><span data-stu-id="32082-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="32082-131">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="32082-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="32082-132">Abra o **manifest.xml** arquivo localizado na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="32082-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="32082-133">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.</span><span class="sxs-lookup"><span data-stu-id="32082-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
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
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
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
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
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
  </VersionOverrides>
</VersionOverrides>
```

<span data-ttu-id="32082-134">O Outlook no Windows usa um arquivo JavaScript, enquanto o Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="32082-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="32082-135">Você deve fornecer referências a ambos os arquivos no nó do manifesto, pois a plataforma do Outlook finalmente determina se deve usar HTML ou JavaScript com base no `Resources` cliente do Outlook.</span><span class="sxs-lookup"><span data-stu-id="32082-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="32082-136">Dessa forma, para configurar a manipulação de eventos, forneça o local do HTML no elemento e, em seguida, em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado `Runtime` `Override` pelo HTML.</span><span class="sxs-lookup"><span data-stu-id="32082-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="32082-137">Para saber mais sobre manifestos para os complementos do Outlook, confira [manifestos de complementos do Outlook.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="32082-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="32082-138">Implementar a manipulação de eventos</span><span class="sxs-lookup"><span data-stu-id="32082-138">Implement event handling</span></span>

<span data-ttu-id="32082-139">Você precisa implementar a manipulação para os eventos selecionados.</span><span class="sxs-lookup"><span data-stu-id="32082-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="32082-140">Neste cenário, você adicionará a manipulação para composição de novos itens.</span><span class="sxs-lookup"><span data-stu-id="32082-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="32082-141">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="32082-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="32082-142">Após a `action` função, insira as seguintes funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="32082-142">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }
    
          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }
    ```

1. <span data-ttu-id="32082-143">Para que as funções funcionem no Outlook na **Web** com esse projeto gerado pelo gerador Yeoman para Os Complementos do Office, adicione as instruções a seguir no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="32082-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="32082-144">Para as funções funcionarem no **Outlook no Windows,** adicione o seguinte código JavaScript no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="32082-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="32082-145">**Observação:** a verificação `Office.actions` garante que o Outlook na Web ignore essas instruções.</span><span class="sxs-lookup"><span data-stu-id="32082-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="32082-146">Experimente</span><span class="sxs-lookup"><span data-stu-id="32082-146">Try it out</span></span>

1. <span data-ttu-id="32082-147">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="32082-147">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="32082-148">Quando você executar esse comando, o servidor Web local será iniciar (se ainda não estiver em execução) e o seu complemento será sideloaded.</span><span class="sxs-lookup"><span data-stu-id="32082-148">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="32082-149">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="32082-149">In Outlook on the web, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem no Outlook na Web com o assunto definido na composição](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="32082-151">No Outlook no Windows, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="32082-151">In Outlook on Windows, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem no Outlook no Windows com o assunto definido na composição](../images/outlook-win-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="32082-153">Comportamento e limitações da ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="32082-153">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="32082-154">Espera-se que os complementos ativados com base em eventos sejam de curta duração, leve e o mais não ofensivo possível.</span><span class="sxs-lookup"><span data-stu-id="32082-154">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="32082-155">Para sinalizar que o seu complemento concluiu o processamento do evento de lançamento, recomendamos que você chame o `event.completed` método.</span><span class="sxs-lookup"><span data-stu-id="32082-155">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="32082-156">Se essa chamada não for feita, o tempo limite do complemento será de aproximadamente 300 segundos, o período máximo permitido para a execução de complementos baseados em eventos. O complemento também termina quando o usuário fecha a janela de redação.</span><span class="sxs-lookup"><span data-stu-id="32082-156">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="32082-157">Se o usuário tiver vários complementos que se inscrevem no mesmo evento, a plataforma do Outlook inicia os complementos sem uma ordem específica.</span><span class="sxs-lookup"><span data-stu-id="32082-157">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="32082-158">Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente.</span><span class="sxs-lookup"><span data-stu-id="32082-158">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="32082-159">Quaisquer outros complementos são pressionados para uma fila e executados à medida que os complementos ativos anteriormente são concluídos ou desativados.</span><span class="sxs-lookup"><span data-stu-id="32082-159">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="32082-160">O usuário pode alternar ou sair do item de email atual onde o complemento começou a ser executado.</span><span class="sxs-lookup"><span data-stu-id="32082-160">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="32082-161">O complemento que foi lançado concluirá sua operação em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="32082-161">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="32082-162">Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas a partir de complementos baseados em eventos. Veja a seguir as APIs bloqueadas:</span><span class="sxs-lookup"><span data-stu-id="32082-162">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="32082-163">Em `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="32082-163">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="32082-164">Em `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="32082-164">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="32082-165">Em `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="32082-165">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="32082-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="32082-166">See also</span></span>

[<span data-ttu-id="32082-167">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="32082-167">Outlook add-in manifests</span></span>](manifests.md)