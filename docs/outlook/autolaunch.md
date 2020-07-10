---
title: Configurar seu suplemento do Outlook para ativação baseada em evento (visualização)
description: Saiba como configurar seu suplemento do Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: 7b14dfb425223dbf486173e7df103d1a2cd09641
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093424"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="4142b-103">Configurar seu suplemento do Outlook para ativação baseada em evento (visualização)</span><span class="sxs-lookup"><span data-stu-id="4142b-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="4142b-104">Sem o recurso de ativação baseada em eventos, um usuário precisa iniciar explicitamente um suplemento para concluir suas tarefas.</span><span class="sxs-lookup"><span data-stu-id="4142b-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="4142b-105">Este recurso permite que seu suplemento Execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a todos os itens.</span><span class="sxs-lookup"><span data-stu-id="4142b-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="4142b-106">Você também pode se integrar ao painel de tarefas e à funcionalidade sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="4142b-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="4142b-107">No momento, os eventos com suporte são os seguintes.</span><span class="sxs-lookup"><span data-stu-id="4142b-107">At present, the supported events are as follows.</span></span>

- <span data-ttu-id="4142b-108">`OnNewMessageCompose`: Ao redigir uma nova mensagem (inclui responder, responder a todos e encaminhar)</span><span class="sxs-lookup"><span data-stu-id="4142b-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="4142b-109">`OnNewAppointmentOrganizer`: Ao criar um novo compromisso</span><span class="sxs-lookup"><span data-stu-id="4142b-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

<span data-ttu-id="4142b-110">Ao final deste passo a passo, você terá um suplemento que é executado sempre que uma nova mensagem é criada.</span><span class="sxs-lookup"><span data-stu-id="4142b-110">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4142b-111">Este recurso só tem suporte para [Visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="4142b-111">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="4142b-112">Veja [como visualizar o recurso de ativação baseado em eventos](#how-to-preview-the-event-based-activation-feature) neste artigo para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="4142b-112">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="4142b-113">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em suplementos de produção.</span><span class="sxs-lookup"><span data-stu-id="4142b-113">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="4142b-114">Como visualizar o recurso de ativação baseado em eventos</span><span class="sxs-lookup"><span data-stu-id="4142b-114">How to preview the event-based activation feature</span></span>

<span data-ttu-id="4142b-115">Convidamos você a experimentar o recurso de ativação baseada em eventos!</span><span class="sxs-lookup"><span data-stu-id="4142b-115">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="4142b-116">Informe-nos seus cenários e saiba como podemos melhorar enviando seus comentários por meio do GitHub (consulte a seção **comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="4142b-116">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="4142b-117">Para visualizar esse recurso:</span><span class="sxs-lookup"><span data-stu-id="4142b-117">To preview this feature:</span></span>

- <span data-ttu-id="4142b-118">Faça referência à biblioteca **beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="4142b-118">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="4142b-119">O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) para compilação TypeScript e IntelliSense é encontrado em CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="4142b-119">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="4142b-120">Você pode instalar esses tipos com o `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="4142b-120">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="4142b-121">Solicite acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário de solicitação](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="4142b-121">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="4142b-122">Vamos informar quando o locatário está pronto.</span><span class="sxs-lookup"><span data-stu-id="4142b-122">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="4142b-123">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="4142b-123">Set up your environment</span></span>

<span data-ttu-id="4142b-124">Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de suplemento com o gerador Yeoman para suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="4142b-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="4142b-125">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="4142b-125">Configure the manifest</span></span>

<span data-ttu-id="4142b-126">Para habilitar a ativação baseada em eventos do suplemento, você deve configurar o elemento de [tempos de execução](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="4142b-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="4142b-127">Por enquanto, `DesktopFormFactor` é o único fator de forma com suporte.</span><span class="sxs-lookup"><span data-stu-id="4142b-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="4142b-128">Em seu editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="4142b-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="4142b-129">Abra o arquivo **manifest.xml** localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4142b-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="4142b-130">Selecione o `<VersionOverrides>` nó inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo seguinte XML.</span><span class="sxs-lookup"><span data-stu-id="4142b-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
  </VersionOverrides>
</VersionOverrides>
```

<span data-ttu-id="4142b-131">O Outlook no Windows usa um arquivo JavaScript, enquanto o Outlook na Web usa um arquivo HTML que faz referência ao mesmo arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4142b-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="4142b-132">Você deve fornecer referências a esses arquivos no manifesto, uma vez que a plataforma do Outlook determina se deve usar HTML ou JavaScript com base no cliente do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4142b-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="4142b-133">Assim, para configurar a manipulação de eventos, forneça o local do HTML no `Runtime` elemento e, em seguida, em seu `Override` elemento filho, forneça o local do arquivo JavaScript embutido ou referenciado pelo HTML.</span><span class="sxs-lookup"><span data-stu-id="4142b-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="4142b-134">Para saber mais sobre manifestos para suplementos do Outlook, confira [manifestos de suplementos do Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="4142b-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="4142b-135">Implementar a manipulação de eventos</span><span class="sxs-lookup"><span data-stu-id="4142b-135">Implement event handling</span></span>

<span data-ttu-id="4142b-136">Você precisa implementar o tratamento dos eventos selecionados.</span><span class="sxs-lookup"><span data-stu-id="4142b-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="4142b-137">Neste cenário, você adicionará o tratamento para compor novos itens.</span><span class="sxs-lookup"><span data-stu-id="4142b-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="4142b-138">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="4142b-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="4142b-139">Após a `action` função, insira as seguintes funções de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4142b-139">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. <span data-ttu-id="4142b-140">No final do arquivo, adicione as instruções a seguir.</span><span class="sxs-lookup"><span data-stu-id="4142b-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="4142b-141">Experimente</span><span class="sxs-lookup"><span data-stu-id="4142b-141">Try it out</span></span>

1. <span data-ttu-id="4142b-142">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4142b-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="4142b-143">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="4142b-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="4142b-144">Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](sideload-outlook-add-ins-for-testing.md)para realizar o sideload do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4142b-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="4142b-145">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4142b-145">In Outlook on the web, create a new message.</span></span>

    ![Uma captura de tela de uma janela de mensagem no Outlook na Web com o assunto definido em redigir.](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="4142b-147">Comportamento e limitações de ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="4142b-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="4142b-148">Os suplementos que ativam com base em eventos são projetados para serem executados de forma curta, até 330 segundos.</span><span class="sxs-lookup"><span data-stu-id="4142b-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="4142b-149">Recomendamos que você tenha seu suplemento chamar o `event.completed` método para sinalizá-lo a concluir o processamento do evento de lançamento.</span><span class="sxs-lookup"><span data-stu-id="4142b-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="4142b-150">O suplemento também termina quando o usuário fecha a janela de redação.</span><span class="sxs-lookup"><span data-stu-id="4142b-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="4142b-151">Se o usuário tiver vários suplementos que se inscreveram no mesmo evento, a plataforma do Outlook iniciará os suplementos sem nenhuma ordem específica.</span><span class="sxs-lookup"><span data-stu-id="4142b-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="4142b-152">Atualmente, apenas cinco suplementos baseados em eventos podem estar ativamente em execução.</span><span class="sxs-lookup"><span data-stu-id="4142b-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="4142b-153">Qualquer suplemento adicional é enviado para uma fila e, em seguida, executado como os suplementos ativos anteriormente são concluídos ou desativados.</span><span class="sxs-lookup"><span data-stu-id="4142b-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="4142b-154">O usuário pode alternar ou navegar para fora do item de email atual onde o suplemento começou a ser executado.</span><span class="sxs-lookup"><span data-stu-id="4142b-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="4142b-155">O suplemento que foi iniciado concluirá a operação em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="4142b-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="4142b-156">Algumas APIs do Office.js que alteram ou alteram a interface do usuário não são permitidas a partir de suplementos baseados em eventos. A seguir estão as APIs bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="4142b-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="4142b-157">Em `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="4142b-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="4142b-158">Em `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="4142b-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="4142b-159">Em `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="4142b-159">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="4142b-160">Confira também</span><span class="sxs-lookup"><span data-stu-id="4142b-160">See also</span></span>

[<span data-ttu-id="4142b-161">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="4142b-161">Outlook add-in manifests</span></span>](manifests.md)
