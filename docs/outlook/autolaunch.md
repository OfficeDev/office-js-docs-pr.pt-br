---
title: Configurar seu Outlook para ativação baseada em eventos (visualização)
description: Saiba como configurar seu Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 04/29/2021
localization_priority: Normal
ms.openlocfilehash: 45f9ff16b3aca0a1fb8f3a8ee3d9ffa8e0f33ea2
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100296"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="cbe7a-103">Configurar seu Outlook para ativação baseada em eventos (visualização)</span><span class="sxs-lookup"><span data-stu-id="cbe7a-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="cbe7a-104">Sem o recurso de ativação baseada em evento, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="cbe7a-105">Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="cbe7a-106">Você também pode se integrar ao painel de tarefas e à funcionalidade sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="cbe7a-107">Atualmente, os seguintes eventos são suportados.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-107">At present, the following events are supported.</span></span>

|<span data-ttu-id="cbe7a-108">Evento</span><span class="sxs-lookup"><span data-stu-id="cbe7a-108">Event</span></span>|<span data-ttu-id="cbe7a-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="cbe7a-109">Description</span></span>|
|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="cbe7a-110">Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar), mas não ao editar, por exemplo, um rascunho.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-110">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="cbe7a-111">Ao criar um novo compromisso, mas não ao editar um existente.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-111">On creating a new appointment but not on editing an existing one.</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="cbe7a-112">Ao adicionar ou remover anexos ao compor uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-112">On adding or removing attachments while composing a message.</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="cbe7a-113">Ao adicionar ou remover anexos durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-113">On adding or removing attachments while composing an appointment.</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="cbe7a-114">Ao adicionar ou remover destinatários ao compor uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-114">On adding or removing recipients while composing a message.</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="cbe7a-115">Ao adicionar ou remover participantes durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-115">On adding or removing attendees while composing an appointment.</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="cbe7a-116">Ao alterar data/hora durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-116">On changing date/time while composing an appointment.</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="cbe7a-117">Ao adicionar, alterar ou remover os detalhes de recorrência ao compor um compromisso.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-117">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="cbe7a-118">Se a data/hora for alterada, `OnAppointmentTimeChanged` o evento também será acionado.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-118">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="cbe7a-119">Ao descartar uma notificação ao compor uma mensagem ou item de compromisso.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-119">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="cbe7a-120">Somente o complemento que adicionou a notificação será notificado.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-120">Only the add-in that added the notification will be notified.</span></span>|

<span data-ttu-id="cbe7a-121">No final deste passo a passo, você terá um complemento que é executado sempre que um novo item é criado e define o assunto.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-121">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cbe7a-122">Esse recurso só é [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) suportado para visualização Outlook na Web e no Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-122">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="cbe7a-123">Confira [Como visualizar o recurso de ativação](#how-to-preview-the-event-based-activation-feature) baseada em evento neste artigo para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-123">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="cbe7a-124">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-124">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="cbe7a-125">Como visualizar o recurso de ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="cbe7a-125">How to preview the event-based activation feature</span></span>

<span data-ttu-id="cbe7a-126">Convidamos você a experimentar o recurso de ativação baseada em evento!</span><span class="sxs-lookup"><span data-stu-id="cbe7a-126">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="cbe7a-127">Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback por meio GitHub (consulte a seção **Comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="cbe7a-127">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="cbe7a-128">Para visualizar esse recurso:</span><span class="sxs-lookup"><span data-stu-id="cbe7a-128">To preview this feature:</span></span>

- <span data-ttu-id="cbe7a-129">Para Outlook na Web:</span><span class="sxs-lookup"><span data-stu-id="cbe7a-129">For Outlook on the web:</span></span>
  - <span data-ttu-id="cbe7a-130">[Configure a versão direcionada em seu Microsoft 365 locatário](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="cbe7a-130">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="cbe7a-131">Fazer referência **à biblioteca beta** no CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="cbe7a-131">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="cbe7a-132">O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para a compilação typeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="cbe7a-132">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="cbe7a-133">Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="cbe7a-133">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="cbe7a-134">Para Outlook no Windows: o build mínimo necessário é 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-134">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="cbe7a-135">Participe do [programa Office Insider](https://insider.office.com) para acessar Office beta.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-135">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="cbe7a-136">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="cbe7a-136">Set up your environment</span></span>

<span data-ttu-id="cbe7a-137">Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido que cria um projeto de complemento com o gerador Yeoman para Office Desempois.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-137">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="cbe7a-138">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="cbe7a-138">Configure the manifest</span></span>

<span data-ttu-id="cbe7a-139">Para habilitar a ativação baseada em evento do seu add-in, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` no nó do manifesto.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-139">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="cbe7a-140">Por enquanto, `DesktopFormFactor` é o único fator de formulário suportado.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-140">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="cbe7a-141">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-141">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="cbe7a-142">Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-142">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="cbe7a-143">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas abertas e próximas) e substitua-o pelo XML a seguir e salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-143">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="cbe7a-144">Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-144">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="cbe7a-145">Você deve fornecer referências a ambos os arquivos no nó do manifesto como a plataforma Outlook finalmente determina se deve usar HTML ou JavaScript com base no cliente `Resources` Outlook.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-145">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="cbe7a-146">Como tal, para configurar o tratamento de eventos, forneça o local do HTML no elemento e, em seguida, em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado `Runtime` `Override` pelo HTML.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-146">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="cbe7a-147">Para saber mais sobre manifestos para Outlook de Outlook, [consulte Outlook manifestos de complemento.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="cbe7a-147">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="cbe7a-148">Implementar o tratamento de eventos</span><span class="sxs-lookup"><span data-stu-id="cbe7a-148">Implement event handling</span></span>

<span data-ttu-id="cbe7a-149">Você precisa implementar o tratamento para os eventos selecionados.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-149">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="cbe7a-150">Nesse cenário, você adicionará a manipulação para compor novos itens.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-150">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="cbe7a-151">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-151">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="cbe7a-152">Após a `action` função, insira as seguintes funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-152">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="cbe7a-153">Para que as funções funcionem Outlook **na Web** com esse projeto gerado pelo gerador Yeoman para Office Desempois, adicione as instruções a seguir no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-153">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="cbe7a-154">Para que as funções funcionem em Outlook **em** Windows, adicione o código JavaScript a seguir no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-154">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="cbe7a-155">**Observação**: verificar `Office.actions` se as Outlook na Web ignoram essas instruções.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-155">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="cbe7a-156">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-156">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="cbe7a-157">Experimente</span><span class="sxs-lookup"><span data-stu-id="cbe7a-157">Try it out</span></span>

1. <span data-ttu-id="cbe7a-158">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-158">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="cbe7a-159">Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-159">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="cbe7a-160">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-160">In Outlook on the web, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem Outlook na Web com o assunto definido como redação](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="cbe7a-162">Em Outlook no Windows, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-162">In Outlook on Windows, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem em Outlook no Windows com o assunto definido como redação](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="cbe7a-164">Se você vir o erro "Não podemos abrir esse complemento do localhost", você precisará habilitar uma isenção de loopback.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-164">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="cbe7a-165">Close Outlook.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-165">Close Outlook.</span></span>
    > 2. <span data-ttu-id="cbe7a-166">Abra o **Gerenciador de Tarefas** e certifique-se de que o **msoadfs.exe** não está em execução.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-166">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="cbe7a-167">Execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-167">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="cbe7a-168">Reinicie o Outlook.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-168">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="cbe7a-169">Depuração</span><span class="sxs-lookup"><span data-stu-id="cbe7a-169">Debug</span></span>

<span data-ttu-id="cbe7a-170">À medida que você implementa sua própria funcionalidade, talvez seja necessário depurar seu código.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-170">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="cbe7a-171">Para obter orientações sobre como depurar a ativação de um add-in baseado em evento, consulte [Depurar](debug-autolaunch.md)seu Outlook de evento.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-171">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="cbe7a-172">Comportamento e limitações de ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="cbe7a-172">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="cbe7a-173">Os complementos que são ativados com base em eventos devem ser curtos, leves e não invasivos possíveis.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-173">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="cbe7a-174">Para sinalizar que o seu complemento concluiu o processamento do evento de lançamento, recomendamos que você chame o método de seu `event.completed` complemento.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-174">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="cbe7a-175">Se essa chamada não for feita, o complemento terá um tempo limite de aproximadamente 300 segundos, o tempo máximo permitido para a execução de complementos baseados em eventos. O complemento também termina quando o usuário fecha a janela de composição.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-175">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="cbe7a-176">Se o usuário tiver vários complementos que se inscrevem no mesmo evento, a plataforma Outlook iniciará os complementos em nenhuma ordem específica.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-176">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="cbe7a-177">Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-177">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="cbe7a-178">Quaisquer complementos adicionais são pressionados para uma fila e executados conforme os complementos ativos anteriormente são concluídos ou desativados.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-178">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="cbe7a-179">O usuário pode alternar ou navegar para longe do item de email atual onde o complemento começou a ser executado.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-179">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="cbe7a-180">O complemento que foi lançado terminará sua operação em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="cbe7a-180">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="cbe7a-181">Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas de complementos baseados em eventos. Veja a seguir as APIs bloqueadas:</span><span class="sxs-lookup"><span data-stu-id="cbe7a-181">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="cbe7a-182">Em `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="cbe7a-182">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="cbe7a-183">Em `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="cbe7a-183">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="cbe7a-184">Em `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="cbe7a-184">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="cbe7a-185">Em `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="cbe7a-185">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="cbe7a-186">Confira também</span><span class="sxs-lookup"><span data-stu-id="cbe7a-186">See also</span></span>

<span data-ttu-id="cbe7a-187">[Outlook manifestos](manifests.md) 
 de complemento [Como depurar os complementos baseados em eventos](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="cbe7a-187">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
