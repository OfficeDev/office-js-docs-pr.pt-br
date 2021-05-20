---
title: Configure seu Outlook complemento para ativação baseada em eventos (visualização)
description: Saiba como configurar seu Outlook complemento para ativação baseada em eventos.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555226"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="a0541-103">Configure seu Outlook complemento para ativação baseada em eventos (visualização)</span><span class="sxs-lookup"><span data-stu-id="a0541-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="a0541-104">Sem o recurso de ativação baseado em eventos, o usuário precisa lançar explicitamente um complemento para concluir suas tarefas.</span><span class="sxs-lookup"><span data-stu-id="a0541-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="a0541-105">Esse recurso permite que seu complemento execute tarefas com base em determinados eventos, particularmente para operações que se aplicam a cada item.</span><span class="sxs-lookup"><span data-stu-id="a0541-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="a0541-106">Você também pode se integrar com a funcionalidade painel de tarefas e sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="a0541-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="a0541-107">Ao final deste passo a passo, você terá um complemento que é executado sempre que um novo item for criado e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="a0541-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a0541-108">Este recurso só é suportado para [visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) em Outlook na web e em Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="a0541-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="a0541-109">Para obter mais detalhes, consulte [Como visualizar o recurso de ativação baseado](#how-to-preview-the-event-based-activation-feature) em eventos neste artigo.</span><span class="sxs-lookup"><span data-stu-id="a0541-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="a0541-110">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.</span><span class="sxs-lookup"><span data-stu-id="a0541-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="a0541-111">Eventos com suporte</span><span class="sxs-lookup"><span data-stu-id="a0541-111">Supported events</span></span>

<span data-ttu-id="a0541-112">No momento, os seguintes eventos são apoiados.</span><span class="sxs-lookup"><span data-stu-id="a0541-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="a0541-113">Evento</span><span class="sxs-lookup"><span data-stu-id="a0541-113">Event</span></span>|<span data-ttu-id="a0541-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0541-114">Description</span></span>|<span data-ttu-id="a0541-115">Clientes</span><span class="sxs-lookup"><span data-stu-id="a0541-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="a0541-116">Ao compor uma nova mensagem (inclui resposta, resposta a todos e adiante), mas não na edição, por exemplo, de um rascunho.</span><span class="sxs-lookup"><span data-stu-id="a0541-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="a0541-117">Windows, web</span><span class="sxs-lookup"><span data-stu-id="a0541-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="a0541-118">Ao criar um novo compromisso, mas não na edição de um já existente.</span><span class="sxs-lookup"><span data-stu-id="a0541-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="a0541-119">Windows, web</span><span class="sxs-lookup"><span data-stu-id="a0541-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="a0541-120">Ao adicionar ou remover anexos enquanto compõe uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0541-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="a0541-121">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="a0541-122">Ao adicionar ou remover anexos enquanto compõe uma consulta.</span><span class="sxs-lookup"><span data-stu-id="a0541-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="a0541-123">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="a0541-124">Ao adicionar ou remover destinatários enquanto compõe uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0541-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="a0541-125">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="a0541-126">Ao adicionar ou remover os participantes enquanto compõe uma consulta.</span><span class="sxs-lookup"><span data-stu-id="a0541-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="a0541-127">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="a0541-128">Ao alterar a data/hora enquanto compõe uma consulta.</span><span class="sxs-lookup"><span data-stu-id="a0541-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="a0541-129">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="a0541-130">Ao adicionar, alterar ou remover os detalhes de recorrência enquanto compõe uma consulta.</span><span class="sxs-lookup"><span data-stu-id="a0541-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="a0541-131">Se a data/hora for alterada, o `OnAppointmentTimeChanged` evento também será acionado.</span><span class="sxs-lookup"><span data-stu-id="a0541-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="a0541-132">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="a0541-133">Ao descartar uma notificação enquanto compõe uma mensagem ou item de nomeação.</span><span class="sxs-lookup"><span data-stu-id="a0541-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="a0541-134">Apenas o complemento que adicionou a notificação será notificado.</span><span class="sxs-lookup"><span data-stu-id="a0541-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="a0541-135">Windows</span><span class="sxs-lookup"><span data-stu-id="a0541-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="a0541-136">Como visualizar o recurso de ativação baseado em eventos</span><span class="sxs-lookup"><span data-stu-id="a0541-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="a0541-137">Convidamos você a experimentar o recurso de ativação baseado em eventos!</span><span class="sxs-lookup"><span data-stu-id="a0541-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="a0541-138">Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback através de GitHub (veja a seção **Feedback** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="a0541-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="a0541-139">Para visualizar este recurso:</span><span class="sxs-lookup"><span data-stu-id="a0541-139">To preview this feature:</span></span>

- <span data-ttu-id="a0541-140">Para Outlook na web:</span><span class="sxs-lookup"><span data-stu-id="a0541-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="a0541-141">[Configure a liberação direcionada no seu inquilino Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="a0541-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="a0541-142">Consulte a biblioteca **beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="a0541-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="a0541-143">O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) para compilação TypeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="a0541-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="a0541-144">Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="a0541-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="a0541-145">Para Outlook em Windows:</span><span class="sxs-lookup"><span data-stu-id="a0541-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="a0541-146">A construção mínima exigida é de 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="a0541-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="a0541-147">Junte-se ao [programa Office Insider](https://insider.office.com) para acesso a Office compilações beta.</span><span class="sxs-lookup"><span data-stu-id="a0541-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="a0541-148">Configure o registro.</span><span class="sxs-lookup"><span data-stu-id="a0541-148">Configure the registry.</span></span> <span data-ttu-id="a0541-149">Outlook inclui uma cópia local da produção e versões beta de Office.js em vez de carregar a partir do CDN.</span><span class="sxs-lookup"><span data-stu-id="a0541-149">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="a0541-150">Por padrão, a cópia de produção local da API é referenciada.</span><span class="sxs-lookup"><span data-stu-id="a0541-150">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="a0541-151">Para mudar para a cópia beta local das APIs javaScript Outlook, você precisa adicionar esta entrada de registro, caso contrário, as APIs beta podem não ser encontradas.</span><span class="sxs-lookup"><span data-stu-id="a0541-151">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="a0541-152">Crie a chave de registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` .</span><span class="sxs-lookup"><span data-stu-id="a0541-152">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="a0541-153">Adicione uma entrada nomeada `EnableBetaAPIsInJavaScript` e defina o valor para `1` .</span><span class="sxs-lookup"><span data-stu-id="a0541-153">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="a0541-154">A imagem a seguir mostra qual deve ser a aparência de registro.</span><span class="sxs-lookup"><span data-stu-id="a0541-154">The following image shows what the registry should look like.</span></span>

        ![Captura de tela do editor de registro com um valor-chave de registro EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="a0541-156">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="a0541-156">Set up your environment</span></span>

<span data-ttu-id="a0541-157">Complete o [Outlook início rápido](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto adicional com o gerador Yeoman para Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="a0541-157">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="a0541-158">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="a0541-158">Configure the manifest</span></span>

<span data-ttu-id="a0541-159">Para ativar a ativação baseada em eventos do seu complemento, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) no `VersionOverridesV1_1` nó do manifesto.</span><span class="sxs-lookup"><span data-stu-id="a0541-159">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="a0541-160">Por enquanto, `DesktopFormFactor` é o único fator de forma suportado.</span><span class="sxs-lookup"><span data-stu-id="a0541-160">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="a0541-161">Em seu editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="a0541-161">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="a0541-162">Abra o **arquivomanifest.xml** localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="a0541-162">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="a0541-163">Selecione todo o `<VersionOverrides>` nó (incluindo tags abertas e fechadas) e substitua-o pelo XML a seguir e salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="a0541-163">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

<span data-ttu-id="a0541-164">Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0541-164">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="a0541-165">Você deve fornecer referências a ambos os arquivos no `Resources` nó do manifesto, pois a plataforma Outlook determina em última instância se deve usar HTML ou JavaScript com base no Outlook cliente.</span><span class="sxs-lookup"><span data-stu-id="a0541-165">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="a0541-166">Como tal, para configurar o manuseio do evento, forneça a localização do HTML no `Runtime` elemento e, em seguida, em seu `Override` elemento filho forneça a localização do arquivo JavaScript ininlinado ou referenciado pelo HTML.</span><span class="sxs-lookup"><span data-stu-id="a0541-166">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="a0541-167">Para saber mais sobre manifestos para Outlook complementos, consulte [Outlook manifestos adicionais](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="a0541-167">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="a0541-168">Implementar o manuseio de eventos</span><span class="sxs-lookup"><span data-stu-id="a0541-168">Implement event handling</span></span>

<span data-ttu-id="a0541-169">Você tem que implementar o manuseio para seus eventos selecionados.</span><span class="sxs-lookup"><span data-stu-id="a0541-169">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="a0541-170">Neste cenário, você adicionará manuseio para compor novos itens.</span><span class="sxs-lookup"><span data-stu-id="a0541-170">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="a0541-171">A partir do mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="a0541-171">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="a0541-172">Após a `action` função, insira as seguintes funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0541-172">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="a0541-173">Adicione o seguinte código JavaScript no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="a0541-173">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="a0541-174">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="a0541-174">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a0541-175">Experimente</span><span class="sxs-lookup"><span data-stu-id="a0541-175">Try it out</span></span>

1. <span data-ttu-id="a0541-176">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="a0541-176">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="a0541-177">Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.</span><span class="sxs-lookup"><span data-stu-id="a0541-177">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="a0541-178">Se o seu complemento não tiver sido carregado automaticamente, siga as instruções em [sideload Outlook complementos para testar](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) para carregar manualmente o complemento Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0541-178">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="a0541-179">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0541-179">In Outlook on the web, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem em Outlook na web com o assunto definido na composição](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="a0541-181">Em Outlook no Windows, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0541-181">In Outlook on Windows, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem em Outlook em Windows com o assunto definido na composição](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="a0541-183">Se você estiver executando seu complemento do localhost e veja o erro "Sentimos muito, não podemos acessar *{seu-add-in-name-aqui}*.</span><span class="sxs-lookup"><span data-stu-id="a0541-183">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="a0541-184">Certifique-se de ter uma conexão de rede.</span><span class="sxs-lookup"><span data-stu-id="a0541-184">Make sure you have a network connection.</span></span> <span data-ttu-id="a0541-185">Se o problema continuar, tente novamente mais tarde.", você pode precisar habilitar uma isenção de loopback.</span><span class="sxs-lookup"><span data-stu-id="a0541-185">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="a0541-186">Close Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0541-186">Close Outlook.</span></span>
    > 1. <span data-ttu-id="a0541-187">Abra o **Gerenciador de Tarefas** e garanta que o processo **demsoadfsb.exe** não esteja em execução.</span><span class="sxs-lookup"><span data-stu-id="a0541-187">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="a0541-188">Execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="a0541-188">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="a0541-189">Reinicie o Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0541-189">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="a0541-190">depurar</span><span class="sxs-lookup"><span data-stu-id="a0541-190">Debug</span></span>

<span data-ttu-id="a0541-191">À medida que você faz alterações no manuseio do evento de lançamento em seu complemento, você deve estar ciente de que:</span><span class="sxs-lookup"><span data-stu-id="a0541-191">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="a0541-192">Se você atualizou o manifesto, [remova o complemento](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) e, em seguida, o der descarga novamente.</span><span class="sxs-lookup"><span data-stu-id="a0541-192">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="a0541-193">Se você fez alterações em outros arquivos além do manifesto, feche e reabra Outlook na Windows ou atualize a guia do navegador executando Outlook na web.</span><span class="sxs-lookup"><span data-stu-id="a0541-193">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="a0541-194">Ao implementar sua própria funcionalidade, você pode precisar depurar seu código.</span><span class="sxs-lookup"><span data-stu-id="a0541-194">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="a0541-195">Para obter orientações sobre como depurar a ativação complementa baseada em eventos, consulte [Depurar seu complemento de Outlook baseado em eventos](debug-autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="a0541-195">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="a0541-196">O registro de tempo de execução também está disponível para este recurso em Windows.</span><span class="sxs-lookup"><span data-stu-id="a0541-196">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="a0541-197">Para obter mais informações, consulte [Depurar seu complemento com o registro de tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows).</span><span class="sxs-lookup"><span data-stu-id="a0541-197">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="a0541-198">Implantar para usuários</span><span class="sxs-lookup"><span data-stu-id="a0541-198">Deploy to users</span></span>

<span data-ttu-id="a0541-199">Você pode implantar complementos baseados em eventos carregando o manifesto através do centro administrativo Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="a0541-199">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="a0541-200">No portal de administração, expanda a seção **Configurações** no painel de navegação e selecione **aplicativos Integrados**.</span><span class="sxs-lookup"><span data-stu-id="a0541-200">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="a0541-201">Na página **aplicativos Integrados,** escolha a Upload ação **de aplicativos personalizados.**</span><span class="sxs-lookup"><span data-stu-id="a0541-201">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Captura de tela da página de aplicativos integrados no centro administrativo Microsoft 365, incluindo a ação de aplicativos personalizados Upload](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="a0541-203">AppSource e lojas de inclientes: A capacidade de implantar complementos baseados em eventos ou atualizar complementos existentes para incluir o recurso de ativação baseado em eventos deve estar disponível em breve.</span><span class="sxs-lookup"><span data-stu-id="a0541-203">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a0541-204">Os complementos baseados em eventos são restritos apenas a implantações gerenciadas por administradores.</span><span class="sxs-lookup"><span data-stu-id="a0541-204">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="a0541-205">Por enquanto, os usuários não podem obter complementos baseados em eventos do AppSource ou lojas de incientes.</span><span class="sxs-lookup"><span data-stu-id="a0541-205">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="a0541-206">Comportamento e limitações de ativação baseadas em eventos</span><span class="sxs-lookup"><span data-stu-id="a0541-206">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="a0541-207">Espera-se que os manipuladores de eventos de lançamento adicionais sejam de curta duração, leves e não invasivos.</span><span class="sxs-lookup"><span data-stu-id="a0541-207">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="a0541-208">Após a ativação, seu complemento será cronometrados dentro de aproximadamente 300 segundos, o tempo máximo permitido para executar complementos baseados em eventos. Para sinalizar que seu complemento completou o processamento de um evento de lançamento, recomendamos que o manipulador associado chame o `event.completed` método.</span><span class="sxs-lookup"><span data-stu-id="a0541-208">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="a0541-209">(Observe que o código incluído após a `event.completed` declaração não é garantido para execução.) Cada vez que um evento que seu complemento é acionado, o complemento é reativado e executa o manipulador de eventos associado e a janela de tempo limite é redefinida.</span><span class="sxs-lookup"><span data-stu-id="a0541-209">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="a0541-210">O complemento termina depois que ele se esgota, ou o usuário fecha a janela de composição ou envia o item.</span><span class="sxs-lookup"><span data-stu-id="a0541-210">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="a0541-211">Se o usuário tiver vários complementos que se inscreveram no mesmo evento, a plataforma Outlook lança os complementos em nenhuma ordem específica.</span><span class="sxs-lookup"><span data-stu-id="a0541-211">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="a0541-212">Atualmente, apenas cinco complementos baseados em eventos podem estar sendo executados ativamente.</span><span class="sxs-lookup"><span data-stu-id="a0541-212">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="a0541-213">O usuário pode alternar ou navegar para longe do item de e-mail atual, onde o complemento começou a ser executado.</span><span class="sxs-lookup"><span data-stu-id="a0541-213">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="a0541-214">O complemento que foi lançado terminará sua operação em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="a0541-214">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="a0541-215">Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas a partir de complementos baseados em eventos. A seguir, as APIs bloqueadas:</span><span class="sxs-lookup"><span data-stu-id="a0541-215">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="a0541-216">Em `OfficeRuntime.auth` :</span><span class="sxs-lookup"><span data-stu-id="a0541-216">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="a0541-217">`getAccessToken`(somente Windows)</span><span class="sxs-lookup"><span data-stu-id="a0541-217">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="a0541-218">Em `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="a0541-218">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="a0541-219">Em `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="a0541-219">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="a0541-220">Em `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="a0541-220">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="a0541-221">Em `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="a0541-221">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="a0541-222">Confira também</span><span class="sxs-lookup"><span data-stu-id="a0541-222">See also</span></span>

- [<span data-ttu-id="a0541-223">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0541-223">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="a0541-224">Como depurar complementos baseados em eventos</span><span class="sxs-lookup"><span data-stu-id="a0541-224">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
