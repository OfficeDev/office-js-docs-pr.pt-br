---
title: Configurar seu Outlook para ativação baseada em eventos
description: Saiba como configurar seu Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 05/20/2021
localization_priority: Normal
ms.openlocfilehash: a0d0b27c9b49132024c78577a4432d85542cf76f
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590502"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a><span data-ttu-id="4d6a7-103">Configurar seu Outlook para ativação baseada em eventos</span><span class="sxs-lookup"><span data-stu-id="4d6a7-103">Configure your Outlook add-in for event-based activation</span></span>

<span data-ttu-id="4d6a7-104">Sem o recurso de ativação baseada em evento, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="4d6a7-105">Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="4d6a7-106">Você também pode se integrar ao painel de tarefas e à funcionalidade sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="4d6a7-107">No final deste passo a passo, você terá um complemento que é executado sempre que um novo item é criado e define o assunto.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!NOTE]
> <span data-ttu-id="4d6a7-108">O suporte para esse recurso foi introduzido no [conjunto de requisitos 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="4d6a7-108">Support for this feature was introduced in [requirement set 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="4d6a7-109">Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-events"></a><span data-ttu-id="4d6a7-110">Eventos com suporte</span><span class="sxs-lookup"><span data-stu-id="4d6a7-110">Supported events</span></span>

<span data-ttu-id="4d6a7-111">Atualmente, os seguintes eventos são suportados.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-111">At present, the following events are supported.</span></span>

|<span data-ttu-id="4d6a7-112">Evento</span><span class="sxs-lookup"><span data-stu-id="4d6a7-112">Event</span></span>|<span data-ttu-id="4d6a7-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="4d6a7-113">Description</span></span>|<span data-ttu-id="4d6a7-114">Clientes</span><span class="sxs-lookup"><span data-stu-id="4d6a7-114">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="4d6a7-115">Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar), mas não ao editar, por exemplo, um rascunho.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-115">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="4d6a7-116">Windows, web</span><span class="sxs-lookup"><span data-stu-id="4d6a7-116">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="4d6a7-117">Ao criar um novo compromisso, mas não ao editar um existente.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-117">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="4d6a7-118">Windows, web</span><span class="sxs-lookup"><span data-stu-id="4d6a7-118">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`\*|<span data-ttu-id="4d6a7-119">Ao adicionar ou remover anexos ao compor uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-119">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="4d6a7-120">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-120">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`\*|<span data-ttu-id="4d6a7-121">Ao adicionar ou remover anexos durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-121">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="4d6a7-122">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-122">Windows</span></span>|
|`OnMessageRecipientsChanged`\*|<span data-ttu-id="4d6a7-123">Ao adicionar ou remover destinatários ao compor uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-123">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="4d6a7-124">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-124">Windows</span></span>|
|`OnAppointmentAttendeesChanged`\*|<span data-ttu-id="4d6a7-125">Ao adicionar ou remover participantes durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-125">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="4d6a7-126">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-126">Windows</span></span>|
|`OnAppointmentTimeChanged`\*|<span data-ttu-id="4d6a7-127">Ao alterar data/hora durante a composição de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-127">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="4d6a7-128">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-128">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`\*|<span data-ttu-id="4d6a7-129">Ao adicionar, alterar ou remover os detalhes de recorrência ao compor um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-129">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="4d6a7-130">Se a data/hora for alterada, `OnAppointmentTimeChanged` o evento também será acionado.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-130">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="4d6a7-131">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-131">Windows</span></span>|
|`OnInfoBarDismissClicked`\*|<span data-ttu-id="4d6a7-132">Ao descartar uma notificação ao compor uma mensagem ou item de compromisso.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-132">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="4d6a7-133">Somente o complemento que adicionou a notificação será notificado.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-133">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="4d6a7-134">Windows</span><span class="sxs-lookup"><span data-stu-id="4d6a7-134">Windows</span></span>|

> [!IMPORTANT]
> <span data-ttu-id="4d6a7-135">\*Esse evento só tem suporte para [visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook no Windows com uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-135">\* This event is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="4d6a7-136">Para obter mais detalhes, [consulte Como visualizar](#how-to-preview) neste artigo.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-136">For more details, see [How to preview](#how-to-preview) in this article.</span></span>
>
> <span data-ttu-id="4d6a7-137">Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-137">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview"></a><span data-ttu-id="4d6a7-138">Como visualizar</span><span class="sxs-lookup"><span data-stu-id="4d6a7-138">How to preview</span></span>

<span data-ttu-id="4d6a7-139">Convidamos você a experimentar os novos eventos!</span><span class="sxs-lookup"><span data-stu-id="4d6a7-139">We invite you to try out the new events!</span></span> <span data-ttu-id="4d6a7-140">Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback por meio GitHub (consulte a seção **Comentários** no final desta página).</span><span class="sxs-lookup"><span data-stu-id="4d6a7-140">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="4d6a7-141">Para visualizar esse recurso:</span><span class="sxs-lookup"><span data-stu-id="4d6a7-141">To preview this feature:</span></span>

- <span data-ttu-id="4d6a7-142">Para Outlook na Web:</span><span class="sxs-lookup"><span data-stu-id="4d6a7-142">For Outlook on the web:</span></span>
  - <span data-ttu-id="4d6a7-143">[Configure a versão direcionada em seu Microsoft 365 locatário](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="4d6a7-143">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="4d6a7-144">Fazer referência **à biblioteca beta** no CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="4d6a7-144">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="4d6a7-145">O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para a compilação typeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="4d6a7-145">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="4d6a7-146">Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="4d6a7-146">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="4d6a7-147">Para Outlook no Windows:</span><span class="sxs-lookup"><span data-stu-id="4d6a7-147">For Outlook on Windows:</span></span>
  - <span data-ttu-id="4d6a7-148">O build mínimo necessário é 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-148">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="4d6a7-149">Participe do [programa Office Insider](https://insider.office.com) para acessar Office beta.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-149">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="4d6a7-150">Configure o Registro.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-150">Configure the registry.</span></span> <span data-ttu-id="4d6a7-151">Outlook inclui uma cópia local das versões de produção e beta do Office.js em vez de carregar do CDN.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-151">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="4d6a7-152">Por padrão, a cópia de produção local da API é referenciada.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-152">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="4d6a7-153">Para alternar para a cópia beta local das APIs javaScript Outlook, você precisa adicionar essa entrada do Registro, caso contrário, as APIs beta podem não ser encontradas.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-153">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="4d6a7-154">Crie a chave do Registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` .</span><span class="sxs-lookup"><span data-stu-id="4d6a7-154">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="4d6a7-155">Adicione uma entrada chamada `EnableBetaAPIsInJavaScript` e desmarcar o valor como `1` .</span><span class="sxs-lookup"><span data-stu-id="4d6a7-155">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="4d6a7-156">A imagem a seguir mostra qual deve ser a aparência de registro.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-156">The following image shows what the registry should look like.</span></span>

        ![Captura de tela do editor do Registro com um valor de chave do Registro EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="4d6a7-158">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="4d6a7-158">Set up your environment</span></span>

<span data-ttu-id="4d6a7-159">Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido que cria um projeto de complemento com o gerador Yeoman para Office Desempois.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-159">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="4d6a7-160">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="4d6a7-160">Configure the manifest</span></span>

<span data-ttu-id="4d6a7-161">Para habilitar a ativação baseada em evento do seu add-in, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) `VersionOverridesV1_1` no nó do manifesto.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-161">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="4d6a7-162">Por enquanto, `DesktopFormFactor` é o único fator de formulário suportado.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-162">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="4d6a7-163">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-163">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="4d6a7-164">Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-164">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="4d6a7-165">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas abertas e próximas) e substitua-o pelo XML a seguir e salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-165">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="4d6a7-166">Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-166">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="4d6a7-167">Você deve fornecer referências a ambos os arquivos no nó do manifesto como a plataforma Outlook finalmente determina se deve usar HTML ou JavaScript com base no cliente `Resources` Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-167">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="4d6a7-168">Como tal, para configurar o tratamento de eventos, forneça o local do HTML no elemento e, em seguida, em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado `Runtime` `Override` pelo HTML.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-168">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="4d6a7-169">Para saber mais sobre manifestos para Outlook de Outlook, [consulte Outlook manifestos de complemento.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="4d6a7-169">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="4d6a7-170">Implementar o tratamento de eventos</span><span class="sxs-lookup"><span data-stu-id="4d6a7-170">Implement event handling</span></span>

<span data-ttu-id="4d6a7-171">Você precisa implementar o tratamento para os eventos selecionados.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-171">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="4d6a7-172">Nesse cenário, você adicionará a manipulação para compor novos itens.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-172">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="4d6a7-173">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-173">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="4d6a7-174">Após a `action` função, insira as seguintes funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-174">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="4d6a7-175">Adicione o código JavaScript a seguir no final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-175">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="4d6a7-176">Salve suas alterações.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-176">Save your changes.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4d6a7-177">Windows: no momento, as importações não são suportadas no arquivo JavaScript onde você implementa o tratamento para a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-177">Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="4d6a7-178">Experimente</span><span class="sxs-lookup"><span data-stu-id="4d6a7-178">Try it out</span></span>

1. <span data-ttu-id="4d6a7-179">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-179">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="4d6a7-180">Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-180">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="4d6a7-181">Se o seu add-in não foi automaticamente sideload, siga as instruções em [Sideload Outlook add-ins](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) for testing to manually sideload the add-in in Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-181">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="4d6a7-182">No Outlook na Web, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-182">In Outlook on the web, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem Outlook na Web com o assunto definido como redação](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="4d6a7-184">Em Outlook no Windows, crie uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-184">In Outlook on Windows, create a new message.</span></span>

    ![Captura de tela de uma janela de mensagem em Outlook no Windows com o assunto definido como redação](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="4d6a7-186">Se você estiver executando o seu complemento no localhost e vir o erro "Lamentamos, não foi possível acessar *{your-add-in-name-here}*.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-186">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="4d6a7-187">Certifique-se de ter uma conexão de rede.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-187">Make sure you have a network connection.</span></span> <span data-ttu-id="4d6a7-188">Se o problema continuar, tente novamente mais tarde.", talvez seja necessário habilitar uma isenção de loopback.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-188">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="4d6a7-189">Close Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-189">Close Outlook.</span></span>
    > 1. <span data-ttu-id="4d6a7-190">Abra o **Gerenciador de Tarefas** e certifique-se de que o **msoadfsb.exe** não está em execução.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-190">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="4d6a7-191">Execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-191">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="4d6a7-192">Reinicie o Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-192">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="4d6a7-193">Depuração</span><span class="sxs-lookup"><span data-stu-id="4d6a7-193">Debug</span></span>

<span data-ttu-id="4d6a7-194">À medida que você faz alterações no tratamento de eventos de início no seu complemento, você deve estar ciente de que:</span><span class="sxs-lookup"><span data-stu-id="4d6a7-194">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="4d6a7-195">Se você atualizou o manifesto, [remova o complemento e,](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) em seguida, o sideload novamente.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-195">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="4d6a7-196">Se você fez alterações em arquivos que não o manifesto, feche e reabra o Outlook no Windows ou atualize a guia do navegador executando Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-196">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="4d6a7-197">Ao implementar sua própria funcionalidade, talvez seja necessário depurar seu código.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-197">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="4d6a7-198">Para obter orientações sobre como depurar a ativação de um add-in baseado em evento, consulte [Depurar](debug-autolaunch.md)seu Outlook de evento.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-198">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="4d6a7-199">O log de tempo de execução também está disponível para esse recurso no Windows.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-199">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="4d6a7-200">Para obter mais informações, consulte [Depurar seu add-in com o log de tempo de execução.](../testing/runtime-logging.md#runtime-logging-on-windows)</span><span class="sxs-lookup"><span data-stu-id="4d6a7-200">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="4d6a7-201">Implantar para usuários</span><span class="sxs-lookup"><span data-stu-id="4d6a7-201">Deploy to users</span></span>

<span data-ttu-id="4d6a7-202">Você pode implantar os complementos baseados em eventos carregando o manifesto por meio do Microsoft 365 de administração.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-202">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="4d6a7-203">No portal de administração, expanda a **seção Configurações** no painel de navegação e selecione **Aplicativos integrados.**</span><span class="sxs-lookup"><span data-stu-id="4d6a7-203">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="4d6a7-204">Na página **Aplicativos integrados,** escolha a ação Upload **aplicativos personalizados.**</span><span class="sxs-lookup"><span data-stu-id="4d6a7-204">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Captura de tela da página Aplicativos integrados no centro de administração Microsoft 365, incluindo a ação Upload aplicativos personalizados](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="4d6a7-206">AppSource e armazenamentos de inclientes: a capacidade de implantar os complementos baseados em eventos ou atualizar os complementos existentes para incluir o recurso de ativação baseada em evento deve estar disponível em breve.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-206">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4d6a7-207">Os complementos baseados em eventos são restritos apenas a implantações gerenciadas pelo administrador.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-207">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="4d6a7-208">Por enquanto, os usuários não podem obter os complementos baseados em eventos do AppSource ou dos armazenamentos de inclientes.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-208">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="4d6a7-209">Comportamento e limitações de ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="4d6a7-209">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="4d6a7-210">Espera-se que os manipuladores de eventos de início do add-in sejam curtos, leves e não invasivos possíveis.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-210">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="4d6a7-211">Após a ativação, o seu complemento terá um tempo limite de aproximadamente 300 segundos, o tempo máximo permitido para a execução de complementos baseados em eventos. Para sinalizar que o seu complemento concluiu o processamento de um evento de lançamento, recomendamos que o manipulador associado chame o `event.completed` método.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-211">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="4d6a7-212">(Observe que o código incluído após `event.completed` a instrução não é garantido para ser executado.) Sempre que um evento que seu complemento lida é disparado, o complemento é reativado e executa o manipulador de eventos associado e a janela de tempo de tempo é redefinida.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-212">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="4d6a7-213">O complemento termina após o tempo final, ou o usuário fecha a janela de redação ou envia o item.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-213">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="4d6a7-214">Se o usuário tiver vários complementos que se inscrevem no mesmo evento, a plataforma Outlook iniciará os complementos em nenhuma ordem específica.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-214">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="4d6a7-215">Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-215">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="4d6a7-216">O usuário pode alternar ou navegar para longe do item de email atual onde o complemento começou a ser executado.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-216">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="4d6a7-217">O complemento que foi lançado terminará sua operação em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-217">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="4d6a7-218">As importações não são suportadas no arquivo JavaScript em que você implementa o tratamento para a ativação baseada em eventos no cliente Windows cliente.</span><span class="sxs-lookup"><span data-stu-id="4d6a7-218">Imports are not supported in the JavaScript file where you implement the handling for event-based activation in the Windows client.</span></span>

<span data-ttu-id="4d6a7-219">Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas de complementos baseados em eventos. Veja a seguir as APIs bloqueadas:</span><span class="sxs-lookup"><span data-stu-id="4d6a7-219">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="4d6a7-220">Em `OfficeRuntime.auth` :</span><span class="sxs-lookup"><span data-stu-id="4d6a7-220">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="4d6a7-221">`getAccessToken`(Windows somente)</span><span class="sxs-lookup"><span data-stu-id="4d6a7-221">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="4d6a7-222">Em `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="4d6a7-222">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="4d6a7-223">Em `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="4d6a7-223">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="4d6a7-224">Em `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="4d6a7-224">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="4d6a7-225">Em `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="4d6a7-225">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="4d6a7-226">Confira também</span><span class="sxs-lookup"><span data-stu-id="4d6a7-226">See also</span></span>

- [<span data-ttu-id="4d6a7-227">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="4d6a7-227">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="4d6a7-228">Como depurar os complementos baseados em eventos</span><span class="sxs-lookup"><span data-stu-id="4d6a7-228">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
