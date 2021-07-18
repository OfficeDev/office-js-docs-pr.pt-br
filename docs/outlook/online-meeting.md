---
title: Criar um Outlook de dispositivo móvel para um provedor de reunião online
description: Discute como configurar um Outlook de celular para um provedor de serviços de reunião online.
ms.topic: article
ms.date: 07/09/2021
localization_priority: Normal
ms.openlocfilehash: f0f9b69c2b8b515df3829ca3ba0714393df79fd1
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455499"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="8d755-103">Criar um Outlook de dispositivo móvel para um provedor de reunião online</span><span class="sxs-lookup"><span data-stu-id="8d755-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="8d755-104">Configurar uma reunião online é uma experiência fundamental para um usuário Outlook, e é fácil criar uma reunião Teams [com](/microsoftteams/teams-add-in-for-outlook) Outlook celular.</span><span class="sxs-lookup"><span data-stu-id="8d755-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="8d755-105">No entanto, criar uma reunião online Outlook com um serviço que não seja da Microsoft pode ser complicado.</span><span class="sxs-lookup"><span data-stu-id="8d755-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="8d755-106">Ao implementar esse recurso, os provedores de serviços podem simplificar a experiência de criação de reuniões online para seus usuários Outlook de complemento.</span><span class="sxs-lookup"><span data-stu-id="8d755-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8d755-107">Esse recurso só tem suporte para Android e iOS com uma assinatura Microsoft 365 assinatura.</span><span class="sxs-lookup"><span data-stu-id="8d755-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="8d755-108">Neste artigo, você aprenderá a configurar seu Outlook de dispositivo móvel para permitir que os usuários organizem e participem de uma reunião usando seu serviço de reunião online.</span><span class="sxs-lookup"><span data-stu-id="8d755-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="8d755-109">Ao longo deste artigo, vamos usar um provedor de serviços de reunião online fictício, "Contoso".</span><span class="sxs-lookup"><span data-stu-id="8d755-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="8d755-110">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="8d755-110">Set up your environment</span></span>

<span data-ttu-id="8d755-111">Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido que cria um projeto de complemento com o gerador Yeoman para Office Desempois.</span><span class="sxs-lookup"><span data-stu-id="8d755-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="8d755-112">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="8d755-112">Configure the manifest</span></span>

<span data-ttu-id="8d755-113">Para permitir que os usuários criem reuniões online com seu complemento, você deve configurar o ponto de extensão [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) no manifesto sob o elemento pai `MobileFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="8d755-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="8d755-114">Outros fatores de formulário não são suportados.</span><span class="sxs-lookup"><span data-stu-id="8d755-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="8d755-115">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="8d755-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="8d755-116">Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="8d755-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="8d755-117">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas abertas e próximas) e substitua-o pelo XML a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d755-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="8d755-118">Para saber mais sobre manifestos para Outlook de Outlook, consulte [Outlook manifestos](manifests.md) de complemento e Adicionar suporte para comandos de complemento para [Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="8d755-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="8d755-119">Implementar a adição de detalhes da reunião online</span><span class="sxs-lookup"><span data-stu-id="8d755-119">Implement adding online meeting details</span></span>

<span data-ttu-id="8d755-120">Nesta seção, saiba como seu script de complemento pode atualizar a reunião de um usuário para incluir detalhes da reunião online.</span><span class="sxs-lookup"><span data-stu-id="8d755-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="8d755-121">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.</span><span class="sxs-lookup"><span data-stu-id="8d755-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="8d755-122">Substitua todo o conteúdo do arquivo **commands.js** pelo JavaScript a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d755-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a><span data-ttu-id="8d755-123">Teste e validação</span><span class="sxs-lookup"><span data-stu-id="8d755-123">Testing and validation</span></span>

<span data-ttu-id="8d755-124">Siga as diretrizes usuais para [testar e validar seu complemento](testing-and-tips.md).</span><span class="sxs-lookup"><span data-stu-id="8d755-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="8d755-125">Depois [de fazer sideload](sideload-outlook-add-ins-for-testing.md) em Outlook na Web, Windows ou Mac, reinicie Outlook em seu dispositivo móvel Android ou iOS.</span><span class="sxs-lookup"><span data-stu-id="8d755-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android or iOS mobile device.</span></span> <span data-ttu-id="8d755-126">Em seguida, em uma nova tela de reunião, verifique se a Microsoft Teams ou Skype alternância foi substituída por sua própria.</span><span class="sxs-lookup"><span data-stu-id="8d755-126">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="8d755-127">Criar interface do usuário de reunião</span><span class="sxs-lookup"><span data-stu-id="8d755-127">Create meeting UI</span></span>

<span data-ttu-id="8d755-128">Como organizador da reunião, você deve ver telas semelhantes às três imagens a seguir ao criar uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8d755-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="8d755-129">[![A tela criar reunião no Android - Contoso alterna.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="8d755-129">[![The create meeting screen on Android - Contoso toggle off.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span></span> <span data-ttu-id="8d755-130">[![A tela criar reunião no Android - carregando a alternância Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="8d755-130">[![The create meeting screen on Android - loading Contoso toggle.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span></span> <span data-ttu-id="8d755-131">[![A tela criar reunião no Android - Contoso é alternada.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="8d755-131">[![The create meeting screen on Android - Contoso toggle on.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="8d755-132">Ingressar na interface do usuário de reunião</span><span class="sxs-lookup"><span data-stu-id="8d755-132">Join meeting UI</span></span>

<span data-ttu-id="8d755-133">Como participante de reunião, você deve ver uma tela semelhante à imagem a seguir ao exibir a reunião.</span><span class="sxs-lookup"><span data-stu-id="8d755-133">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="8d755-134">[![Captura de tela da tela de reunião de junção no Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="8d755-134">[![Screenshot of join meeting screen on Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8d755-135">Se você não vir o **link** Ingressar, pode ser que o modelo de reunião online para seu serviço não esteja registrado em nossos servidores.</span><span class="sxs-lookup"><span data-stu-id="8d755-135">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="8d755-136">Consulte a [seção Registrar seu modelo de reunião online](#register-your-online-meeting-template) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="8d755-136">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="8d755-137">Registrar seu modelo de reunião online</span><span class="sxs-lookup"><span data-stu-id="8d755-137">Register your online-meeting template</span></span>

<span data-ttu-id="8d755-138">Se você quiser registrar o modelo de reunião online para seu serviço, você pode criar um problema GitHub com os detalhes.</span><span class="sxs-lookup"><span data-stu-id="8d755-138">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="8d755-139">Depois disso, entraremos em contato com você para coordenar a linha do tempo do registro.</span><span class="sxs-lookup"><span data-stu-id="8d755-139">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="8d755-140">Vá para a **seção Comentários** no final deste artigo.</span><span class="sxs-lookup"><span data-stu-id="8d755-140">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="8d755-141">Pressione o link **Esta** página.</span><span class="sxs-lookup"><span data-stu-id="8d755-141">Press the **This page** link.</span></span>
1. <span data-ttu-id="8d755-142">De definir **o Título** do novo problema como "Registrar o modelo de reunião online para meu serviço", substituindo pelo nome `my-service` do serviço.</span><span class="sxs-lookup"><span data-stu-id="8d755-142">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="8d755-143">No corpo do problema, substitua a cadeia de caracteres "[Insira comentários aqui]" pela cadeia de caracteres definida na variável ou semelhante na seção Implementar a adição de detalhes da reunião online anteriormente `newBody` neste artigo. [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="8d755-143">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="8d755-144">Clique **em Enviar novo problema**.</span><span class="sxs-lookup"><span data-stu-id="8d755-144">Click **Submit new issue**.</span></span>

![screenshot of new GitHub issue screen with Contoso sample content.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="8d755-146">APIs disponíveis</span><span class="sxs-lookup"><span data-stu-id="8d755-146">Available APIs</span></span>

<span data-ttu-id="8d755-147">As APIs a seguir estão disponíveis para esse recurso.</span><span class="sxs-lookup"><span data-stu-id="8d755-147">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="8d755-148">APIs organizadores de compromissos</span><span class="sxs-lookup"><span data-stu-id="8d755-148">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="8d755-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="8d755-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="8d755-150">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-150">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-152">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-152">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-155">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Hora](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-155">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-156">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Assunto](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-156">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="8d755-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="8d755-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="8d755-158">Manipular fluxo de auth</span><span class="sxs-lookup"><span data-stu-id="8d755-158">Handle auth flow</span></span>
  - [<span data-ttu-id="8d755-159">APIs de caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="8d755-159">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="8d755-160">Restrições</span><span class="sxs-lookup"><span data-stu-id="8d755-160">Restrictions</span></span>

<span data-ttu-id="8d755-161">Várias restrições se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d755-161">Several restrictions apply.</span></span>

- <span data-ttu-id="8d755-162">Aplicável somente a provedores de serviços de reunião online.</span><span class="sxs-lookup"><span data-stu-id="8d755-162">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="8d755-163">Somente os complementos instalados pelo administrador aparecerão na tela de redação da reunião, substituindo a opção padrão Teams ou Skype de reunião.</span><span class="sxs-lookup"><span data-stu-id="8d755-163">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="8d755-164">Os complementos instalados pelo usuário não serão ativados.</span><span class="sxs-lookup"><span data-stu-id="8d755-164">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="8d755-165">O ícone do add-in deve estar em escala de cinza usando código hexaxa ou seu equivalente `#919191` em [outros formatos de cor.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="8d755-165">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="8d755-166">Somente um comando sem interface do usuário é suportado no modo Organizador de Compromissos (redação).</span><span class="sxs-lookup"><span data-stu-id="8d755-166">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>
- <span data-ttu-id="8d755-167">O complemento deve atualizar os detalhes da reunião no formulário de compromisso dentro do período de tempo de um minuto.</span><span class="sxs-lookup"><span data-stu-id="8d755-167">The add-in should update the meeting details in the appointment form within the one-minute timeout period.</span></span> <span data-ttu-id="8d755-168">No entanto, qualquer tempo gasto em uma caixa de diálogo que o add-in abriu para autenticação, etc. é excluído do período de tempo de tempo.</span><span class="sxs-lookup"><span data-stu-id="8d755-168">However, any time spent in a dialog box the add-in opened for authentication, etc. is excluded from the timeout period.</span></span>

## <a name="see-also"></a><span data-ttu-id="8d755-169">Confira também</span><span class="sxs-lookup"><span data-stu-id="8d755-169">See also</span></span>

- [<span data-ttu-id="8d755-170">Suplementos do Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="8d755-170">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="8d755-171">Adicionar suporte para comandos de suplementos para Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="8d755-171">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
