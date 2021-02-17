---
title: Criar um complemento móvel do Outlook para um provedor de reunião online
description: Discute como configurar um complemento móvel do Outlook para um provedor de serviços de reunião online.
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: fb98ddeeef8615476659a0abb798ea7901d81248
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270739"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="fca49-103">Criar um complemento móvel do Outlook para um provedor de reunião online</span><span class="sxs-lookup"><span data-stu-id="fca49-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="fca49-104">Configurar uma reunião online é uma experiência fundamental para um usuário do Outlook e é fácil criar uma reunião do [Teams com o Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile.</span><span class="sxs-lookup"><span data-stu-id="fca49-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="fca49-105">No entanto, criar uma reunião online no Outlook com um serviço que não seja da Microsoft pode ser complicado.</span><span class="sxs-lookup"><span data-stu-id="fca49-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="fca49-106">Implementando esse recurso, os provedores de serviços podem simplificar a experiência de criação de reuniões online para seus usuários de complementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fca49-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fca49-107">Esse recurso só é suportado no Android e iOS com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="fca49-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="fca49-108">Neste artigo, você aprenderá a configurar seu complemento móvel do Outlook para permitir que os usuários organizem e participem de uma reunião usando seu serviço de reunião online.</span><span class="sxs-lookup"><span data-stu-id="fca49-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="fca49-109">Neste artigo, vamos usar um provedor de serviços de reunião online fictício, "Contoso".</span><span class="sxs-lookup"><span data-stu-id="fca49-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="fca49-110">Configurar seu ambiente</span><span class="sxs-lookup"><span data-stu-id="fca49-110">Set up your environment</span></span>

<span data-ttu-id="fca49-111">Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.</span><span class="sxs-lookup"><span data-stu-id="fca49-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="fca49-112">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="fca49-112">Configure the manifest</span></span>

<span data-ttu-id="fca49-113">Para permitir que os usuários criem reuniões online com seu complemento, você deve configurar o ponto de extensão [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) no manifesto sob o elemento `MobileFormFactor` pai.</span><span class="sxs-lookup"><span data-stu-id="fca49-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="fca49-114">Não há suporte para outros fatores forma.</span><span class="sxs-lookup"><span data-stu-id="fca49-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="fca49-115">No editor de código, abra o projeto de início rápido.</span><span class="sxs-lookup"><span data-stu-id="fca49-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="fca49-116">Abra o **manifest.xml** arquivo localizado na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="fca49-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="fca49-117">Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.</span><span class="sxs-lookup"><span data-stu-id="fca49-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="fca49-118">Para saber mais sobre manifestos para os complementos do [Outlook,](manifests.md) confira manifestos de complementos do Outlook e adicione suporte para comandos de complemento para [o Outlook Mobile.](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="fca49-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="fca49-119">Implementar a adição de detalhes da reunião online</span><span class="sxs-lookup"><span data-stu-id="fca49-119">Implement adding online meeting details</span></span>

<span data-ttu-id="fca49-120">Nesta seção, saiba como seu script de complemento pode atualizar a reunião de um usuário para incluir detalhes da reunião online.</span><span class="sxs-lookup"><span data-stu-id="fca49-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="fca49-121">No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** editor de código.</span><span class="sxs-lookup"><span data-stu-id="fca49-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="fca49-122">Substitua todo o conteúdo do arquivo **commands.js** pelo JavaScript a seguir.</span><span class="sxs-lookup"><span data-stu-id="fca49-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

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

## <a name="testing-and-validation"></a><span data-ttu-id="fca49-123">Teste e validação</span><span class="sxs-lookup"><span data-stu-id="fca49-123">Testing and validation</span></span>

<span data-ttu-id="fca49-124">Siga as orientações usuais [para testar e validar seu complemento.](testing-and-tips.md)</span><span class="sxs-lookup"><span data-stu-id="fca49-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="fca49-125">Depois [de fazer sideload](sideload-outlook-add-ins-for-testing.md) no Outlook na Web, no Windows ou no Mac, reinicie o Outlook em seu dispositivo móvel Android.</span><span class="sxs-lookup"><span data-stu-id="fca49-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.</span></span> <span data-ttu-id="fca49-126">(O Android é o único cliente com suporte por enquanto.) Em seguida, em uma nova tela de reunião, verifique se a alternância do Microsoft Teams ou do Skype foi substituída pelo seu próprio.</span><span class="sxs-lookup"><span data-stu-id="fca49-126">(Android is the only supported client for now.) Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="fca49-127">Criar interface do usuário de reunião</span><span class="sxs-lookup"><span data-stu-id="fca49-127">Create meeting UI</span></span>

<span data-ttu-id="fca49-128">Como organizador da reunião, você deverá ver telas semelhantes às três imagens a seguir ao criar uma reunião.</span><span class="sxs-lookup"><span data-stu-id="fca49-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="fca49-129">[ ![ screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - loading Contoso toggle screenshot](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![ of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="fca49-129">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="fca49-130">Ingressar na interface do usuário da reunião</span><span class="sxs-lookup"><span data-stu-id="fca49-130">Join meeting UI</span></span>

<span data-ttu-id="fca49-131">Como participante de uma reunião, você deverá ver uma tela semelhante à imagem a seguir ao exibir a reunião.</span><span class="sxs-lookup"><span data-stu-id="fca49-131">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="fca49-132">[![captura de tela da tela participar da reunião no Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="fca49-132">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fca49-133">Se você não vir o **link** Ingressar, pode ser que o modelo de reunião online para seu serviço não esteja registrado em nossos servidores.</span><span class="sxs-lookup"><span data-stu-id="fca49-133">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="fca49-134">Consulte a [seção Registrar seu modelo de reunião online](#register-your-online-meeting-template) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="fca49-134">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="fca49-135">Registrar seu modelo de reunião online</span><span class="sxs-lookup"><span data-stu-id="fca49-135">Register your online-meeting template</span></span>

<span data-ttu-id="fca49-136">Se você quiser registrar o modelo de reunião online para seu serviço, poderá criar um problema do GitHub com os detalhes.</span><span class="sxs-lookup"><span data-stu-id="fca49-136">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="fca49-137">Depois disso, entraremos em contato com você para coordenar a linha do tempo do registro.</span><span class="sxs-lookup"><span data-stu-id="fca49-137">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="fca49-138">Vá para a **seção** Comentários no final deste artigo.</span><span class="sxs-lookup"><span data-stu-id="fca49-138">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="fca49-139">Pressione o link **Esta** página.</span><span class="sxs-lookup"><span data-stu-id="fca49-139">Press the **This page** link.</span></span>
1. <span data-ttu-id="fca49-140">De definir **o Título** do novo problema como "Registrar o modelo de reunião online para meu serviço", substituindo pelo nome `my-service` do serviço.</span><span class="sxs-lookup"><span data-stu-id="fca49-140">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="fca49-141">No corpo do problema, substitua a cadeia de caracteres "[Insira comentários aqui]" pela cadeia de caracteres que você definiu na variável ou semelhante na seção Implementar adicionando detalhes da reunião online anteriormente `newBody` neste artigo. [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="fca49-141">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="fca49-142">Clique **em Enviar novo problema.**</span><span class="sxs-lookup"><span data-stu-id="fca49-142">Click **Submit new issue**.</span></span>

![captura de tela da nova tela de problema do GitHub com o conteúdo de exemplo da Contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="fca49-144">APIs disponíveis</span><span class="sxs-lookup"><span data-stu-id="fca49-144">Available APIs</span></span>

<span data-ttu-id="fca49-145">As seguintes APIs estão disponíveis para esse recurso.</span><span class="sxs-lookup"><span data-stu-id="fca49-145">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="fca49-146">APIs do Organizador de Compromissos</span><span class="sxs-lookup"><span data-stu-id="fca49-146">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="fca49-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="fca49-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="fca49-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="fca49-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="fca49-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="fca49-156">Manipular fluxo de auth</span><span class="sxs-lookup"><span data-stu-id="fca49-156">Handle auth flow</span></span>
  - [<span data-ttu-id="fca49-157">APIs de caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="fca49-157">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="fca49-158">Restrições</span><span class="sxs-lookup"><span data-stu-id="fca49-158">Restrictions</span></span>

<span data-ttu-id="fca49-159">Várias restrições se aplicam.</span><span class="sxs-lookup"><span data-stu-id="fca49-159">Several restrictions apply.</span></span>

- <span data-ttu-id="fca49-160">Aplicável somente a provedores de serviços de reunião online.</span><span class="sxs-lookup"><span data-stu-id="fca49-160">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="fca49-161">Somente os complementos instalados pelo administrador aparecerão na tela de redação da reunião, substituindo a opção padrão do Teams ou do Skype.</span><span class="sxs-lookup"><span data-stu-id="fca49-161">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="fca49-162">Os complementos instalados pelo usuário não serão ativados.</span><span class="sxs-lookup"><span data-stu-id="fca49-162">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="fca49-163">O ícone de add-in deve estar em escala de cinza usando código hexaxa ou `#919191` seu equivalente em outros [formatos de cor.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="fca49-163">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="fca49-164">Somente um comando sem interface do usuário é suportado no modo Organizador de Compromissos (redação).</span><span class="sxs-lookup"><span data-stu-id="fca49-164">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="fca49-165">Confira também</span><span class="sxs-lookup"><span data-stu-id="fca49-165">See also</span></span>

- [<span data-ttu-id="fca49-166">Suplementos do Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="fca49-166">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="fca49-167">Adicionar suporte para comandos de suplementos para Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="fca49-167">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
