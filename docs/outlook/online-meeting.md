---
title: Criar um suplemento do Outlook Mobile para um provedor de reunião online
description: Descreve como configurar um suplemento móvel do Outlook para um provedor de serviços de reunião online.
ms.topic: article
ms.date: 05/19/2020
localization_priority: Normal
ms.openlocfilehash: 1d42ec82e12e9f34f0211ca9926f5ae8b92c7804
ms.sourcegitcommit: 8499a4247d1cb1e96e99c17cb520f4a8a41667e3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2020
ms.locfileid: "44292284"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="4fb9b-103">Criar um suplemento do Outlook Mobile para um provedor de reunião online</span><span class="sxs-lookup"><span data-stu-id="4fb9b-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="4fb9b-104">A configuração de uma reunião online é uma experiência principal para um usuário do Outlook e é fácil [criar uma reunião do teams com o Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="4fb9b-105">No entanto, a criação de uma reunião online no Outlook com um serviço que não seja da Microsoft pode ser complicada.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="4fb9b-106">Ao implementar esse recurso, os provedores de serviços podem simplificar a experiência de criação de reunião online para os usuários de suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4fb9b-107">Este recurso só é suportado no Android com uma assinatura do Office 365.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-107">This feature is only supported on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="4fb9b-108">Neste artigo, você aprenderá como configurar seu suplemento do Outlook Mobile para permitir que os usuários organizem e ingressem em uma reunião usando o serviço de reunião online.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="4fb9b-109">Neste artigo, vamos usar um provedor de serviço de reunião online fictício, "contoso".</span><span class="sxs-lookup"><span data-stu-id="4fb9b-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="4fb9b-110">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="4fb9b-110">Configure the manifest</span></span>

<span data-ttu-id="4fb9b-111">Para permitir que os usuários criem reuniões online com seu suplemento, você deve configurar o `MobileOnlineMeetingCommandSurface` ponto de extensão no manifesto no elemento pai `MobileFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="4fb9b-111">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="4fb9b-112">Não há suporte para outros fatores de formulário.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-112">Other form factors are not supported.</span></span>

<span data-ttu-id="4fb9b-113">O exemplo a seguir mostra um trecho do manifesto que inclui o `MobileFormFactor` elemento e o `MobileOnlineMeetingCommandSurface` ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-113">The following example shows an excerpt from the manifest that includes the `MobileFormFactor` element and `MobileOnlineMeetingCommandSurface` extension point.</span></span>

> [!TIP]
> <span data-ttu-id="4fb9b-114">Para saber mais sobre manifestos para suplementos do Outlook, confira [manifestos de suplemento do Outlook](manifests.md) e [Adicione suporte para comandos de suplemento do Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="4fb9b-114">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="3" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="4fb9b-115">Implementar adicionando detalhes da reunião online</span><span class="sxs-lookup"><span data-stu-id="4fb9b-115">Implement adding online meeting details</span></span>

<span data-ttu-id="4fb9b-116">Nesta seção, saiba como o script do seu suplemento pode atualizar a reunião de um usuário para incluir detalhes online da reunião.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-116">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

<span data-ttu-id="4fb9b-117">O exemplo a seguir mostra como você cria detalhes da reunião online.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-117">The following example shows how you construct online meeting details.</span></span> <span data-ttu-id="4fb9b-118">Não mostrado é como obter a ID do organizador da reunião e outros detalhes do serviço.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-118">Not shown is how to get the meeting organizer's ID and other details from your service.</span></span>

```js
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
```

<span data-ttu-id="4fb9b-119">O exemplo a seguir mostra como definir uma função sem interface do usuário chamada `insertContosoMeeting` referenciada no manifesto para atualizar o corpo da reunião com os detalhes da reunião online.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-119">The following example shows how to define a UI-less function named `insertContosoMeeting` referenced in the manifest to update the meeting body with the online meeting details.</span></span>

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

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
```

<span data-ttu-id="4fb9b-120">O exemplo a seguir mostra uma implementação da função de suporte `updateBody` usada no exemplo anterior que acrescenta os detalhes da reunião online ao corpo atual da reunião.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-120">The following example shows an implementation of the supporting function `updateBody` used in the previous example that appends the online meeting details to the current body of the meeting.</span></span>

```js
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
```

## <a name="testing-and-validation"></a><span data-ttu-id="4fb9b-121">Teste e validação</span><span class="sxs-lookup"><span data-stu-id="4fb9b-121">Testing and validation</span></span>

<span data-ttu-id="4fb9b-122">Siga as orientações usuais para [testar e validar o suplemento](testing-and-tips.md).</span><span class="sxs-lookup"><span data-stu-id="4fb9b-122">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="4fb9b-123">Após o [Sideload](sideload-outlook-add-ins-for-testing.md) no Outlook na Web, no Windows ou no Mac, reinicie o Outlook no seu dispositivo móvel Android (o Android é o único cliente com suporte para agora).</span><span class="sxs-lookup"><span data-stu-id="4fb9b-123">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device (Android is the only supported client for now).</span></span> <span data-ttu-id="4fb9b-124">Em seguida, em uma nova tela de reunião, verifique se o Microsoft Teams ou o alternância do Skype foi substituído por seu próprio.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-124">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="4fb9b-125">Criar IU de reunião</span><span class="sxs-lookup"><span data-stu-id="4fb9b-125">Create meeting UI</span></span>

<span data-ttu-id="4fb9b-126">Como organizador da reunião, você deve ver telas semelhantes às três imagens a seguir ao criar uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-126">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="4fb9b-127">[ ![ captura de tela da tela criar reunião no Android-ativar/desativar](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![ captura de tela da tela criar reunião no Android-carregando a](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![ captura de tela da Contoso Toggle Screen do botão criar reunião no Android-ativar/desativar](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="4fb9b-127">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="4fb9b-128">Ingressar na IU da reunião</span><span class="sxs-lookup"><span data-stu-id="4fb9b-128">Join meeting UI</span></span>

<span data-ttu-id="4fb9b-129">Como participante da reunião, você verá uma tela semelhante à seguinte imagem ao exibir a reunião.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-129">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="4fb9b-130">[![captura de tela da tela ingressar na reunião no Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="4fb9b-130">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

## <a name="available-apis"></a><span data-ttu-id="4fb9b-131">APIs disponíveis</span><span class="sxs-lookup"><span data-stu-id="4fb9b-131">Available APIs</span></span>

<span data-ttu-id="4fb9b-132">As seguintes APIs estão disponíveis para este recurso.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-132">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="4fb9b-133">APIs do organizador de compromissos</span><span class="sxs-lookup"><span data-stu-id="4fb9b-133">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="4fb9b-134">[Office. Context. Mailbox. Item. Subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([assunto](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-134">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-135">[Office. Context. Mailbox. Item. Start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-135">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-136">[Office. Context. Mailbox. Item. end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-136">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-137">[Office. Context. Mailbox. Item. Location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([local](/javascript/api/outlook/office.location?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-137">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-138">[Office. Context. Mailbox. Item. optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-138">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-139">[Office. Context. Mailbox. Item. requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-139">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-140">[Office. Context. Mailbox. Item. Body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body. getasync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body. setasync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-140">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="4fb9b-141">[Office. Context. Mailbox. Item. loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-141">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="4fb9b-142">[Office. Context. roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="4fb9b-142">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span></span>
- <span data-ttu-id="4fb9b-143">Gerenciar fluxo de autenticação</span><span class="sxs-lookup"><span data-stu-id="4fb9b-143">Handle auth flow</span></span>
  - [<span data-ttu-id="4fb9b-144">APIs de caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="4fb9b-144">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="4fb9b-145">Restriction</span><span class="sxs-lookup"><span data-stu-id="4fb9b-145">Restrictions</span></span>

<span data-ttu-id="4fb9b-146">Várias restrições se aplicam.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-146">Several restrictions apply.</span></span>

- <span data-ttu-id="4fb9b-147">Aplicável somente aos provedores de serviço de reunião online.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-147">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="4fb9b-148">No momento, o Android é o único cliente com suporte.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-148">At present, Android is the only supported client.</span></span> <span data-ttu-id="4fb9b-149">O suporte ao iOS estará disponível em breve.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-149">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="4fb9b-150">Somente os suplementos instalados pelo administrador serão exibidos na tela de redação da reunião, substituindo a opção Teams ou Skype padrão.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-150">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="4fb9b-151">Os suplementos instalados pelo usuário não serão ativados.</span><span class="sxs-lookup"><span data-stu-id="4fb9b-151">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="4fb9b-152">O ícone do suplemento deve estar em escala de cinza usando o código hex `#919191` ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="4fb9b-152">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="4fb9b-153">Só há suporte para um comando sem interface do usuário no modo de organizador de compromisso (compor).</span><span class="sxs-lookup"><span data-stu-id="4fb9b-153">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="4fb9b-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="4fb9b-154">See also</span></span>

- [<span data-ttu-id="4fb9b-155">Suplementos do Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="4fb9b-155">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="4fb9b-156">Adicionar suporte para comandos de suplementos para Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="4fb9b-156">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
