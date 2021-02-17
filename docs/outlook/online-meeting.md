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
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>Criar um complemento móvel do Outlook para um provedor de reunião online

Configurar uma reunião online é uma experiência fundamental para um usuário do Outlook e é fácil criar uma reunião do [Teams com o Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile. No entanto, criar uma reunião online no Outlook com um serviço que não seja da Microsoft pode ser complicado. Implementando esse recurso, os provedores de serviços podem simplificar a experiência de criação de reuniões online para seus usuários de complementos do Outlook.

> [!IMPORTANT]
> Esse recurso só é suportado no Android e iOS com uma assinatura do Microsoft 365.

Neste artigo, você aprenderá a configurar seu complemento móvel do Outlook para permitir que os usuários organizem e participem de uma reunião usando seu serviço de reunião online. Neste artigo, vamos usar um provedor de serviços de reunião online fictício, "Contoso".

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para permitir que os usuários criem reuniões online com seu complemento, você deve configurar o ponto de extensão [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) no manifesto sob o elemento `MobileFormFactor` pai. Não há suporte para outros fatores forma.

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do projeto.

1. Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.

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
> Para saber mais sobre manifestos para os complementos do [Outlook,](manifests.md) confira manifestos de complementos do Outlook e adicione suporte para comandos de complemento para [o Outlook Mobile.](add-mobile-support.md)

## <a name="implement-adding-online-meeting-details"></a>Implementar a adição de detalhes da reunião online

Nesta seção, saiba como seu script de complemento pode atualizar a reunião de um usuário para incluir detalhes da reunião online.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** editor de código.

1. Substitua todo o conteúdo do arquivo **commands.js** pelo JavaScript a seguir.

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

## <a name="testing-and-validation"></a>Teste e validação

Siga as orientações usuais [para testar e validar seu complemento.](testing-and-tips.md) Depois [de fazer sideload](sideload-outlook-add-ins-for-testing.md) no Outlook na Web, no Windows ou no Mac, reinicie o Outlook em seu dispositivo móvel Android. (O Android é o único cliente com suporte por enquanto.) Em seguida, em uma nova tela de reunião, verifique se a alternância do Microsoft Teams ou do Skype foi substituída pelo seu próprio.

### <a name="create-meeting-ui"></a>Criar interface do usuário de reunião

Como organizador da reunião, você deverá ver telas semelhantes às três imagens a seguir ao criar uma reunião.

[ ![ screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - loading Contoso toggle screenshot](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![ of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Ingressar na interface do usuário da reunião

Como participante de uma reunião, você deverá ver uma tela semelhante à imagem a seguir ao exibir a reunião.

[![captura de tela da tela participar da reunião no Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Se você não vir o **link** Ingressar, pode ser que o modelo de reunião online para seu serviço não esteja registrado em nossos servidores. Consulte a [seção Registrar seu modelo de reunião online](#register-your-online-meeting-template) para obter detalhes.

## <a name="register-your-online-meeting-template"></a>Registrar seu modelo de reunião online

Se você quiser registrar o modelo de reunião online para seu serviço, poderá criar um problema do GitHub com os detalhes. Depois disso, entraremos em contato com você para coordenar a linha do tempo do registro.

1. Vá para a **seção** Comentários no final deste artigo.
1. Pressione o link **Esta** página.
1. De definir **o Título** do novo problema como "Registrar o modelo de reunião online para meu serviço", substituindo pelo nome `my-service` do serviço.
1. No corpo do problema, substitua a cadeia de caracteres "[Insira comentários aqui]" pela cadeia de caracteres que você definiu na variável ou semelhante na seção Implementar adicionando detalhes da reunião online anteriormente `newBody` neste artigo. [](#implement-adding-online-meeting-details)
1. Clique **em Enviar novo problema.**

![captura de tela da nova tela de problema do GitHub com o conteúdo de exemplo da Contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>APIs disponíveis

As seguintes APIs estão disponíveis para esse recurso.

- APIs do Organizador de Compromissos
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Manipular fluxo de auth
  - [APIs de caixa de diálogo](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Restrições

Várias restrições se aplicam.

- Aplicável somente a provedores de serviços de reunião online.
- Somente os complementos instalados pelo administrador aparecerão na tela de redação da reunião, substituindo a opção padrão do Teams ou do Skype. Os complementos instalados pelo usuário não serão ativados.
- O ícone de add-in deve estar em escala de cinza usando código hexaxa ou `#919191` seu equivalente em outros [formatos de cor.](https://convertingcolors.com/hex-color-919191.html)
- Somente um comando sem interface do usuário é suportado no modo Organizador de Compromissos (redação).

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook Mobile](outlook-mobile-addins.md)
- [Adicionar suporte para comandos de suplementos para Outlook Mobile](add-mobile-support.md)
