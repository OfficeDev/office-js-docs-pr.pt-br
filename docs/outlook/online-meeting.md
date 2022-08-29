---
title: Criar um suplemento do Outlook para um provedor de reunião online
description: Discute como configurar um suplemento do Outlook para um provedor de serviços de reunião online.
ms.topic: article
ms.date: 08/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: e1775d8cf8cc45887dfb1058603c103583d5e5dc
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2022
ms.locfileid: "67320655"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>Criar um suplemento do Outlook para um provedor de reunião online

Configurar uma reunião online é uma experiência fundamental para um usuário do Outlook e é fácil criar uma reunião do [Teams com o Outlook](/microsoftteams/teams-add-in-for-outlook). No entanto, criar uma reunião online no Outlook com um serviço que não seja da Microsoft pode ser complicado. Ao implementar esse recurso, os provedores de serviços podem simplificar a criação e a experiência de ingresso na reunião online para os usuários do suplemento do Outlook.

> [!IMPORTANT]
> Esse recurso é compatível com Outlook na Web, Windows, Mac, Android e iOS com uma assinatura do Microsoft 365.

Neste artigo, você aprenderá a configurar seu suplemento do Outlook para permitir que os usuários organizem e ingressem em uma reunião usando seu serviço de reunião online. Ao longo deste artigo, usaremos um provedor de serviços de reunião online fictício, "Contoso".

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , que cria um projeto de suplemento com o gerador Yeoman para suplementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para permitir que os usuários criem reuniões online com seu suplemento, você deve configurar o **\<VersionOverrides\>** nó no manifesto. Se você estiver criando um suplemento que só terá suporte no Outlook na Web, Windows e Mac, selecione a guia **Windows, Mac e Web** para obter diretrizes. No entanto, se o suplemento também tiver suporte no Outlook no Android e no iOS, selecione **a guia** Móvel.

# <a name="windows-mac-web"></a>[Windows, Mac, Web](#tab/non-mobile)

1. No editor de código, abra o projeto de início rápido do Outlook que você criou.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.

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

# <a name="mobile"></a>[Dispositivo móvel](#tab/mobile)

Para permitir que os usuários criem uma reunião online a partir de seu dispositivo móvel, o ponto de extensão [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) é configurado no manifesto sob o elemento pai **\<MobileFormFactor\>**. Esse ponto de extensão não tem suporte em outros fatores forma.

1. No editor de código, abra o projeto de início rápido do Outlook que você criou.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.

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

---

> [!TIP]
> Para saber mais sobre manifestos para suplementos do Outlook, consulte [manifestos de suplemento do Outlook](manifests.md) e Adicionar suporte para comandos de suplemento [para Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Implementar a adição de detalhes da reunião online

Nesta seção, saiba como o script de suplemento pode atualizar a reunião de um usuário para incluir detalhes da reunião online. A seguir, aplica-se a todas as plataformas com suporte.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

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

    let mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define and register a function command named `insertContosoMeeting` (referenced in the manifest)
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
    // Register the function.
    Office.actions.associate("insertContosoMeeting", insertContosoMeeting);

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
    ```

## <a name="testing-and-validation"></a>Teste e validação

Siga as diretrizes usuais para [testar e validar](testing-and-tips.md) seu suplemento e, em seguida, fazer [sideload](sideload-outlook-add-ins-for-testing.md) do manifesto no Outlook na Web, Windows ou Mac. Se o suplemento também der suporte a dispositivos móveis, reinicie o Outlook em seu dispositivo Android ou iOS após o sideload. Depois que o suplemento for sideload, crie uma nova reunião e verifique se a alternância do Microsoft Teams ou skype foi substituída pela sua.

### <a name="create-meeting-ui"></a>Criar interface do usuário da reunião

Como organizador da reunião, você deverá ver telas semelhantes às três imagens a seguir ao criar uma reunião.

[![A tela criar reunião no Android com a opção Contoso desativada.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![A tela criar reunião no Android com uma alternância de carregamento da Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![A tela criar reunião no Android com a opção Contoso ativada.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Ingressar na interface do usuário da reunião

Como participante da reunião, você deverá ver uma tela semelhante à imagem a seguir ao exibir a reunião.

[![A tela ingressar na reunião no Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> O **botão** Ingressar só tem suporte no Outlook na Web, Mac, Android e iOS. Se você vir apenas um link de reunião, mas não vir o botão  Ingressar em um cliente com suporte, pode ser que o modelo de reunião online para seu serviço não esteja registrado em nossos servidores. Consulte a [seção Registrar seu modelo de reunião online](#register-your-online-meeting-template) para obter detalhes.

## <a name="register-your-online-meeting-template"></a>Registrar seu modelo de reunião online

Registrar seu suplemento de reunião online é opcional. Ela só se aplica se você quiser exibir o botão **Ingressar** em reuniões, além do link da reunião. Depois de desenvolver seu suplemento de reunião online e quiser registrá-lo, crie um problema do GitHub usando as diretrizes a seguir. Entraremos em contato com você para coordenar uma linha do tempo de registro.

> [!IMPORTANT]
> O **botão** Ingressar só tem suporte no Outlook na Web, Mac, Android e iOS.

1. Crie um [novo problema do GitHub](https://github.com/OfficeDev/office-js/issues/new).
1. Defina **o Título** do novo problema como "Outlook: Registrar o modelo de reunião online para meu serviço", substituindo `my-service` pelo nome do serviço.
1. No corpo do problema, substitua `newBody` o texto existente pela cadeia de caracteres definida na variável ou semelhante na seção Implementar a adição de detalhes da reunião [online](#implement-adding-online-meeting-details) anteriormente neste artigo.
1. Clique **em Enviar novo problema**.

![Uma nova tela de problema do GitHub com o conteúdo de exemplo da Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>APIs disponíveis

As APIs a seguir estão disponíveis para esse recurso.

- APIs do Organizador de Compromissos
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Hora](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([Local](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Hora](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Assunto](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Manipular fluxo de autenticação
  - [APIs de caixa de diálogo](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Restrições

Várias restrições se aplicam.

- Aplicável somente a provedores de serviços de reunião online.
- Somente suplementos instalados pelo administrador aparecerão na tela de composição da reunião, substituindo a opção padrão do Teams ou do Skype. Os suplementos instalados pelo usuário não serão ativados.
- O ícone de suplemento deve estar em escala de cinza usando código hexadecimal `#919191` ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).
- Há suporte para apenas um comando de função no modo Organizador de Compromissos (redigir).
- O suplemento deve atualizar os detalhes da reunião no formulário de compromisso dentro do período de tempo limite de um minuto. No entanto, qualquer tempo gasto em uma caixa de diálogo que o suplemento abriu para autenticação, por exemplo, é excluído do período de tempo limite.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook Mobile](outlook-mobile-addins.md)
- [Adicionar suporte para comandos de suplementos para Outlook Mobile](add-mobile-support.md)
