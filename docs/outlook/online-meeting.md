---
title: Criar um suplemento do Outlook para um provedor de reunião online
description: Discute como configurar um suplemento do Outlook para um provedor de serviços de reunião online.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c2cdb9f6369fd851a13fe45df132482b0ccdc0e
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767179"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>Criar um suplemento do Outlook para um provedor de reunião online

Configurar uma reunião online é uma experiência fundamental para um usuário do Outlook e é fácil [criar uma reunião do Teams com o Outlook](/microsoftteams/teams-add-in-for-outlook). No entanto, criar uma reunião online no Outlook com um serviço que não seja da Microsoft pode ser complicado. Ao implementar esse recurso, os provedores de serviços podem simplificar a criação de reuniões online e a experiência de junção para os usuários do suplemento do Outlook.

> [!IMPORTANT]
> Esse recurso tem suporte em Outlook na Web, Windows, Mac, Android e iOS com uma assinatura do Microsoft 365.

Neste artigo, você aprenderá a configurar seu suplemento do Outlook para permitir que os usuários organizem e participem de uma reunião usando seu serviço de reunião online. Ao longo deste artigo, usaremos um provedor de serviços de reunião online fictício, "Contoso".

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de suplemento com o gerador Yeoman para Suplementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para permitir que os usuários criem reuniões online com seu suplemento, você deve configurar o manifesto. A marcação é diferente dependendo de duas variáveis:

- O tipo de plataforma de destino; móvel ou não móvel.
- O tipo de manifesto; [XML ou o manifesto do Teams para suplementos do Office (versão prévia)](../develop/json-manifest-overview.md).

Se o suplemento usar um manifesto XML e o suplemento só tiver suporte em Outlook na Web, Windows e Mac, selecione a guia **Windows, Mac e Web** para obter diretrizes. No entanto, se o suplemento também tiver suporte no Outlook no Android e no iOS, selecione a guia **Móvel** .

Se o suplemento usar o manifesto do Teams (versão prévia), selecione a guia **Manifesto do Teams (versão prévia do desenvolvedor).**

> [!IMPORTANT]
> Os provedores de reunião online ainda não têm suporte para o manifesto do Teams (versão prévia). Estamos trabalhando para fornecer esse suporte em breve.

# <a name="windows-mac-web"></a>[Windows, Mac, Web](#tab/non-mobile)

1. No editor de código, abra o projeto de início rápido do Outlook que você criou.

1. Abra o arquivo **manifest.xml** localizado na raiz do projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas abertas e fechadas) e substitua-o pelo XML a seguir.

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

Para permitir que os usuários criem uma reunião online a partir de seu dispositivo móvel, o [ponto de extensão MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) está configurado no manifesto no elemento **\<MobileFormFactor\>** pai . Esse ponto de extensão não tem suporte em outros fatores de formulário.

1. No editor de código, abra o projeto de início rápido do Outlook que você criou.

1. Abra o arquivo **manifest.xml** localizado na raiz do projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas abertas e fechadas) e substitua-o pelo XML a seguir.

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

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

> [!IMPORTANT]
> Os provedores de reunião online ainda não têm suporte para o [manifesto do Teams para Suplementos do Office (versão prévia)](../develop/json-manifest-overview.md). Essa guia é para uso futuro.

1. Abra o arquivo **manifest.json** .

1. Localize o *primeiro* objeto na matriz "authorization.permissions.resourceSpecific" e defina sua propriedade "name" como "MailboxItem.ReadWrite.User". Deve ser assim quando terminar.

    ```json
    {
        "name": "MailboxItem.ReadWrite.User",
        "type": "Delegated"
    }
    ```

1. Na matriz "validDomains", altere a URL para "https://contoso.com", que é a URL do provedor de reunião online fictício. A matriz deve ser assim quando você terminar.

    ```json
    "validDomains": [
        "https://contoso.com"
    ],
    ```

1. Adicione o objeto a seguir à matriz "extensions.runtimes". Observe o seguinte sobre este código.

   - A "minVersion" do conjunto de requisitos da caixa de correio está definida como "1.3" para que o runtime não seja iniciado em plataformas e versões do Office em que esse recurso não tenha suporte.
   - A "id" do runtime é definida como o nome descritivo "online_meeting_runtime".
   - A propriedade "code.page" é definida como a URL do arquivo HTML sem interface do usuário que carregará o comando da função.
   - A propriedade "lifetime" é definida como "curta", o que significa que o runtime é iniciado quando o botão de comando da função é selecionado e é desligado quando a função é concluída. (Em certos casos raros, o runtime é desligado antes da conclusão do manipulador. Consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).)
   - Há uma ação para executar uma função chamada "insertContosoMeeting". Você criará essa função em uma etapa posterior.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "online_meeting_runtime",
        "type": "general",
        "code": {
            "page": "https://contoso.com/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "insertContosoMeeting",
                "type": "executeFunction",
                "displayName": "insertContosoMeeting"
            }
        ]
    }
    ```

1. Substitua a matriz "extensions.ribbons" pelo seguinte. Observe o seguinte sobre esta marcação.

   - A "minVersion" do conjunto de requisitos da caixa de correio é definida como "1.3" para que as personalizações de faixa de opções não apareçam em plataformas e versões do Office em que esse recurso não tem suporte.
   - A matriz "contextos" especifica que a faixa de opções está disponível apenas na janela organizador de detalhes da reunião.
   - Haverá um grupo de controle personalizado na guia faixa de opções padrão (da janela organizador de detalhes da reunião) rotulada de **reunião contoso**.
   - O grupo terá um botão rotulado **Adicionar uma reunião contoso**.
   - O "actionId" do botão foi definido como "insertContosoMeeting", que corresponde à "id" da ação criada na etapa anterior.

    ```json
    "ribbons": [
      {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "scopes": [
                "mail"
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "contexts": [
            "meetingDetailsOrganizer"
        ],
        "tabs": [
            {
                "builtInTabId": "TabDefault",
                "groups": [
                    {
                        "id": "apptComposeGroup",
                        "label": "Contoso meeting",
                        "controls": [
                            {
                                "id": "insertMeetingButton",
                                "type": "button",
                                "label": "Add a Contoso meeting",
                                "icons": [
                                    {
                                        "size": 16,
                                        "file": "icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "file": "icon-32.png"
                                    },
                                    {
                                        "size": 64,
                                        "file": "icon-64_02.png"
                                    },
                                    {
                                        "size": 80,
                                        "file": "icon-80.png"
                                    }
                                ],
                                "supertip": {
                                    "title": "Add a Contoso meeting",
                                    "description": "Add a Contoso meeting to this appointment."
                                },
                                "actionId": "insertContosoMeeting",
                            }
                        ]
                    }
                ]
            }
        ]
      }
    ]
    ```

---

> [!TIP]
> Para saber mais sobre manifestos para suplementos do Outlook, confira [Manifestos de suplementos do Outlook](manifests.md) e [Adicionar suporte para comandos de suplemento para o Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Implementar a adição de detalhes da reunião online

Nesta seção, saiba como seu script de suplemento pode atualizar a reunião de um usuário para incluir detalhes da reunião online. O seguinte se aplica a todas as plataformas com suporte.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

1. Substitua todo o conteúdo do **arquivocommands.js** pelo JavaScript a seguir.

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

Siga as diretrizes habituais para [testar e validar o suplemento](testing-and-tips.md) e, em seguida, [carregar](sideload-outlook-add-ins-for-testing.md) o manifesto em Outlook na Web, Windows ou Mac. Se o suplemento também for compatível com o celular, reinicie o Outlook em seu dispositivo Android ou iOS após o sideload. Depois que o suplemento estiver sideload, crie uma nova reunião e verifique se o alternância do Microsoft Teams ou do Skype é substituído pelo seu.

### <a name="create-meeting-ui"></a>Criar interface do usuário da reunião

Como organizador da reunião, você deve ver telas semelhantes às três imagens a seguir ao criar uma reunião.

[![A tela criar reunião no Android com o alternância contoso.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![A tela criar reunião no Android com um alternância contoso de carregamento.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![A tela criar reunião no Android com o alternância Contoso.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Ingressar na interface do usuário da reunião

Como participante da reunião, você deve ver uma tela semelhante à imagem a seguir ao exibir a reunião.

[![A tela de reunião de junção no Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> O botão **Ingressar** só tem suporte em Outlook na Web, Mac, Android e iOS. Se você vir apenas um link de reunião, mas não vir o botão **Ingressar** em um cliente com suporte, pode ser que o modelo de reunião online para seu serviço não esteja registrado em nossos servidores. Consulte a seção [Registrar seu modelo de reunião online](#register-your-online-meeting-template) para obter detalhes.

## <a name="register-your-online-meeting-template"></a>Registrar seu modelo de reunião online

Registrar seu suplemento de reunião online é opcional. Ela só se aplica se você quiser exibir o botão **Ingressar** em reuniões, além do link de reunião. Depois de desenvolver seu suplemento de reunião online e desejar registrá-lo, crie um problema do GitHub usando as diretrizes a seguir. Entraremos em contato com você para coordenar uma linha do tempo de registro.

> [!IMPORTANT]
> O botão **Ingressar** só tem suporte em Outlook na Web, Mac, Android e iOS.

1. Crie um [novo problema do GitHub](https://github.com/OfficeDev/office-js/issues/new).
1. Defina o **título** do novo problema como "Outlook: registre o modelo de reunião online para meu serviço", substituindo `my-service` pelo nome do serviço.
1. No corpo do problema, substitua o texto existente pela cadeia de caracteres definida na `newBody` variável ou semelhante da seção [Implementar a adição de detalhes da reunião online](#implement-adding-online-meeting-details) anteriormente neste artigo.
1. Clique **em Enviar novo problema**.

![Uma nova tela de problemas do GitHub com o conteúdo de exemplo da Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>APIs disponíveis

As APIs a seguir estão disponíveis para esse recurso.

- APIs organizadoras de compromissos
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Hora](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([Local](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Hora](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Assunto](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Manipular fluxo de auth
  - [APIs de caixa de diálogo](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Restrições

Várias restrições se aplicam.

- Aplicável somente a provedores de serviços de reunião online.
- Somente os suplementos instalados pelo administrador serão exibidos na tela de composição da reunião, substituindo a opção padrão teams ou Skype. Os suplementos instalados pelo usuário não serão ativados.
- O ícone de suplemento deve estar em escala de cinza usando o código `#919191` hex ou o equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).
- Há suporte para apenas um comando de função no modo Organizador de Compromissos (compose).
- O suplemento deve atualizar os detalhes da reunião no formulário de compromisso dentro do período limite de um minuto. No entanto, qualquer tempo gasto em uma caixa de diálogo, o suplemento aberto para autenticação, por exemplo, é excluído do período limite.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook Mobile](outlook-mobile-addins.md)
- [Adicionar suporte para comandos de suplementos para Outlook Mobile](add-mobile-support.md)
