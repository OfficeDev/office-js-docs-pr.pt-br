---
title: Registrar anotações de compromisso em um aplicativo externo em suplementos móveis do Outlook
description: Saiba como configurar um suplemento móvel do Outlook para registrar anotações de compromisso e outros detalhes em um aplicativo externo.
ms.topic: article
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: a980b68c603154c42112f525ec6285b740ce38a5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607580"
---
# <a name="log-appointment-notes-to-an-external-application-in-outlook-mobile-add-ins"></a>Registrar anotações de compromisso em um aplicativo externo em suplementos móveis do Outlook

Salvar suas anotações de compromisso e outros detalhes em um CRM (gerenciamento de relacionamento com o cliente) ou aplicativo de anotações pode ajudá-lo a acompanhar as reuniões que você participou.

Neste artigo, você aprenderá a configurar seu suplemento móvel do Outlook para permitir que os usuários registrem anotações e outros detalhes sobre seus compromissos com seu CRM ou aplicativo de anotações. Ao longo deste artigo, usaremos um provedor de serviços crm fictício chamado "Contoso".

> [!IMPORTANT]
> Esse recurso só tem suporte no Android com uma assinatura do Microsoft 365.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) para criar um projeto de suplemento com o gerador Yeoman para suplementos do Office.

## <a name="capture-and-view-appointment-notes"></a>Capturar e exibir anotações de compromisso

Você pode optar por implementar um comando de função ou um painel de tarefas. Para atualizar o suplemento, selecione a guia para o comando de função ou o painel de tarefas e siga as instruções.

# <a name="function-command"></a>[Comando de função](#tab/noui)

Essa opção permitirá que um usuário registre e exiba suas anotações e outros detalhes sobre seus compromissos ao selecionar um comando de função na faixa de opções.

### <a name="configure-the-manifest"></a>Configurar o manifesto

Para permitir que os usuários registrem anotações de compromisso com seu suplemento, você deve configurar o ponto de extensão [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) no manifesto sob o elemento pai `MobileFormFactor`. Não há suporte para outros fatores forma.

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó inteiro `<VersionOverrides>` (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir. Substitua todas as referências à **Contoso** com as informações da sua empresa.

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
              <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="apptReadGroup">
                    <Label resid="residDescription"/>
                    <Control xsi:type="Button" id="apptReadOpenPaneButton">
                      <Label resid="residLabel"/>
                      <Supertip>
                        <Title resid="residLabel"/>
                        <Description resid="residTooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="icon-16"/>
                        <bt:Image size="32" resid="icon-32"/>
                        <bt:Image size="80" resid="icon-80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>logCRMEvent</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
            <MobileFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <FunctionName>logCRMEvent</FunctionName>
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
            <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
            <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Para saber mais sobre manifestos para suplementos do Outlook, consulte [manifestos de suplemento do Outlook](manifests.md) e Adicionar suporte para comandos de suplemento [para Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Capturar anotações de compromisso

Nesta seção, saiba como o suplemento pode extrair detalhes do compromisso quando o usuário seleciona o **botão Log** .

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

1. Substitua todo o conteúdo do arquivo **commands.js** pelo JavaScript a seguir.

    ```js
    var event;

    Office.initialize = function (reason) {
      // Add any initialization code here.
    };

    function logCRMEvent(appointmentEvent) {
      event = appointmentEvent;
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
          } else {
            console.error("Failed to get body.");
            event.completed({ allowEvent: false });
          }
        }
      );
    }

    // Register the function.
    Office.actions.associate("logCRMEvent", logCRMEvent);
    ```

Em seguida, atualize **ocommands.html** para fazer **referênciacommands.js**.

1. No mesmo projeto de início rápido, abra o **arquivo ./src/commands/commands.html** no editor de código.

1. Localize e substitua `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` pelo seguinte:

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="commands.js"></script>
    ```

### <a name="view-appointment-notes"></a>Exibir anotações de compromisso

O **rótulo** do botão Log pode ser alternado para exibir o **Modo** de Exibição definindo a **propriedade personalizada EventLogged** reservada para essa finalidade. Quando o usuário seleciona o **botão Exibir** , ele pode examinar suas anotações registradas para esse compromisso.

Seu suplemento define a experiência de exibição de log. Por exemplo, você pode exibir as anotações de compromisso registradas em uma caixa de diálogo quando o usuário seleciona o **botão Exibir** . Para obter detalhes sobre como usar diálogos, [consulte Usar a API de diálogo do Office em seus Suplementos do Office](../develop/dialog-api-in-office-add-ins.md).

Adicione a seguinte função **a ./src/commands/commands.js**. Essa função define a **propriedade personalizada EventLogged** no item de compromisso atual.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
              event.completed({ allowEvent: true });
              event = undefined;
            }
          }
        );
      }
    }
  );
}
```

Em seguida, chame-o depois que o suplemento registra com êxito as anotações do compromisso. Por exemplo, você pode chamá-lo de **logCRMEvent** , conforme mostrado na função a seguir.

```js
function logCRMEvent(appointmentEvent) {
  event = appointmentEvent;
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Replace `event.completed({ allowEvent: true });` with the following statement.
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
        event.completed({ allowEvent: false });
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Excluir o log de compromissos

Se você quiser permitir que os usuários desfaçam o registro em log ou excluam as anotações de compromisso registradas para que um log de substituição possa ser salvo, você tem duas opções.

1. Use o Microsoft Graph para [limpar o objeto de propriedades personalizadas](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) quando o usuário selecionar o botão apropriado na faixa de opções.
1. Adicione a seguinte função **a ./src/commands/commands.js** para limpar a propriedade personalizada **EventLogged** no item de compromisso atual.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                  event.completed({ allowEvent: true });
                  event = undefined;
                }
              }
            );
          }
        }
      );
    }
    ```

Em seguida, chame-o quando quiser limpar a propriedade personalizada. Por exemplo, você pode chamá-lo do **logCRMEvent** se a configuração do log falhar de alguma forma, conforme mostrado na função a seguir.

  ```js
  function logCRMEvent(appointmentEvent) {
    event = appointmentEvent;
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          // Replace `event.completed({ allowEvent: false });` with the following statement.
          clearCustomProperties();
        }
      }
    );
  }
  ```

# <a name="task-pane"></a>[Painel de tarefas](#tab/taskpane)

Essa opção permitirá que um usuário registre e exiba suas anotações e outros detalhes sobre seus compromissos em um painel de tarefas.

### <a name="configure-the-manifest"></a>Configurar o manifesto

Para permitir que os usuários registrem anotações de compromisso com seu suplemento, você deve configurar o ponto de extensão [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) no manifesto sob o elemento pai `MobileFormFactor`. Não há suporte para outros fatores forma.

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó inteiro `<VersionOverrides>` (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir. Substitua todas as referências à **Contoso** com as informações da sua empresa.

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
                <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                  <OfficeTab id="TabDefault">
                    <Group id="apptReadGroup">
                      <Label resid="residDescription"/>
                      <Control xsi:type="Button" id="apptReadOpenPaneButton">
                        <Label resid="residLabel"/>
                        <Supertip>
                          <Title resid="residLabel"/>
                          <Description resid="residTooltip"/>
                        </Supertip>
                        <Icon>
                          <bt:Image size="16" resid="icon-16"/>
                          <bt:Image size="32" resid="icon-32"/>
                          <bt:Image size="80" resid="icon-80"/>
                        </Icon>
                        <Action xsi:type="ShowTaskpane">
                          <SourceLocation resid="Taskpane.Url"/>
                        </Action>
                      </Control>
                    </Group>
                  </OfficeTab>
                </ExtensionPoint>
              </DesktopFormFactor>
              <MobileFormFactor>
                <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
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
              <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
              <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
              <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
              <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
              <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
            </bt:ShortStrings>
            <bt:LongStrings>
              <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
            </bt:LongStrings>
          </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Para saber mais sobre manifestos para suplementos do Outlook, consulte [manifestos de suplemento do Outlook](manifests.md) e Adicionar suporte para comandos de suplemento [para Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Capturar anotações de compromisso

Nesta seção, saiba como exibir as anotações de compromisso registradas e outros detalhes em um painel de tarefas quando o usuário seleciona o **botão Log** .

1. No mesmo projeto de início rápido, abra o arquivo **./src/taskpane/taskpane.js** no editor de código.

1. Substitua todo o conteúdo do arquivo **taskpane.js** pelo JavaScript a seguir.

    ```js
    // Office is ready.
    Office.onReady(function () {
        getEventData();
      }
    );

    function getEventData() {
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("event logged successfully");
          } else {
            console.error("Failed to get body.");
          }
        }
      );
    }
    ```

Em seguida, atualize **otaskpane.html** para fazer **referênciataskpane.js**.

1. No mesmo projeto de início rápido, abra o **arquivo ./src/taskpane/taskpane.html** no editor de código.

1. Localize e substitua `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` pelo seguinte:

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="taskpane.js"></script>
    ```

### <a name="view-appointment-notes"></a>Exibir anotações de compromisso

O **rótulo** do botão Log pode ser alternado para exibir o **Modo** de Exibição definindo a **propriedade personalizada EventLogged** reservada para essa finalidade. Quando o usuário seleciona o **botão Exibir** , ele pode examinar suas anotações registradas para esse compromisso. Seu suplemento define a experiência de exibição de log.

Adicione a seguinte função **a ./src/taskpane/taskpane.js**. Essa função define a **propriedade personalizada EventLogged** no item de compromisso atual.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
            }
          }
        );
      }
    }
  );
}
```

Em seguida, chame-o depois que o suplemento registra com êxito as anotações do compromisso. Por exemplo, você pode chamá-lo de **getEventData** , conforme mostrado na função a seguir.

```js
function getEventData() {
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("event logged successfully");
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Excluir o log de compromissos

Se você quiser permitir que os usuários desfaçam o registro em log ou excluam as anotações de compromisso registradas para que um log de substituição possa ser salvo, você tem duas opções.

1. Use o Microsoft Graph [para limpar o objeto de propriedades personalizadas](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) quando o usuário selecionar o botão apropriado no painel de tarefas.
1. Adicione a função a seguir **a ./src/taskpane/taskpane.js** para limpar a propriedade **personalizada EventLogged** no item de compromisso atual.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                }
              }
            );
          }
        }
      );
    }
    ```

Em seguida, chame-o quando quiser limpar a propriedade personalizada. Por exemplo, você pode chamá-lo de **getEventData** se a configuração do log falhar de alguma forma, conforme mostrado na função a seguir.

  ```js
  function getEventData() {
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("event logged successfully");
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          clearCustomProperties();
        }
      }
    );
  }
  ```

---

## <a name="test-and-validate"></a>Testar e validar

1. Siga as diretrizes usuais [para testar e validar seu suplemento](testing-and-tips.md).
1. Depois de [fazer o sideload](sideload-outlook-add-ins-for-testing.md) do suplemento no Outlook na Web, Windows ou Mac, reinicie o Outlook em seu dispositivo móvel Android.
1. Abra um compromisso como participante e verifique se, no cartão **Insights** da Reunião, há um novo cartão com o nome do suplemento junto com o **botão Log** .

### <a name="ui-log-the-appointment-notes"></a>Interface do usuário: registrar as anotações do compromisso

Como participante da reunião, você deverá ver uma tela semelhante à imagem a seguir ao abrir uma reunião.

![Captura de tela mostrando o botão Log em uma tela de compromisso no Android.](../images/outlook-android-log-appointment-details.jpg)

### <a name="ui-view-the-appointment-log"></a>Interface do usuário: exibir o log de compromissos

Depois de registrar com êxito as anotações do compromisso, o botão agora deve ser rotulado **Exibir** em vez de **Log**. Você deverá ver uma tela semelhante à imagem a seguir.

![Captura de tela mostrando o botão Exibir em uma tela de compromisso no Android.](../images/outlook-android-view-appointment-log.jpg)

## <a name="available-apis"></a>APIs disponíveis

As APIs a seguir estão disponíveis para esse recurso.

- [APIs de caixa de diálogo](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- [APIs de Leitura de Compromisso (participante),](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true) **exceto** o seguinte:
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## <a name="restrictions"></a>Restrições

Várias restrições se aplicam.

- O **nome do botão** Log não pode ser alterado. No entanto, há uma maneira de um rótulo diferente ser exibido definindo uma propriedade personalizada no item de compromisso. Para obter mais detalhes, consulte a seção Exibir **anotações de compromisso** para [o comando de função](?tabs=noui#view-appointment-notes) ou [o painel de tarefas,](?tabs=taskpane#view-appointment-notes-1) conforme apropriado.
- A **propriedade personalizada EventLogged** deve ser usada se você quiser alternar o rótulo do botão **Log** para **Exibir** e voltar.
- O ícone de suplemento deve estar em escala de cinza usando código hexadecimal `#919191` ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).
- O suplemento deve extrair os detalhes da reunião do formulário de compromisso dentro do período de tempo limite de um minuto. No entanto, qualquer tempo gasto em uma caixa de diálogo que o suplemento abriu para autenticação, por exemplo, é excluído do período de tempo limite.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook Mobile](outlook-mobile-addins.md)
- [Adicionar suporte para comandos de suplementos para Outlook Mobile](add-mobile-support.md)
