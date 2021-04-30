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
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Configurar seu Outlook para ativação baseada em eventos (visualização)

Sem o recurso de ativação baseada em evento, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas. Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item. Você também pode se integrar ao painel de tarefas e à funcionalidade sem interface do usuário. Atualmente, os seguintes eventos são suportados.

|Evento|Descrição|
|---|---|
|`OnNewMessageCompose`|Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar), mas não ao editar, por exemplo, um rascunho.|
|`OnNewAppointmentOrganizer`|Ao criar um novo compromisso, mas não ao editar um existente.|
|`OnMessageAttachmentsChanged`|Ao adicionar ou remover anexos ao compor uma mensagem.|
|`OnAppointmentAttachmentsChanged`|Ao adicionar ou remover anexos durante a composição de um compromisso.|
|`OnMessageRecipientsChanged`|Ao adicionar ou remover destinatários ao compor uma mensagem.|
|`OnAppointmentAttendeesChanged`|Ao adicionar ou remover participantes durante a composição de um compromisso.|
|`OnAppointmentTimeChanged`|Ao alterar data/hora durante a composição de um compromisso.|
|`OnAppointmentRecurrenceChanged`|Ao adicionar, alterar ou remover os detalhes de recorrência ao compor um compromisso. Se a data/hora for alterada, `OnAppointmentTimeChanged` o evento também será acionado.|
|`OnInfoBarDismissClicked`|Ao descartar uma notificação ao compor uma mensagem ou item de compromisso. Somente o complemento que adicionou a notificação será notificado.|

No final deste passo a passo, você terá um complemento que é executado sempre que um novo item é criado e define o assunto.

> [!IMPORTANT]
> Esse recurso só é [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) suportado para visualização Outlook na Web e no Windows com uma assinatura Microsoft 365. Confira [Como visualizar o recurso de ativação](#how-to-preview-the-event-based-activation-feature) baseada em evento neste artigo para obter mais detalhes.
>
> Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.

## <a name="how-to-preview-the-event-based-activation-feature"></a>Como visualizar o recurso de ativação baseada em evento

Convidamos você a experimentar o recurso de ativação baseada em evento! Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback por meio GitHub (consulte a seção **Comentários** no final desta página).

Para visualizar esse recurso:

- Para Outlook na Web:
  - [Configure a versão direcionada em seu Microsoft 365 locatário](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).
  - Fazer referência **à biblioteca beta** no CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para a compilação typeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .
- Para Outlook no Windows: o build mínimo necessário é 16.0.13729.20000. Participe do [programa Office Insider](https://insider.office.com) para acessar Office beta.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido que cria um projeto de complemento com o gerador Yeoman para Office Desempois.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar a ativação baseada em evento do seu add-in, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` no nó do manifesto. Por enquanto, `DesktopFormFactor` é o único fator de formulário suportado.

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas abertas e próximas) e substitua-o pelo XML a seguir e salve suas alterações.

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

Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript. Você deve fornecer referências a ambos os arquivos no nó do manifesto como a plataforma Outlook finalmente determina se deve usar HTML ou JavaScript com base no cliente `Resources` Outlook. Como tal, para configurar o tratamento de eventos, forneça o local do HTML no elemento e, em seguida, em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado `Runtime` `Override` pelo HTML.

> [!TIP]
> Para saber mais sobre manifestos para Outlook de Outlook, [consulte Outlook manifestos de complemento.](manifests.md)

## <a name="implement-event-handling"></a>Implementar o tratamento de eventos

Você precisa implementar o tratamento para os eventos selecionados.

Nesse cenário, você adicionará a manipulação para compor novos itens.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

1. Após a `action` função, insira as seguintes funções JavaScript.

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

1. Para que as funções funcionem Outlook **na Web** com esse projeto gerado pelo gerador Yeoman para Office Desempois, adicione as instruções a seguir no final do arquivo.

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. Para que as funções funcionem em Outlook **em** Windows, adicione o código JavaScript a seguir no final do arquivo.

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **Observação**: verificar `Office.actions` se as Outlook na Web ignoram essas instruções.

1. Salve suas alterações.

## <a name="try-it-out"></a>Experimente

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.

    ```command&nbsp;line
    npm start
    ```

1. No Outlook na Web, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem Outlook na Web com o assunto definido como redação](../images/outlook-web-autolaunch-1.png)

1. Em Outlook no Windows, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem em Outlook no Windows com o assunto definido como redação](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Se você vir o erro "Não podemos abrir esse complemento do localhost", você precisará habilitar uma isenção de loopback.
    >
    > 1. Close Outlook.
    > 2. Abra o **Gerenciador de Tarefas** e certifique-se de que o **msoadfs.exe** não está em execução.
    > 3. Execute o seguinte comando.
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. Reinicie o Outlook.

## <a name="debug"></a>Depuração

À medida que você implementa sua própria funcionalidade, talvez seja necessário depurar seu código. Para obter orientações sobre como depurar a ativação de um add-in baseado em evento, consulte [Depurar](debug-autolaunch.md)seu Outlook de evento.

## <a name="event-based-activation-behavior-and-limitations"></a>Comportamento e limitações de ativação baseada em evento

Os complementos que são ativados com base em eventos devem ser curtos, leves e não invasivos possíveis. Para sinalizar que o seu complemento concluiu o processamento do evento de lançamento, recomendamos que você chame o método de seu `event.completed` complemento. Se essa chamada não for feita, o complemento terá um tempo limite de aproximadamente 300 segundos, o tempo máximo permitido para a execução de complementos baseados em eventos. O complemento também termina quando o usuário fecha a janela de composição.

Se o usuário tiver vários complementos que se inscrevem no mesmo evento, a plataforma Outlook iniciará os complementos em nenhuma ordem específica. Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente. Quaisquer complementos adicionais são pressionados para uma fila e executados conforme os complementos ativos anteriormente são concluídos ou desativados.

O usuário pode alternar ou navegar para longe do item de email atual onde o complemento começou a ser executado. O complemento que foi lançado terminará sua operação em segundo plano.

Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas de complementos baseados em eventos. Veja a seguir as APIs bloqueadas:

- Em `Office.context.auth` :
  - `getAccessToken`
  - `getAccessTokenAsync`
- Em `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Em `Office.context.mailbox.item` :
  - `close`
- Em `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>Confira também

[Outlook manifestos](manifests.md) 
 de complemento [Como depurar os complementos baseados em eventos](debug-autolaunch.md)
