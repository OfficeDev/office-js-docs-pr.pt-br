---
title: Configurar seu complemento do Outlook para ativação baseada em eventos (visualização)
description: Saiba como configurar seu complemento do Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 02/03/2021
localization_priority: Normal
ms.openlocfilehash: d9108b4debea5e59503f3c935a537e5fafde00c8
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234272"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Configurar seu complemento do Outlook para ativação baseada em eventos (visualização)

Sem o recurso de ativação baseada em eventos, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas. Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item. Você também pode integrar com o painel de tarefas e a funcionalidade sem interface do usuário. Atualmente, os seguintes eventos são suportados.

- `OnNewMessageCompose`: Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar)
- `OnNewAppointmentOrganizer`: Ao criar um novo compromisso

  > [!IMPORTANT]
  > Esse recurso não **é** ativado na edição de um item, por exemplo, um rascunho ou um compromisso existente.

No final deste passo a passo, você terá um complemento que é executado sempre que uma nova mensagem é criada.

> [!IMPORTANT]
> Esse recurso só tem suporte para [visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web e no Windows com uma assinatura do Microsoft 365. Veja [como visualizar o recurso de ativação baseada em eventos](#how-to-preview-the-event-based-activation-feature) neste artigo para obter mais detalhes.
>
> Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.

## <a name="how-to-preview-the-event-based-activation-feature"></a>Como visualizar o recurso de ativação baseada em eventos

Convidamos você a experimentar o recurso de ativação baseada em eventos! Conheça seus cenários e como podemos melhorar nos fazendo comentários por meio do GitHub (confira a seção **Comentários** no final desta página).

Para visualizar esse recurso:

- Para o Outlook na Web:
  - [Configure o lançamento direcionado no locatário do Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)
  - Fazer referência **à biblioteca beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . O [arquivo de definição de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) tipo para compilação de TypeScript e IntelliSense é encontrado na CDN e [definitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .
- Para o Outlook no Windows: o build mínimo necessário é 16.0.13729.20000. Participe do [programa Office Insider](https://insider.office.com) para acessar as versões beta do Office.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar a ativação baseada em eventos do seu complemento, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) no `VersionOverridesV1_1` nó do manifesto. Por enquanto, `DesktopFormFactor` é o único fator forma com suporte.

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do projeto.

1. Selecione o nó `<VersionOverrides>` inteiro (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir.

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

O Outlook no Windows usa um arquivo JavaScript, enquanto o Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript. Você deve fornecer referências a ambos os arquivos no nó do manifesto, pois a plataforma do Outlook finalmente determina se deve usar HTML ou JavaScript com base no `Resources` cliente do Outlook. Dessa forma, para configurar a manipulação de eventos, forneça o local do HTML no elemento e, em seguida, em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado `Runtime` `Override` pelo HTML.

> [!TIP]
> Para saber mais sobre manifestos para os complementos do Outlook, confira [manifestos de complementos do Outlook.](manifests.md)

## <a name="implement-event-handling"></a>Implementar a manipulação de eventos

Você precisa implementar a manipulação para os eventos selecionados.

Neste cenário, você adicionará a manipulação para composição de novos itens.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** seu editor de código.

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

1. Para que as funções funcionem no Outlook na **Web** com esse projeto gerado pelo gerador Yeoman para Os Complementos do Office, adicione as instruções a seguir no final do arquivo.

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. Para as funções funcionarem no **Outlook no Windows,** adicione o seguinte código JavaScript no final do arquivo.

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **Observação:** a verificação `Office.actions` garante que o Outlook na Web ignore essas instruções.

## <a name="try-it-out"></a>Experimente

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executar esse comando, o servidor Web local será iniciar (se ainda não estiver em execução) e o seu complemento será sideloaded.

    ```command&nbsp;line
    npm start
    ```

1. No Outlook na Web, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem no Outlook na Web com o assunto definido na composição](../images/outlook-web-autolaunch-1.png)

1. No Outlook no Windows, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem no Outlook no Windows com o assunto definido na composição](../images/outlook-win-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>Comportamento e limitações da ativação baseada em eventos

Espera-se que os complementos ativados com base em eventos sejam de curta duração, leve e o mais não ofensivo possível. Para sinalizar que o seu complemento concluiu o processamento do evento de lançamento, recomendamos que você chame o `event.completed` método. Se essa chamada não for feita, o tempo limite do complemento será de aproximadamente 300 segundos, o período máximo permitido para a execução de complementos baseados em eventos. O complemento também termina quando o usuário fecha a janela de redação.

Se o usuário tiver vários complementos que se inscrevem no mesmo evento, a plataforma do Outlook inicia os complementos sem uma ordem específica. Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente. Quaisquer outros complementos são pressionados para uma fila e executados à medida que os complementos ativos anteriormente são concluídos ou desativados.

O usuário pode alternar ou sair do item de email atual onde o complemento começou a ser executado. O complemento que foi lançado concluirá sua operação em segundo plano.

Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas a partir de complementos baseados em eventos. Veja a seguir as APIs bloqueadas:

- Em `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Em `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`
- Em `Office.context.auth` :
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a>Confira também

[Manifestos de suplementos do Outlook](manifests.md)