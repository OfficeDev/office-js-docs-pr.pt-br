---
title: Configure seu Outlook complemento para ativação baseada em eventos (visualização)
description: Saiba como configurar seu Outlook complemento para ativação baseada em eventos.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555226"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Configure seu Outlook complemento para ativação baseada em eventos (visualização)

Sem o recurso de ativação baseado em eventos, o usuário precisa lançar explicitamente um complemento para concluir suas tarefas. Esse recurso permite que seu complemento execute tarefas com base em determinados eventos, particularmente para operações que se aplicam a cada item. Você também pode se integrar com a funcionalidade painel de tarefas e sem interface do usuário.

Ao final deste passo a passo, você terá um complemento que é executado sempre que um novo item for criado e definir o assunto.

> [!IMPORTANT]
> Este recurso só é suportado para [visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) em Outlook na web e em Windows com uma assinatura Microsoft 365. Para obter mais detalhes, consulte [Como visualizar o recurso de ativação baseado](#how-to-preview-the-event-based-activation-feature) em eventos neste artigo.
>
> Como os recursos de visualização estão sujeitos a alterações sem aviso prévio, eles não devem ser usados em complementos de produção.

## <a name="supported-events"></a>Eventos com suporte

No momento, os seguintes eventos são apoiados.

|Evento|Descrição|Clientes|
|---|---|---|
|`OnNewMessageCompose`|Ao compor uma nova mensagem (inclui resposta, resposta a todos e adiante), mas não na edição, por exemplo, de um rascunho.|Windows, web|
|`OnNewAppointmentOrganizer`|Ao criar um novo compromisso, mas não na edição de um já existente.|Windows, web|
|`OnMessageAttachmentsChanged`|Ao adicionar ou remover anexos enquanto compõe uma mensagem.|Windows|
|`OnAppointmentAttachmentsChanged`|Ao adicionar ou remover anexos enquanto compõe uma consulta.|Windows|
|`OnMessageRecipientsChanged`|Ao adicionar ou remover destinatários enquanto compõe uma mensagem.|Windows|
|`OnAppointmentAttendeesChanged`|Ao adicionar ou remover os participantes enquanto compõe uma consulta.|Windows|
|`OnAppointmentTimeChanged`|Ao alterar a data/hora enquanto compõe uma consulta.|Windows|
|`OnAppointmentRecurrenceChanged`|Ao adicionar, alterar ou remover os detalhes de recorrência enquanto compõe uma consulta. Se a data/hora for alterada, o `OnAppointmentTimeChanged` evento também será acionado.|Windows|
|`OnInfoBarDismissClicked`|Ao descartar uma notificação enquanto compõe uma mensagem ou item de nomeação. Apenas o complemento que adicionou a notificação será notificado.|Windows|

## <a name="how-to-preview-the-event-based-activation-feature"></a>Como visualizar o recurso de ativação baseado em eventos

Convidamos você a experimentar o recurso de ativação baseado em eventos! Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback através de GitHub (veja a seção **Feedback** no final desta página).

Para visualizar este recurso:

- Para Outlook na web:
  - [Configure a liberação direcionada no seu inquilino Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).
  - Consulte a biblioteca **beta** na CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) para compilação TypeScript e IntelliSense é encontrado no CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview` .
- Para Outlook em Windows:
  - A construção mínima exigida é de 16.0.14026.20000. Junte-se ao [programa Office Insider](https://insider.office.com) para acesso a Office compilações beta.
  - Configure o registro. Outlook inclui uma cópia local da produção e versões beta de Office.js em vez de carregar a partir do CDN. Por padrão, a cópia de produção local da API é referenciada. Para mudar para a cópia beta local das APIs javaScript Outlook, você precisa adicionar esta entrada de registro, caso contrário, as APIs beta podem não ser encontradas.
    1. Crie a chave de registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` .
    1. Adicione uma entrada nomeada `EnableBetaAPIsInJavaScript` e defina o valor para `1` . A imagem a seguir mostra qual deve ser a aparência de registro.

        ![Captura de tela do editor de registro com um valor-chave de registro EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Complete o [Outlook início rápido](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto adicional com o gerador Yeoman para Office Add-ins.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para ativar a ativação baseada em eventos do seu complemento, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) no `VersionOverridesV1_1` nó do manifesto. Por enquanto, `DesktopFormFactor` é o único fator de forma suportado.

1. Em seu editor de código, abra o projeto de início rápido.

1. Abra o **arquivomanifest.xml** localizado na raiz do seu projeto.

1. Selecione todo o `<VersionOverrides>` nó (incluindo tags abertas e fechadas) e substitua-o pelo XML a seguir e salve suas alterações.

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

Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript. Você deve fornecer referências a ambos os arquivos no `Resources` nó do manifesto, pois a plataforma Outlook determina em última instância se deve usar HTML ou JavaScript com base no Outlook cliente. Como tal, para configurar o manuseio do evento, forneça a localização do HTML no `Runtime` elemento e, em seguida, em seu `Override` elemento filho forneça a localização do arquivo JavaScript ininlinado ou referenciado pelo HTML.

> [!TIP]
> Para saber mais sobre manifestos para Outlook complementos, consulte [Outlook manifestos adicionais](manifests.md).

## <a name="implement-event-handling"></a>Implementar o manuseio de eventos

Você tem que implementar o manuseio para seus eventos selecionados.

Neste cenário, você adicionará manuseio para compor novos itens.

1. A partir do mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** em seu editor de código.

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

1. Adicione o seguinte código JavaScript no final do arquivo.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. Salve suas alterações.

## <a name="try-it-out"></a>Experimente

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Se o seu complemento não tiver sido carregado automaticamente, siga as instruções em [sideload Outlook complementos para testar](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) para carregar manualmente o complemento Outlook.

1. No Outlook na Web, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem em Outlook na web com o assunto definido na composição](../images/outlook-web-autolaunch-1.png)

1. Em Outlook no Windows, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem em Outlook em Windows com o assunto definido na composição](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Se você estiver executando seu complemento do localhost e veja o erro "Sentimos muito, não podemos acessar *{seu-add-in-name-aqui}*. Certifique-se de ter uma conexão de rede. Se o problema continuar, tente novamente mais tarde.", você pode precisar habilitar uma isenção de loopback.
    >
    > 1. Close Outlook.
    > 1. Abra o **Gerenciador de Tarefas** e garanta que o processo **demsoadfsb.exe** não esteja em execução.
    > 1. Execute o seguinte comando.
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. Reinicie o Outlook.

## <a name="debug"></a>depurar

À medida que você faz alterações no manuseio do evento de lançamento em seu complemento, você deve estar ciente de que:

- Se você atualizou o manifesto, [remova o complemento](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) e, em seguida, o der descarga novamente.
- Se você fez alterações em outros arquivos além do manifesto, feche e reabra Outlook na Windows ou atualize a guia do navegador executando Outlook na web.

Ao implementar sua própria funcionalidade, você pode precisar depurar seu código. Para obter orientações sobre como depurar a ativação complementa baseada em eventos, consulte [Depurar seu complemento de Outlook baseado em eventos](debug-autolaunch.md).

O registro de tempo de execução também está disponível para este recurso em Windows. Para obter mais informações, consulte [Depurar seu complemento com o registro de tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows).

## <a name="deploy-to-users"></a>Implantar para usuários

Você pode implantar complementos baseados em eventos carregando o manifesto através do centro administrativo Microsoft 365. No portal de administração, expanda a seção **Configurações** no painel de navegação e selecione **aplicativos Integrados**. Na página **aplicativos Integrados,** escolha a Upload ação **de aplicativos personalizados.**

![Captura de tela da página de aplicativos integrados no centro administrativo Microsoft 365, incluindo a ação de aplicativos personalizados Upload](../images/outlook-deploy-event-based-add-ins.png)

AppSource e lojas de inclientes: A capacidade de implantar complementos baseados em eventos ou atualizar complementos existentes para incluir o recurso de ativação baseado em eventos deve estar disponível em breve.

> [!IMPORTANT]
> Os complementos baseados em eventos são restritos apenas a implantações gerenciadas por administradores. Por enquanto, os usuários não podem obter complementos baseados em eventos do AppSource ou lojas de incientes.

## <a name="event-based-activation-behavior-and-limitations"></a>Comportamento e limitações de ativação baseadas em eventos

Espera-se que os manipuladores de eventos de lançamento adicionais sejam de curta duração, leves e não invasivos. Após a ativação, seu complemento será cronometrados dentro de aproximadamente 300 segundos, o tempo máximo permitido para executar complementos baseados em eventos. Para sinalizar que seu complemento completou o processamento de um evento de lançamento, recomendamos que o manipulador associado chame o `event.completed` método. (Observe que o código incluído após a `event.completed` declaração não é garantido para execução.) Cada vez que um evento que seu complemento é acionado, o complemento é reativado e executa o manipulador de eventos associado e a janela de tempo limite é redefinida. O complemento termina depois que ele se esgota, ou o usuário fecha a janela de composição ou envia o item.

Se o usuário tiver vários complementos que se inscreveram no mesmo evento, a plataforma Outlook lança os complementos em nenhuma ordem específica. Atualmente, apenas cinco complementos baseados em eventos podem estar sendo executados ativamente.

O usuário pode alternar ou navegar para longe do item de e-mail atual, onde o complemento começou a ser executado. O complemento que foi lançado terminará sua operação em segundo plano.

Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas a partir de complementos baseados em eventos. A seguir, as APIs bloqueadas:

- Em `OfficeRuntime.auth` :
  - `getAccessToken`(somente Windows)
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

- [Manifestos de suplementos do Outlook](manifests.md)
- [Como depurar complementos baseados em eventos](debug-autolaunch.md)
