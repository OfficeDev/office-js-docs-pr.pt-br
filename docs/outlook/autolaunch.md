---
title: Configurar o suplemento do Outlook para ativação baseada em eventos
description: Saiba como configurar o suplemento do Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: ce2821ed5d226ff2c6a2b3c718d5711689523ac6
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664676"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Configurar o suplemento do Outlook para ativação baseada em eventos

Sem o recurso de ativação baseado em eventos, um usuário precisa iniciar explicitamente um suplemento para concluir suas tarefas. Esse recurso permite que seu suplemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item. Você também pode se integrar ao painel de tarefas e comandos de função.

Ao final deste passo a passo, você terá um suplemento que é executado sempre que um novo item for criado e definirá o assunto.

> [!NOTE]
> O suporte para esse recurso foi introduzido no [conjunto de requisitos 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), com eventos adicionais agora disponíveis em conjuntos de requisitos subsequentes. Para obter detalhes sobre o conjunto de requisitos mínimos de um evento e os clientes e plataformas que dão suporte a ele, consulte [Eventos com suporte](#supported-events) e [conjuntos de requisitos com suporte por servidores do Exchange e clientes do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
>
> A ativação baseada em eventos não tem suporte no Outlook no iOS ou no Android.

## <a name="supported-events"></a>Eventos com suporte

A tabela a seguir lista os eventos que estão disponíveis no momento e os clientes com suporte para cada evento. Quando um evento é gerado, o manipulador recebe um `event` objeto que pode incluir detalhes específicos para o tipo de evento. A coluna **Description** inclui um link para o objeto relacionado, quando aplicável.

|Nome canônico do evento</br>e nome do manifesto XML|Nome do manifesto do Teams|Descrição|Requisitos mínimos definidos e clientes com suporte|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |Ao compor uma nova mensagem (inclui resposta, responder tudo e encaminhar), mas não na edição, por exemplo, de um rascunho.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|Ao criar um novo compromisso, mas não na edição de um existente.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|Ao adicionar ou remover anexos ao compor uma mensagem.<br><br>Objeto de dados específico do evento: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|Ao adicionar ou remover anexos durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|Ao adicionar ou remover destinatários ao compor uma mensagem.<br><br>Objeto de dados específico do evento: [DestinatáriosChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|Ao adicionar ou remover participantes durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [DestinatáriosChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|Ao alterar a data/hora durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|Ao adicionar, alterar ou remover os detalhes de recorrência durante a composição de um compromisso. Se a data/hora for alterada, o `OnAppointmentTimeChanged` evento também será acionado.<br><br>Objeto de dados específico do evento: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|Ao descartar uma notificação ao compor uma mensagem ou item de compromisso. Somente o suplemento que adicionou a notificação será notificado.<br><br>Objeto de dados específico do evento: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>– Navegador da Web<br>- Nova interface do usuário do Mac|
|`OnMessageSend`|messageSending|Ao enviar um item de mensagem. Para saber mais, confira o [passo a passo alertas inteligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>– Navegador da Web|
|`OnAppointmentSend`|appointmentSending|Ao enviar um item de compromisso. Para saber mais, confira o [passo a passo alertas inteligentes](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>– Navegador da Web|
|`OnMessageCompose`|messageComposeOpened|Ao compor uma nova mensagem (inclui responder, responder tudo e encaminhar) ou editar um rascunho.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>– Navegador da Web|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|Ao criar um novo compromisso ou editar um existente.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>– Navegador da Web|

> [!NOTE]
> <sup>1</sup> Os suplementos baseados em eventos no Outlook no Windows exigem um mínimo de Windows 10 versão 1903 (Build 18362) ou Windows Server 2019 Versão 1903 para execução.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de suplemento com o [gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md).

> [!NOTE]
> Se você quiser usar o [manifesto do Teams para suplementos do Office (versão prévia),](../develop/json-manifest-overview.md) conclua o início rápido alternativo no [Outlook com um manifesto do Teams (versão prévia),](../quickstarts/outlook-quickstart-json-manifest.md) mas ignore todas as seções após a seção **Experimentar** .

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para configurar o manifesto, selecione a guia para o tipo de manifesto que você está usando.

# <a name="xml-manifest"></a>[Manifesto XML](#tab/xmlmanifest)

Para habilitar a ativação baseada em evento do suplemento, você deve configurar o elemento [Runtimes](/javascript/api/manifest/runtimes) e o `VersionOverridesV1_1` ponto de extensão [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) no nó do manifesto. Por enquanto, `DesktopFormFactor` é o único fator de formulário com suporte.

1. No editor de código, abra o projeto de início rápido.

1. Abra o arquivo **manifest.xml** localizado na raiz do projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas abertas e fechadas) e substitua-o pelo XML a seguir e salve suas alterações.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
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
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
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
        <!-- Entry needed for Outlook on Windows. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
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

O Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web e na nova interface do usuário do Mac usam um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript. Você deve fornecer referências a esses dois arquivos no `Resources` nó do manifesto, pois a plataforma do Outlook finalmente determina se deve usar HTML ou JavaScript com base no cliente do Outlook. Como tal, para configurar o tratamento de eventos, forneça o local do HTML no elemento e, em **\<Runtime\>** seguida, em seu `Override` elemento filho forneça o local do arquivo JavaScript inlined ou referenciado pelo HTML.

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

1. Abra o arquivo **manifest.json** .

1. Adicione o objeto a seguir à matriz "extensions.runtimes". Observe o seguinte sobre esta marcação:

   - A "minVersion" do conjunto de requisitos da caixa de correio é definida como "1.10" porque a tabela anterior neste artigo especifica que esta é a versão mais baixa do conjunto de requisitos que dá suporte aos `OnNewMessageCompose` eventos e `OnNewAppointmentCompose` .
   - A "id" do runtime é definida como o nome descritivo "autorun_runtime".
   - A propriedade "code" tem uma propriedade filho "page" definida como um arquivo HTML e uma propriedade filho "script" definida como um arquivo JavaScript. Você criará ou editará esses arquivos em etapas posteriores. O Office usa um desses valores dependendo da plataforma.
       - O Office no Windows executa os manipuladores de eventos em um runtime somente JavaScript, que carrega um arquivo JavaScript diretamente.
       - O Office no Mac e a Web executam os manipuladores em um runtime do navegador, que carrega um arquivo HTML. Esse arquivo, por sua vez, contém uma `<script>` marca que carrega o arquivo JavaScript.
     Para obter mais informações, consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).
   - A propriedade "lifetime" é definida como "curta", o que significa que o runtime é iniciado quando um dos eventos é disparado e desligado quando o manipulador é concluído. (Em certos casos raros, o runtime é desligado antes da conclusão do manipulador. Consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).)
   - Há dois tipos de "ações" que podem ser executadas no runtime. Você criará funções para corresponder a essas ações em uma etapa posterior.

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. Adicione a seguinte matriz "autoRunEvents" como uma propriedade do objeto na matriz "extensões".

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Adicione o objeto a seguir à matriz "autoRunEvents". A propriedade "eventos" mapeia manipuladores para eventos, conforme descrito na tabela anterior neste artigo. Os nomes do manipulador devem corresponder aos usados nas propriedades "id" dos objetos na matriz "actions" em uma etapa anterior.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - Para saber mais sobre runtimes em suplementos, consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).
> - Para saber mais sobre manifestos para suplementos do Outlook, confira [Manifestos de suplementos do Outlook](manifests.md).

## <a name="implement-event-handling"></a>Implementar o tratamento de eventos

Você precisa implementar o tratamento para seus eventos selecionados.

Nesse cenário, você adicionará o tratamento para compor novos itens.

1. No mesmo projeto de início rápido, crie uma nova pasta chamada **launchevent** no diretório **./src** .

1. Na pasta **./src/launchevent** , crie um novo arquivo chamado **launchevent.js**.

1. Abra o arquivo **./src/launchevent/launchevent.js** no editor de código e adicione o código JavaScript a seguir.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
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

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. Salve suas alterações.

> [!IMPORTANT]
> Windows: no momento, as importações não têm suporte no arquivo JavaScript em que você implementa o tratamento para ativação baseada em eventos.

## <a name="update-the-commands-html-file"></a>Atualizar o arquivo HTML de comandos

1. Na pasta **./src/commands** , abra **commands.html**.

1. Imediatamente antes da marca **de cabeça** de fechamento (`</head>`), adicione uma entrada de script para incluir o código JavaScript que manipula eventos.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Salve suas alterações.

## <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

1. Abra o arquivo **webpack.config.js** encontrado no diretório raiz do projeto e conclua as etapas a seguir.

1. Localize a `plugins` matriz dentro do `config` objeto e adicione este novo objeto no início da matriz.

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. Salve suas alterações.

## <a name="try-it-out"></a>Experimente

1. Execute os comandos a seguir no diretório raiz do seu projeto. Quando você executar `npm start`, o servidor Web local será iniciado (se ainda não estiver em execução) e o suplemento será sideload.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Se o suplemento não foi carregado automaticamente, siga as instruções nos [suplementos do Sideload Outlook para testar](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) para carregar manualmente o suplemento no Outlook.

1. No Outlook na Web, crie uma nova mensagem.

    ![Uma janela de mensagem no Outlook na Web com o assunto definido em composição.](../images/outlook-web-autolaunch-1.png)

1. No Outlook, na nova interface do usuário do Mac, crie uma nova mensagem.

    ![Uma janela de mensagem no Outlook sobre a nova interface do usuário do Mac com o assunto definido em composição.](../images/outlook-mac-autolaunch.png)

1. No Outlook no Windows, crie uma nova mensagem.

    ![Uma janela de mensagem no Outlook no Windows com o assunto definido em composição.](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>Depurar

Ao fazer alterações no tratamento de eventos de inicialização no suplemento, você deve estar ciente de que:

- Se você atualizou o manifesto, [remova o suplemento](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) e, em seguida, acaricie-o novamente. Se você estiver usando o Outlook no Windows, feche e reabra o Outlook.
- Se você fez alterações em arquivos diferentes do manifesto, feche e reabra o Outlook no Windows ou atualize a guia do navegador em execução Outlook na Web.

Ao implementar sua própria funcionalidade, talvez seja necessário depurar seu código. Para obter diretrizes sobre como depurar a ativação de suplemento baseada em eventos, consulte [Depurar seu suplemento do Outlook baseado em eventos](debug-autolaunch.md).

O log de runtime também está disponível para esse recurso no Windows. Para obter mais informações, confira [Depurar seu suplemento com o log de runtime](../testing/runtime-logging.md#runtime-logging-on-windows).

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>Implantar em usuários

Você pode implantar suplementos baseados em evento carregando o manifesto por meio do Centro de administração do Microsoft 365. No portal de administração, expanda a seção **Configurações** no painel de navegação e selecione **Aplicativos integrados**. Na página **Aplicativos integrados** , escolha a ação **Carregar aplicativos personalizados** .

![A página Aplicativos integrados no Centro de administração do Microsoft 365 com a ação Carregar aplicativos personalizados realçada.](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> Os suplementos baseados em eventos são restritos apenas a implantações gerenciadas por administradores. Os usuários não podem ativar suplementos baseados em eventos do AppSource ou da Office Store no aplicativo. Para saber mais, confira [Opções de listagem do AppSource para seu suplemento do Outlook baseado em eventos](autolaunch-store-options.md).

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>Comportamento e limitações de ativação baseados em eventos

Espera-se que os manipuladores de eventos de lançamento de suplemento sejam de execução curta, leve e o mais não invasivo possível. Após a ativação, o suplemento terá um tempo limite em aproximadamente 300 segundos, o tempo máximo permitido para executar suplementos baseados em evento. Para sinalizar que o suplemento concluiu o processamento de um evento de inicialização, o manipulador de eventos associado deve chamar o `event.completed` método. (Observe que o código incluído após a `event.completed` instrução não tem garantia de execução.) Sempre que um evento que seu suplemento manipula é disparado, o suplemento é reativado e executa o manipulador de eventos associado e a janela de tempo limite é redefinida. O suplemento termina após o tempo limite ou o usuário fecha a janela de composição ou envia o item.

Se o usuário tiver vários suplementos que se inscreveram no mesmo evento, a plataforma do Outlook iniciará os suplementos em nenhuma ordem específica. Atualmente, apenas cinco suplementos baseados em eventos podem estar em execução ativa.

Em todos os clientes do Outlook com suporte, o usuário deve permanecer no item de email atual em que o suplemento foi ativado para que ele seja concluído em execução. Navegar para longe do item atual (por exemplo, alternar para outra janela ou guia de composição) encerra a operação de suplemento. O suplemento também interrompe a operação quando o usuário envia a mensagem ou o compromisso que está compondo.

Não há suporte para importações no arquivo JavaScript em que você implementa o tratamento para ativação baseada em eventos no cliente Windows.

Algumas APIs Office.js que alteram ou alteram a interface do usuário não são permitidas de suplementos baseados em eventos. A seguir estão as APIs bloqueadas.

- Em `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > O [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) tem suporte em todas as versões do Outlook que dão suporte a ativação baseada em evento e SSO (logon único), enquanto [o Office.auth](/javascript/api/office/office.auth) só tem suporte em determinados builds do Outlook. Para obter mais informações, confira [Habilitar o SSO (logon único) em suplementos do Outlook que usam ativação baseada em evento](use-sso-in-event-based-activation.md).
- Em `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Em `Office.context.mailbox.item`:
  - `close`
- Em `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>Como solicitar dados externos

Você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou usando [XMLHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) uma API Web padrão que emite solicitações HTTP para interagir com servidores.

Lembre-se de que você deve usar medidas de segurança adicionais ao usar objetos XMLHttpRequest, exigindo [a Mesma Política de Origem](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) e [CORS simples (Compartilhamento de Recursos entre Origens)](https://developer.mozilla.org/docs/Web/HTTP/CORS).

Uma [implementação de CORS simples](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) :

- Não é possível usar cookies.
- Só dá suporte a métodos simples, como `GET`, `HEAD`e `POST`.
- Aceita cabeçalhos simples com nomes `Accept`de campo , `Accept-Language`ou `Content-Language`.
- Pode usar o `Content-Type`, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.
- Não é possível ter ouvintes de evento registrados no objeto retornado por `XMLHttpRequest.upload`.
- Não é possível usar `ReadableStream` objetos em solicitações.

> [!NOTE]
> O suporte completo do CORS está disponível em Outlook na Web, Mac e Windows (a partir da versão 2201, build 16.0.14813.10000).

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Como depurar suplementos baseados em eventos](debug-autolaunch.md)
- [Opções de listagem do AppSource para seu suplemento do Outlook baseado em evento](autolaunch-store-options.md)
- [Alertas inteligentes e passo a passo do OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Exemplos de código de suplementos do Office:
  - [Use a ativação baseada em eventos do Outlook para criptografar anexos, processar os participantes da solicitação de reunião e reagir às alterações de data/hora do compromisso](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Use a ativação baseada em eventos do Outlook para definir a assinatura](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Use a ativação baseada em eventos do Outlook para marcar destinatários externos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Usar Alertas Inteligentes do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
