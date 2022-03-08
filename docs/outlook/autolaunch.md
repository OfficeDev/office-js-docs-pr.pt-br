---
title: Configurar seu complemento do Outlook para ativação baseada em eventos
description: Saiba como configurar seu complemento do Outlook para ativação baseada em eventos.
ms.topic: article
ms.date: 03/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7d63e814875ee36a24bf7a919da0b62562433af0
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340285"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Configurar seu complemento do Outlook para ativação baseada em eventos

Sem o recurso de ativação baseada em evento, um usuário precisa iniciar explicitamente um complemento para concluir suas tarefas. Esse recurso permite que o seu complemento execute tarefas com base em determinados eventos, especialmente para operações que se aplicam a cada item. Você também pode se integrar ao painel de tarefas e à funcionalidade sem interface do usuário.

No final deste passo a passo, você terá um complemento que é executado sempre que um novo item é criado e define o assunto.

> [!NOTE]
> O suporte para esse recurso foi introduzido no [conjunto de requisitos 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md). Confira, [clientes e plataformas](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="supported-events"></a>Eventos com suporte

A tabela a seguir lista os eventos que estão disponíveis no momento e os clientes com suporte para cada evento. Quando um evento é gerado, o manipulador recebe um `event` objeto que pode incluir detalhes específicos do tipo de evento. A **coluna Descrição** inclui um link para o objeto relacionado quando aplicável.

> [!IMPORTANT]
> Os eventos ainda em visualização só podem estar disponíveis com uma assinatura do Microsoft 365 e em um conjunto limitado de clientes com suporte, conforme notado na tabela a seguir. Para obter detalhes de configuração do cliente, [consulte Como visualizar](#how-to-preview) neste artigo. Eventos de visualização não devem ser usados em complementos de produção.

|Evento|Descrição|Conjunto de requisitos mínimos e clientes com suporte|
|---|---|---|
|`OnNewMessageCompose`|Ao compor uma nova mensagem (inclui responder, responder a todos e encaminhar), mas não ao editar, por exemplo, um rascunho.|[1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br><br>- Windows<br>- Navegador da Web<br>- Nova visualização da interface do usuário do Mac|
|`OnNewAppointmentOrganizer`|Ao criar um novo compromisso, mas não ao editar um existente.|[1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br><br>- Windows<br>- Navegador da Web<br>- Nova visualização da interface do usuário do Mac|
|`OnMessageAttachmentsChanged`|Ao adicionar ou remover anexos ao compor uma mensagem.<br><br>Objeto de dados específico do evento: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnAppointmentAttachmentsChanged`|Ao adicionar ou remover anexos durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnMessageRecipientsChanged`|Ao adicionar ou remover destinatários ao compor uma mensagem.<br><br>Objeto de dados específico do evento: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnAppointmentAttendeesChanged`|Ao adicionar ou remover participantes durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnAppointmentTimeChanged`|Ao alterar data/hora durante a composição de um compromisso.<br><br>Objeto de dados específico do evento: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnAppointmentRecurrenceChanged`|Ao adicionar, alterar ou remover os detalhes de recorrência ao compor um compromisso. Se a data/hora for alterada, o `OnAppointmentTimeChanged` evento também será acionado.<br><br>Objeto de dados específicos do evento: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnInfoBarDismissClicked`|Ao descartar uma notificação ao compor uma mensagem ou item de compromisso. Somente o complemento que adicionou a notificação será notificado.<br><br>Objeto de dados específico do evento: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](../reference/objectmodel/requirement-set-1.11/outlook-requirement-set-1.11.md)<br><br>- Windows<br>- Navegador da Web|
|`OnMessageSend`|Ao enviar um item de mensagem. Para saber mais, consulte o passo a passo [alertas inteligentes](smart-alerts-onmessagesend-walkthrough.md).|[Visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)<br><br>- Windows|
|`OnAppointmentSend`|Ao enviar um item de compromisso. Para saber mais, consulte o passo a passo [alertas inteligentes](smart-alerts-onmessagesend-walkthrough.md).|[Visualização](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)<br><br>- Windows|

### <a name="how-to-preview"></a>Como visualizar

Convidamos você a experimentar os eventos agora na visualização! Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback por meio do GitHub (consulte a seção **Comentários** no final desta página).

Para visualizar esses eventos quando disponível:

- Para o Outlook na Web:
  - [Configure a versão direcionada em seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).
  - Fazer referência **à biblioteca beta** na CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) da compilação TypeScript e IntelliSense pode ser encontrado na CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview`.
- Para o Outlook na nova visualização da interface do usuário do Mac:
  - O build mínimo necessário é 16,54 (21101001). Participe do [programa Office Insider](https://insider.office.com/join/Mac) e escolha o **Canal Beta** para acesso a builds beta do Office.
- Para o Outlook no Windows:
  - O build mínimo necessário é 16.0.14511.10000. Participe do [programa Office Insider](https://insider.office.com/join/windows) e escolha o **Canal Beta** para acesso a builds beta do Office.
  - Configure o Registro. O Outlook inclui uma cópia local das versões de produção e beta do Office.js em vez de carregar da CDN (rede de distribuição de conteúdo). Por padrão, a cópia de produção local da API é referenciada. Para alternar para a cópia beta local das APIs JavaScript do Outlook, você precisa adicionar essa entrada do Registro, caso contrário, as APIs beta podem não ser encontradas.
    1. Crie a chave do Registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.
    1. Adicione uma entrada chamada `EnableBetaAPIsInJavaScript` e desmarcar o valor como `1`. A imagem a seguir mostra qual deve ser a aparência de registro.

        ![Captura de tela do editor do Registro com um valor de chave do Registro EnableBetaAPIsInJavaScript.](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de complemento com o gerador Yeoman para Os Complementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar a ativação baseada em evento do seu add-in, você deve configurar o elemento [Runtimes](../reference/manifest/runtimes.md) e o ponto de extensão [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) `VersionOverridesV1_1` no nó do manifesto. Por enquanto, `DesktopFormFactor` é o único fator de formulário suportado.

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione o nó inteiro `<VersionOverrides>` (incluindo marcas abertas e próximas) e substitua-o pelo XML a seguir e salve suas alterações.

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
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
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
              
              <!-- Other available events (currently released) -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              -->

              <!-- Other available events (currently in preview) -->
              <!--
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
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
        <!-- Entry needed for Outlook Desktop. -->
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

Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web e na nova visualização da interface do usuário do Mac usam um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript. Você deve fornecer referências `Resources` a ambos os arquivos no nó do manifesto, pois a plataforma Outlook determina se deve usar HTML ou JavaScript com base no cliente Outlook. Como tal, para configurar o tratamento de eventos, forneça o local do HTML `Runtime` no elemento e, em seguida, `Override` em seu elemento filho forneça o local do arquivo JavaScript embutido ou referenciado pelo HTML.

> [!TIP]
> Para saber mais sobre manifestos para Outlook de Outlook, [consulte Outlook manifestos de complemento](manifests.md).

## <a name="implement-event-handling"></a>Implementar o tratamento de eventos

Você precisa implementar o tratamento para os eventos selecionados.

Nesse cenário, você adicionará a manipulação para compor novos itens.

1. No mesmo projeto de início rápido, crie uma nova pasta chamada **launchevent** no **diretório /src/** .

1. Na pasta **./src/launchevent** , crie um novo arquivo chamado **launchevent.js**.

1. Abra o arquivo **./src/launchevent/launchevent.js** no editor de código e adicione o seguinte código JavaScript.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

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
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. Salve suas alterações.

> [!IMPORTANT]
> Windows: no momento, as importações não são suportadas no arquivo JavaScript onde você implementa o tratamento para a ativação baseada em eventos.

## <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

Abra o **webpack.config.js** arquivo encontrado no diretório raiz do projeto e conclua as etapas a seguir.

1. Localize `plugins` a matriz dentro do `config` objeto e adicione esse novo objeto no início da matriz.

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

1. Execute os seguintes comandos no diretório raiz do seu projeto. Quando você executar `npm start`, o servidor Web local será acionado (se ele ainda não estiver em execução) e o seu complemento será sideload.

    ```command&nbsp;line
    npm run build
    ```
    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Se o seu add-in não foi automaticamente sideload, siga as instruções em [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manualmente sideload the add-in in Outlook.

1. No Outlook na Web, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem Outlook na Web com o assunto definido na composição.](../images/outlook-web-autolaunch-1.png)

1. Em Outlook na nova visualização da interface do usuário do Mac, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem Outlook na nova visualização da interface do usuário do Mac com o assunto definido na composição.](../images/outlook-mac-autolaunch.png)

1. Em Outlook no Windows, crie uma nova mensagem.

    ![Captura de tela de uma janela de mensagem Outlook no Windows com o assunto definido na redação.](../images/outlook-win-autolaunch.png)

    [!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="debug"></a>Depurar

À medida que você faz alterações no tratamento de eventos de início no seu complemento, você deve estar ciente de que:

- Se você atualizou o manifesto, [remova o complemento](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) e, em seguida, o sideload novamente. Se você estiver usando o Outlook no Windows, feche-o e reabra-o.
- Se você fez alterações em arquivos que não o manifesto, feche e reabra o Outlook no Windows ou atualize a guia do navegador executando Outlook na Web.

Ao implementar sua própria funcionalidade, talvez seja necessário depurar seu código. Para obter orientações sobre como depurar a ativação de um add-in baseado em eventos, consulte [Depurar seu Outlook de evento](debug-autolaunch.md).

O log de tempo de execução também está disponível para esse recurso em Windows. Para obter mais informações, consulte [Depurar seu add-in com o log de tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows).

## <a name="deploy-to-users"></a>Implantar para usuários

Você pode implantar os complementos baseados em eventos carregando o manifesto por meio do Centro de administração do Microsoft 365. No portal de administração, expanda a **seção Configurações** no painel de navegação e selecione **Aplicativos integrados**. Na página **Aplicativos integrados**, escolha a **ação Upload aplicativos personalizados**.

![Captura de tela da página Aplicativos integrados no Centro de administração do Microsoft 365, incluindo a ação Upload aplicativos personalizados.](../images/outlook-deploy-event-based-add-ins.png)

AppSource e no Office Store: a capacidade de implantar os complementos baseados em eventos ou atualizar os complementos existentes para incluir o recurso de ativação baseada em eventos deve estar disponível em breve.

> [!IMPORTANT]
> Os complementos baseados em eventos são restritos apenas a implantações gerenciadas pelo administrador. Por enquanto, os usuários não podem obter os complementos baseados em eventos no AppSource ou no Office Store. Para saber mais, consulte As opções de [listagem do AppSource para seu Outlook de evento](autolaunch-store-options.md).

## <a name="event-based-activation-behavior-and-limitations"></a>Comportamento e limitações de ativação baseada em evento

Espera-se que os manipuladores de eventos de início do add-in sejam curtos, leves e não invasivos possíveis. Após a ativação, o seu complemento terá um tempo limite de aproximadamente 300 segundos, o tempo máximo permitido para a execução de complementos baseados em eventos. Para sinalizar que o seu complemento concluiu o processamento de um evento de lançamento, recomendamos que o manipulador associado chame o `event.completed` método. (Observe que o código incluído após a instrução `event.completed` não é garantido para ser executado.) Sempre que um evento que seu complemento lida é disparado, o complemento é reativado e executa o manipulador de eventos associado e a janela de tempo de tempo é redefinida. O complemento termina após o tempo final, ou o usuário fecha a janela de redação ou envia o item.

Se o usuário tiver vários complementos inscritos no mesmo evento, a plataforma Outlook iniciará os complementos em nenhuma ordem específica. Atualmente, apenas cinco complementos baseados em eventos podem ser executados ativamente.

O usuário pode alternar ou navegar para longe do item de email atual onde o complemento começou a ser executado. O complemento que foi lançado terminará sua operação em segundo plano.

As importações não são suportadas no arquivo JavaScript em que você implementa o tratamento para a ativação baseada em eventos no cliente Windows.

Algumas Office.js APIs que alteram ou alteram a interface do usuário não são permitidas de complementos baseados em eventos. A seguir estão as APIs bloqueadas.

- Em `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > `OfficeRuntime.auth` tem suporte. Para obter mais informações, consulte [Enable single sign-on (SSO) in Outlook add-ins that use event-based activation](use-sso-in-event-based-activation.md).
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

Você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)uma API Web padrão que emite solicitações HTTP para interagir com servidores.

Esteja ciente de que você deve usar medidas de segurança adicionais ao criar XmlHttpRequests, exigindo a Política de Mesma [Origem](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) e [o CORS simples](https://www.w3.org/TR/cors/).

Uma implementação de CORS simples não pode usar cookies e só oferece suporte a métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar um `Content-Type` header em CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.

O suporte completo ao CORS está chegando em breve.

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Como depurar os complementos baseados em eventos](debug-autolaunch.md)
- [Opções de listagem do AppSource para seu Outlook de evento](autolaunch-store-options.md)
- [Alertas Inteligentes e Passo a passo do OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Exemplos pnP:
  - [Use Outlook ativação baseada em eventos para criptografar anexos, processar os participantes da solicitação de reunião e reagir às alterações de data/hora do compromisso](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Use a ativação baseada em eventos do Outlook para definir a assinatura](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Use a ativação baseada em eventos do Outlook para marcar destinatários externos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
