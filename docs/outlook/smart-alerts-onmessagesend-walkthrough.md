---
title: Usar Alertas Inteligentes e os eventos OnMessageSend e OnAppointmentSend no suplemento Outlook (versão prévia)
description: Saiba como lidar com os eventos ao enviar em seu suplemento Outlook usando a ativação baseada em evento.
ms.topic: article
ms.date: 05/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0174d766423a9b70c67b0c2cf559f5b1ea24c9fe
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833912"
---
# <a name="use-smart-alerts-and-the-onmessagesend-and-onappointmentsend-events-in-your-outlook-add-in-preview"></a>Usar Alertas Inteligentes e os eventos OnMessageSend e OnAppointmentSend no suplemento Outlook (versão prévia)

Os `OnMessageSend` e `OnAppointmentSend` os eventos aproveitam os Alertas Inteligentes, que permitem que você execute a lógica depois que um  usuário seleciona Enviar em sua Outlook mensagem ou compromisso. O manipulador de eventos permite que você conceda aos usuários a oportunidade de melhorar seus emails e convites de reunião antes que eles sejam enviados.

O passo a passo a seguir usa o `OnMessageSend` evento. Ao final deste passo a passo, você terá um suplemento executado sempre que uma mensagem estiver sendo enviada e verificará se o usuário esqueceu de adicionar um documento ou imagem mencionado no email.

> [!IMPORTANT]
> Os `OnMessageSend` eventos `OnAppointmentSend` e os eventos só estão disponíveis em versão prévia com uma assinatura Microsoft 365 no Outlook no Windows. Para obter mais detalhes, [consulte Como visualizar](autolaunch.md#how-to-preview). Eventos de visualização não devem ser usados em suplementos de produção.

## <a name="prerequisites"></a>Pré-requisitos

O `OnMessageSend` evento está disponível por meio do recurso de ativação baseada em evento. Para entender como configurar seu suplemento para usar esse recurso, use outros eventos disponíveis, configure a visualização para esse evento, depure seu suplemento e muito mais, consulte Configurar seu suplemento [do Outlook para ativação](autolaunch.md) baseada em evento.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido, que cria um projeto de suplemento com o gerador Yeoman para Office suplementos.

## <a name="configure-the-manifest"></a>Configurar o manifesto

1. No editor de código, abra o projeto de início rápido.

1. Abra o **manifest.xml** arquivo localizado na raiz do seu projeto.

1. Selecione todo o **nó VersionOverrides** (incluindo marcas de abertura e fechamento) e substitua-o pelo XML a seguir e salve as alterações.

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

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
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

> [!TIP]
>
> - Para **obter as opções sendMode** disponíveis com os eventos `OnMessageSend` `OnAppointmentSend` e os eventos, consulte [as opções de SendMode disponíveis](/javascript/api/manifest/launchevent#available-sendmode-options-preview).
> - Para saber mais sobre manifestos para Outlook suplementos, [consulte Outlook manifestos de suplemento](manifests.md).

## <a name="implement-event-handling"></a>Implementar a manipulação de eventos

Você precisa implementar a manipulação para o evento selecionado.

Nesse cenário, você adicionará a manipulação para enviar uma mensagem. O suplemento verificará se há determinadas palavras-chave na mensagem. Se alguma dessas palavras-chave for encontrada, ela verificará se há anexos. Se não houver anexos, o suplemento recomendará que o usuário adicione o anexo possivelmente ausente.

1. No mesmo projeto de início rápido, crie uma nova pasta chamada **launchevent** no **diretório ./src** .

1. Na pasta **./src/launchevent** , crie um novo arquivo chamado **launchevent.js**.

1. Abra o arquivo **./src/launchevent/launchevent.js** no editor de código e adicione o código JavaScript a seguir.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
    }

    function getBodyCallback(asyncResult){
      let event = asyncResult.asyncContext;
      let body = "";
      if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        body = asyncResult.value;
      } else {
        let message = "Failed to get body text";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      let matches = hasMatches(body);
      if (matches) {
        Office.context.mailbox.item.getAttachmentsAsync(
          { asyncContext: event },
          getAttachmentsCallback);
      } else {
        event.completed({ allowEvent: true });
      }
    }

    function hasMatches(body) {
      if (body == null || body == "") {
        return false;
      }

      const arrayOfTerms = ["send", "picture", "document", "attachment"];
      for (let index = 0; index < arrayOfTerms.length; index++) {
        const term = arrayOfTerms[index].trim();
        const regex = RegExp(term, 'i');
        if (regex.test(body)) {
          return true;
        }
      }

      return false;
    }

    function getAttachmentsCallback(asyncResult) {
      let event = asyncResult.asyncContext;
      if (asyncResult.value.length > 0) {
        for (let i = 0; i < asyncResult.value.length; i++) {
          if (asyncResult.value[i].isInline == false) {
            event.completed({ allowEvent: true });
            return;
          }
        }

        event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
      } else {
        event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
      }
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. Salve suas alterações.

## <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

1. Abra o **webpack.config.js** arquivo encontrado no diretório raiz do projeto e conclua as etapas a seguir.

1. Localize `plugins` a matriz dentro do `config` objeto e adicione esse novo objeto ao início da matriz.

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

1. Execute os comandos a seguir no diretório raiz do seu projeto. Quando você executar `npm start`, o servidor Web local será iniciado (se ele ainda não estiver em execução) e seu suplemento será sideload.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Se o suplemento não foi carregado automaticamente no sideload, siga as instruções nos [suplementos do Sideload Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) para testar o sideload manual do suplemento no Outlook.

1. Em Outlook no Windows, crie uma nova mensagem e defina o assunto. No corpo, adicione texto como "Ei, confira esta foto do meu cachorro!".
1. Envie a mensagem. Uma caixa de diálogo deve aparecer com uma recomendação para adicionar um anexo.

    ![Caixa de diálogo recomendando que o usuário inclua um anexo.](../images/outlook-win-smart-alert.png)

1. Adicione um anexo e envie a mensagem novamente. Não deve haver nenhum alerta desta vez.

## <a name="smart-alerts-feature-behavior-and-scenarios"></a>Cenários e comportamento de recursos de Alertas Inteligentes

As descrições das **opções e recomendações do SendMode** para quando usá-las são detalhadas [nas opções Do SendMode disponíveis](/javascript/api/manifest/launchevent). O exemplo a seguir descreve o comportamento do recurso para determinados cenários.

### <a name="add-in-is-unavailable"></a>O suplemento não está disponível

Se o suplemento não estiver disponível quando uma mensagem ou compromisso estiver sendo enviado (por exemplo, ocorrerá um erro que impede o carregamento do suplemento), o usuário será alertado. As opções disponíveis para o usuário diferem dependendo da **opção SendMode** aplicada ao suplemento.

Se a opção ou for usada, o usuário poderá escolher  Enviar Mesmo Assim para enviar o item sem que o suplemento o verifique ou Tente  Mais Tarde para permitir que o item seja verificado pelo suplemento quando ele ficar disponível novamente.`PromptUser` `SoftBlock`

![Caixa de diálogo que alerta o usuário de que o suplemento não está disponível e dá ao usuário a opção de enviar o item agora ou mais tarde.](../images/outlook-soft-block-promptUser-unavailable.png)

Se a `Block` opção for usada, o usuário não poderá enviar o item até que o suplemento fique disponível.

![Caixa de diálogo que alerta o usuário de que o suplemento não está disponível. O usuário só pode enviar o item quando o suplemento estiver disponível novamente.](../images/outlook-hard-block-unavailable.png)

### <a name="long-running-add-in-operations"></a>Operações de suplemento de execução longa

Se o suplemento for executado por mais de cinco segundos, mas menos de cinco minutos, o usuário será alertado de que o suplemento está demorando mais do que o esperado para processar a mensagem ou o compromisso.

Se a `PromptUser` opção for usada, o usuário poderá escolher **Enviar** Mesmo Assim para enviar o item sem que o suplemento conclua sua verificação. Como alternativa, o usuário pode selecionar **Não Enviar** para interromper o processamento do suplemento.

![Caixa de diálogo que alerta o usuário de que o suplemento está demorando mais do que o esperado para processar o item. O usuário pode optar por enviar o item sem que o suplemento conclua sua verificação ou impedir que o suplemento processe o item.](../images/outlook-promptUser-long-running.png)

No entanto, se `SoftBlock` a opção ou `Block` for usada, o usuário não poderá enviar o item até que o suplemento conclua o processamento.

![Caixa de diálogo que alerta o usuário de que o suplemento está demorando mais do que o esperado para processar o item. O usuário deve aguardar até que o suplemento conclua o processamento do item antes que ele possa ser enviado.](../images/outlook-soft-hard-block-long-running.png)

`OnMessageSend` e `OnAppointmentSend` os suplementos devem ser de curta duração e leve. Para evitar a caixa de diálogo de operação de execução longa, use outros eventos para processar verificações condicionais antes que `OnMessageSend` o evento `OnAppointmentSend` ou o evento seja ativado. Por exemplo, se o usuário precisar criptografar anexos para cada mensagem ou compromisso, `OnMessageAttachmentsChanged` `OnAppointmentAttachmentsChanged` considere usar o evento ou o evento para executar a verificação.

### <a name="add-in-timed-out"></a>Tempo limite do suplemento

Se o suplemento for executado por cinco minutos ou mais, ele terá o tempo limite limite. Se a `PromptUser` opção for usada, o usuário poderá escolher **Enviar** Mesmo Assim para enviar o item sem que o suplemento conclua sua verificação. Como alternativa, o usuário pode escolher **Não Enviar**.

![Caixa de diálogo que alerta o usuário de que o processo de suplemento passou do tempo limite. O usuário pode optar por enviar o item sem que o suplemento conclua sua verificação ou não envie o item.](../images/outlook-promptUser-timeout.png)

Se a `SoftBlock` opção `Block` ou for usada, o usuário não poderá enviar o item até que o suplemento conclua sua verificação. O usuário deve tentar enviar o item novamente para reativar o suplemento.

![Caixa de diálogo que alerta o usuário de que o processo de suplemento passou do tempo limite. O usuário deve tentar enviar o item novamente para ativar o suplemento antes de poder enviar a mensagem ou o compromisso.](../images/outlook-soft-hard-block-timeout.png)

## <a name="limitations"></a>Limitações

Como os `OnMessageSend` eventos e `OnAppointmentSend` os eventos têm suporte por meio do recurso de ativação baseada em evento, as mesmas limitações de recursos se aplicam aos suplementos que são ativados como resultado desses eventos. Para obter uma descrição dessas limitações, consulte o comportamento e as limitações de [ativação baseada em eventos](autolaunch.md#event-based-activation-behavior-and-limitations).

Além dessas restrições, apenas uma instância de cada evento `OnMessageSend` `OnAppointmentSend` pode ser declarada no manifesto. Se você precisar de vários `OnMessageSend` ou `OnAppointmentSend` eventos, deverá declarar cada um deles em um manifesto ou suplemento separado.

Embora uma mensagem de diálogo alertas inteligentes possa ser alterada para se adequar ao seu cenário de suplemento usando a propriedade [errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions) do método event.completed, o seguinte não pode ser personalizado.

- A barra de título do diálogo. O nome do suplemento sempre é exibido lá.
- O formato da mensagem. Por exemplo, você não pode alterar o tamanho e a cor da fonte do texto nem inserir uma lista com marcadores.
- As opções de diálogo. Por exemplo, as **opções Enviar Mesmo** **Assim e** Não Enviar são fixas e dependem da [opção SendMode selecionada](/javascript/api/manifest/launchevent) .
- Caixas de diálogo de informações de progresso e processamento de ativação baseada em evento. Por exemplo, o texto e as opções que aparecem nas caixas de diálogo tempo limite e operação de execução longa não podem ser alterados.

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Configurar seu Outlook para ativação baseada em evento](autolaunch.md)
- [Como depurar suplementos baseados em eventos](debug-autolaunch.md)
- [Opções de listagem do AppSource para seu suplemento de Outlook baseado em evento](autolaunch-store-options.md)
