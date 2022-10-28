---
title: Usar alertas inteligentes e os eventos OnMessageSend e OnAppointmentSend no suplemento do Outlook
description: Saiba como lidar com os eventos enviados no suplemento do Outlook usando a ativação baseada em eventos.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: a0fca566862455cd8a3981c1cfffba117145b39f
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767165"
---
# <a name="use-smart-alerts-and-the-onmessagesend-and-onappointmentsend-events-in-your-outlook-add-in"></a>Usar alertas inteligentes e os eventos OnMessageSend e OnAppointmentSend no suplemento do Outlook

Os `OnMessageSend` eventos e `OnAppointmentSend` aproveitam alertas inteligentes, o que permite que você execute a lógica depois que um usuário seleciona **Enviar** sua mensagem ou compromisso do Outlook. Seu manipulador de eventos permite que você dê aos seus usuários a oportunidade de melhorar seus emails e convites de reunião antes que eles sejam enviados.

O passo a passo a seguir usa o `OnMessageSend` evento. Ao final deste passo a passo, você terá um suplemento que é executado sempre que uma mensagem estiver sendo enviada e verifica se o usuário esqueceu de adicionar um documento ou imagem mencionado em seu email.

> [!NOTE]
> Os `OnMessageSend` eventos e `OnAppointmentSend` foram introduzidos no [conjunto de requisitos 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12). Confira, [clientes e plataformas](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) que oferecem suporte a esse conjunto de requisitos.

## <a name="prerequisites"></a>Pré-requisitos

O `OnMessageSend` evento está disponível por meio do recurso de ativação baseado em evento. Para entender como configurar o suplemento para usar esse recurso, use outros eventos disponíveis, depure seu suplemento e muito mais, consulte [Configurar o suplemento do Outlook para ativação baseada em eventos](autolaunch.md).

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), que cria um projeto de suplemento com o [gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para configurar o manifesto, selecione a guia para o tipo de manifesto que você está usando.

# <a name="xml-manifest"></a>[Manifesto XML](#tab/xmlmanifest)

1. No editor de código, abra o projeto de início rápido.

1. Abra o arquivo **manifest.xml** localizado na raiz do projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas abertas e fechadas) e substitua-o pelo XML a seguir e salve suas alterações.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.12">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and on the new Mac UI. -->
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

> [!TIP]
>
> - Para opções **SendMode** disponíveis com os `OnMessageSend` eventos e `OnAppointmentSend` , consulte [Opções de SendMode disponíveis](/javascript/api/manifest/launchevent#available-sendmode-options).
> - Para saber mais sobre manifestos para suplementos do Outlook, confira [Manifestos de suplementos do Outlook](manifests.md).

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

> [!IMPORTANT]
> Ainda não há suporte para alertas inteligentes para o [manifesto do Teams para Suplementos do Office (versão prévia)](../develop/json-manifest-overview.md). Essa guia é para uso futuro.

1. Abra o arquivo **manifest.json** .

1. Adicione o objeto a seguir à matriz "extensions.runtimes". Observe o seguinte sobre esta marcação:

   - A "minVersion" do conjunto de requisitos da caixa de correio é definida como "1.12" porque a [tabela de eventos com suporte](autolaunch.md#supported-events) especifica que esta é a versão mais baixa do conjunto de requisitos que dá suporte ao `OnMessageSend` evento.
   - A "id" do runtime é definida como o nome descritivo "autorun_runtime".
   - A propriedade "code" tem uma propriedade filho "page" definida como um arquivo HTML e uma propriedade filho "script" definida como um arquivo JavaScript. Você criará ou editará esses arquivos em etapas posteriores. O Office usa um desses valores ou outro, dependendo da plataforma.
       - O Office no Windows executa o manipulador de eventos em um runtime somente JavaScript, que carrega um arquivo JavaScript diretamente.
       - O Office no Mac e a Web executam o manipulador em um runtime do navegador, que carrega um arquivo HTML. Esse arquivo, por sua vez, contém uma `<script>` marca que carrega o arquivo JavaScript.
     Para obter mais informações, consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).
   - A propriedade "lifetime" é definida como "curta", o que significa que o runtime é iniciado quando o evento é disparado e desligado quando o manipulador é concluído. (Em certos casos raros, o runtime é desligado antes da conclusão do manipulador. Consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).)
   - Há uma ação para executar um manipulador para o `OnMessageSend` evento. Você criará a função de manipulador em uma etapa posterior.

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.12"
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
                "id": "onMessageSendHandler",
                "type": "executeFunction",
                "displayName": "onMessageSendHandler"
            }
        ]
    }
    ```

1. Adicione a seguinte matriz "autoRunEvents" como uma propriedade do objeto na matriz "extensões".

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Adicione o objeto a seguir à matriz "autoRunEvents". Observe o seguinte sobre este código:

   - O objeto de evento atribui uma função manipulador ao `OnMessageSend` evento (usando o nome do manifesto do Teams do evento, "messageSending", conforme descrito na [tabela de eventos com suporte](autolaunch.md#supported-events)). O nome da função fornecido em "actionId" deve corresponder ao nome usado na propriedade "id" do objeto na matriz "actions" em uma etapa anterior.
   - A opção "sendMode" está definida como "promptUser". Isso significa que, se a mensagem não atender às condições que o suplemento define para envio, o usuário será solicitado a cancelar o envio ou enviar de qualquer maneira.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.12"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
            {
                "type": "messageSending",
                "actionId": "onMessageSendHandler",
                "options": {
                    "sendMode": "promptUser"
                }
            }
          ]
      }
    ```

---

## <a name="implement-event-handling"></a>Implementar o tratamento de eventos

Você precisa implementar o tratamento para o evento selecionado.

Nesse cenário, você adicionará tratamento para enviar uma mensagem. Seu suplemento verificará se há determinadas palavras-chave na mensagem. Se alguma dessas palavras-chave for encontrada, ela verificará se há anexos. Se não houver anexos, seu suplemento recomendará ao usuário adicionar o anexo possivelmente ausente.

1. No mesmo projeto de início rápido, crie uma nova pasta chamada **launchevent** no diretório **./src** .

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

## <a name="update-the-commands-html-file"></a>Atualizar o arquivo HTML de comandos

1. Na pasta **./src/commands** , abra **commands.html**.

1. Imediatamente antes da marca **de cabeça** de fechamento (`</head>`), adicione uma entrada de script para o código JavaScript que manipula eventos.

   ```js
   <script type="text/javascript" src="../launchevent/launchevent.js"></script> 
   ```

1. Salve suas alterações.

## <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

1. Abra o arquivo **webpack.config.js** encontrado no diretório raiz do projeto e conclua as etapas a seguir.

1. Localize a `plugins` matriz dentro do `config` objeto e adicione este novo objeto ao início da matriz.

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

1. No Outlook no Windows, crie uma nova mensagem e defina o assunto. No corpo, adicione um texto como "Ei, confira esta foto do meu cachorro!".
1. Envie a mensagem. Uma caixa de diálogo deve aparecer com uma recomendação para que você adicione um anexo.

    ![A caixa de diálogo recomenda que o usuário inclua um anexo.](../images/outlook-win-smart-alert.png)

1. Adicione um anexo e envie a mensagem novamente. Não deve haver alerta desta vez.

## <a name="deploy-to-users"></a>Implantar em usuários

Semelhante a outros suplementos baseados em eventos, os suplementos que usam o recurso Alertas Inteligentes devem ser implantados pelo administrador de uma organização. Para obter diretrizes sobre como implantar seu suplemento por meio do Centro de administração do Microsoft 365, consulte a seção **Implantar para usuários** em [Configurar seu suplemento do Outlook para ativação baseada em eventos](autolaunch.md#deploy-to-users).

> [!IMPORTANT]
> Os suplementos que usam o recurso Alertas Inteligentes só poderão ser publicados no AppSource se a [propriedade SendMode](/javascript/api/manifest/launchevent#available-sendmode-options) do manifesto estiver definida como a opção `SoftBlock` ou `PromptUser` . Se a propriedade **SendMode** de um suplemento estiver definida como `Block`, ela só poderá ser implantada pelo administrador de uma organização, pois falhará na validação do AppSource. Para saber mais sobre como publicar seu suplemento baseado em evento no AppSource, confira [Opções de listagem do AppSource para seu suplemento do Outlook baseado em eventos](autolaunch-store-options.md).

## <a name="smart-alerts-feature-behavior-and-scenarios"></a>Comportamento e cenários de recursos de Alertas Inteligentes

As descrições das opções e recomendações **sendmode** para quando usá-las são detalhadas nas [opções SendMode disponíveis](/javascript/api/manifest/launchevent#available-sendmode-options). A seguir, descreve o comportamento do recurso para determinados cenários.

### <a name="add-in-is-unavailable"></a>O suplemento não está disponível

Se o suplemento não estiver disponível quando uma mensagem ou compromisso estiver sendo enviado (por exemplo, ocorrerá um erro que impede o carregamento do suplemento), o usuário será alertado. As opções disponíveis para o usuário diferem dependendo da opção **SendMode** aplicada ao suplemento.

Se a opção `PromptUser` ou `SoftBlock` for usada, o usuário poderá escolher **Enviar De qualquer maneira** para enviar o item sem que o suplemento o verifique ou **Tente Posteriormente** para permitir que o item seja verificado pelo suplemento quando ele ficar disponível novamente.

![Caixa de diálogo que alerta o usuário de que o suplemento não está disponível e dá ao usuário a opção de enviar o item agora ou posterior.](../images/outlook-soft-block-promptUser-unavailable.png)

Se a opção `Block` for usada, o usuário não poderá enviar o item até que o suplemento fique disponível. (A opção `Block` não terá suporte se o suplemento usar um manifesto do Teams (versão prévia).

![Caixa de diálogo que alerta o usuário de que o suplemento não está disponível. O usuário só pode enviar o item quando o suplemento estiver disponível novamente.](../images/outlook-hard-block-unavailable.png)

### <a name="long-running-add-in-operations"></a>Operações de suplemento de execução longa

Se o suplemento for executado por mais de cinco segundos, mas menos de cinco minutos, o usuário será alertado de que o suplemento está demorando mais do que o esperado para processar a mensagem ou o compromisso.

Se a opção `PromptUser` for usada, o usuário poderá escolher **Enviar De qualquer maneira** para enviar o item sem que o suplemento conclua sua verificação. Como alternativa, o usuário pode selecionar **Não Enviar** para impedir o processamento do suplemento.

![Caixa de diálogo que alerta o usuário de que o suplemento está demorando mais do que o esperado para processar o item. O usuário pode optar por enviar o item sem que o suplemento conclua sua verificação ou impeça que o suplemento processe o item.](../images/outlook-promptUser-long-running.png)

No entanto, se a opção `SoftBlock` ou `Block` for usada, o usuário não poderá enviar o item até que o suplemento conclua o processamento.

![Caixa de diálogo que alerta o usuário de que o suplemento está demorando mais do que o esperado para processar o item. O usuário deve aguardar até que o suplemento conclua o processamento do item antes que ele possa ser enviado.](../images/outlook-soft-hard-block-long-running.png)

`OnMessageSend` e `OnAppointmentSend` os suplementos devem ser de execução curta e leve. Para evitar a caixa de diálogo de operação de longa execução, use outros eventos para processar verificações condicionais antes que o `OnMessageSend` evento ou `OnAppointmentSend` seja ativado. Por exemplo, se o usuário for obrigado a criptografar anexos para cada mensagem ou compromisso, considere usar o `OnMessageAttachmentsChanged` evento ou `OnAppointmentAttachmentsChanged` para executar a verificação.

### <a name="add-in-timed-out"></a>Tempo limite de complemento

Se o suplemento for executado por cinco minutos ou mais, ele terá um tempo limite. Se a opção `PromptUser` for usada, o usuário poderá escolher **Enviar De qualquer maneira** para enviar o item sem que o suplemento conclua sua verificação. Como alternativa, o usuário pode escolher **Não Enviar**.

![Caixa de diálogo que alerta o usuário de que o processo de suplemento está com o tempo limite. O usuário pode optar por enviar o item sem que o suplemento conclua sua verificação ou não envie o item.](../images/outlook-promptUser-timeout.png)

Se a opção `SoftBlock` ou `Block` for usada, o usuário não poderá enviar o item até que o suplemento conclua sua verificação. O usuário deve tentar enviar o item novamente para reativar o suplemento.

![Caixa de diálogo que alerta o usuário de que o processo de suplemento está com o tempo limite. O usuário deve tentar enviar o item novamente para ativar o suplemento antes de poder enviar a mensagem ou compromisso.](../images/outlook-soft-hard-block-timeout.png)

## <a name="limitations"></a>Limitações

Como os `OnMessageSend` eventos e `OnAppointmentSend` têm suporte por meio do recurso de ativação baseado em evento, as mesmas limitações de recurso se aplicam a suplementos que são ativados como resultado desses eventos. Para obter uma descrição dessas limitações, consulte [Comportamento e limitações de ativação baseados em eventos](autolaunch.md#event-based-activation-behavior-and-limitations).

Além dessas restrições, apenas uma instância do `OnMessageSend` evento e `OnAppointmentSend` pode ser declarada no manifesto. Se você precisar de vários `OnMessageSend` eventos ou `OnAppointmentSend` , você deve declarar cada um em um suplemento separado.

Embora uma mensagem de diálogo Alertas Inteligentes possa ser alterada para se adequar ao cenário de suplemento usando a [propriedade errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions) do método event.completed, o seguinte não pode ser personalizado.

- A barra de título da caixa de diálogo. O nome do suplemento sempre é exibido lá.
- O formato da mensagem. Por exemplo, você não pode alterar o tamanho e a cor da fonte do texto ou inserir uma lista com marcadores.
- As opções de caixa de diálogo. Por exemplo, as opções **Enviar De Qualquer Maneira** e **Não Enviar** são fixas e dependem da [opção SendMode](/javascript/api/manifest/launchevent#available-sendmode-options) selecionada.
- Diálogos de informações de processamento e progresso baseados em eventos. Por exemplo, o texto e as opções que aparecem nas caixas de diálogo de operação de tempo limite e de longa execução não podem ser alteradas.

## <a name="differences-between-smart-alerts-and-the-on-send-feature"></a>Diferenças entre alertas inteligentes e o recurso de envio

Embora alertas inteligentes e o [recurso de envio](outlook-on-send-addins.md) forneçam aos seus usuários a oportunidade de melhorar suas mensagens e convites de reunião antes de serem enviados, alertas inteligentes são um recurso mais recente que oferece mais flexibilidade com a forma como você solicita aos usuários mais ações. As principais diferenças entre os dois recursos são descritas na tabela a seguir.

> [!IMPORTANT]
> Ainda não há suporte para alertas inteligentes para o manifesto do Teams (versão prévia). Estamos trabalhando para fornecer esse suporte em breve.

|Atributo|Alertas Inteligentes|Em envio|
|-----|-----|-----|
|**Conjunto de requisitos com suporte mínimo**|[Caixa de correio 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)|[Caixa de correio 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8)|
|**Clientes do Outlook com suporte**|-Windows<br>– Navegador da Web (interface do usuário moderna)<br>- Mac (nova interface do usuário)|-Windows<br>– Navegador da Web (interface do usuário clássica e moderna)<br>- Mac (interface do usuário clássica e nova) |
|**Eventos com suporte**|**Manifesto XML**<br>- `OnMessageSend`<br>- `OnAppointmentSend`<br><br>**Manifesto do Teams (versão prévia)**<br>- "messageSending"<br>- "appointmentSending"|**Manifesto XML**<br>- `ItemSend`<br><br>**Manifesto do Teams (versão prévia)**<br>– Sem suporte|
|**Propriedade de extensão manifest**|**Manifesto XML**<br>- `LaunchEvent`<br><br>**Manifesto do Teams (versão prévia)**<br>- "AutoRunEvents"|**Manifesto XML**<br>- `Events`<br><br>**Manifesto do Teams (versão prévia)**<br>– Sem suporte|
|**Opções de modo de envio com suporte**|– Usuário prompt<br>- Bloco macio<br>- Bloquear (sem suporte se o suplemento usar um manifesto do Teams (versão prévia))|Bloquear|
|**Número máximo de eventos com suporte em um suplemento**|Um `OnMessageSend` e um `OnAppointmentSend` evento.|Um `ItemSend` evento.|
|**Implantação de suplemento**|O suplemento poderá ser publicado no AppSource se sua `SendMode` propriedade estiver definida como a opção `SoftBlock` ou `PromptUser` . Caso contrário, o suplemento deve ser implantado pelo administrador de uma organização.|O suplemento não pode ser publicado no AppSource. Ele deve ser implantado pelo administrador de uma organização.|
|**Configuração adicional para instalação de suplemento**|Nenhuma configuração adicional é necessária depois que o manifesto é carregado no Centro de administração do Microsoft 365.|Dependendo dos padrões de conformidade da organização e do cliente do Outlook usado, determinadas políticas de caixa de correio devem ser configuradas para instalar o suplemento.|

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Configurar o suplemento do Outlook para ativação baseada em eventos](autolaunch.md)
- [Como depurar suplementos baseados em eventos](debug-autolaunch.md)
- [Opções de listagem do AppSource para seu suplemento do Outlook baseado em evento](autolaunch-store-options.md)
- [Exemplo de código de suplementos do Office: usar alertas inteligentes do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
