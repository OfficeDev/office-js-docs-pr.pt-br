---
title: Use Alertas Inteligentes e o evento OnMessageSend no seu Outlook de usuário (visualização)
description: Saiba como lidar com o evento enviar mensagem em seu Outlook de envio usando a ativação baseada em evento.
ms.topic: article
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78e10f8609264d69ba32b78badc14c626c210d76
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681757"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>Use Alertas Inteligentes e o evento OnMessageSend no seu Outlook de usuário (visualização)

O evento tira proveito dos Alertas Inteligentes que permitem executar a lógica depois que um usuário seleciona `OnMessageSend` **Enviar** em sua Outlook mensagem. O manipulador de eventos permite que você dê aos usuários a oportunidade de melhorar seus emails antes que eles são enviados. O `OnAppointmentSend` evento é semelhante, mas se aplica a um compromisso.

No final deste passo a passo, você terá um complemento que é executado sempre que uma mensagem estiver sendo enviada e verifica se o usuário esqueceu de adicionar um documento ou imagem mencionado no email.

> [!IMPORTANT]
> Os eventos e só estão disponíveis na visualização com uma assinatura Microsoft 365 no Outlook `OnMessageSend` `OnAppointmentSend` no Windows. Para obter mais detalhes, consulte [Como visualizar](autolaunch.md#how-to-preview). Eventos de visualização não devem ser usados em complementos de produção.

## <a name="prerequisites"></a>Pré-requisitos

O `OnMessageSend` evento está disponível por meio do recurso de ativação baseada em evento. Para entender sobre como configurar seu add-in para usar esse recurso, eventos disponíveis, como visualizar esse evento, depuração, limitações de recursos e muito mais, consulte [Configure your Outlook add-in for event-based activation](autolaunch.md).

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [Outlook início](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) rápido que cria um projeto de complemento com o gerador Yeoman para Office Desempois.

## <a name="configure-the-manifest"></a>Configurar o manifesto

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

> [!TIP]
>
> - Para **opções sendMode** disponíveis com o `OnMessageSend` evento, consulte Opções de [SendMode disponíveis.](../reference/manifest/launchevent.md#available-sendmode-options-preview)
> - Para saber mais sobre manifestos para Outlook de Outlook, [consulte Outlook manifestos de complemento.](manifests.md)

## <a name="implement-event-handling"></a>Implementar o tratamento de eventos

Você precisa implementar o tratamento para o evento selecionado.

Nesse cenário, você adicionará a manipulação para enviar uma mensagem. O seu complemento verificará se há determinadas palavras-chave na mensagem. Se alguma dessas palavras-chave for encontrada, ela verificará se há anexos. Se não houver anexos, o seu complemento recomendará ao usuário adicionar o anexo possivelmente ausente.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

1. Após a `action` função, insira as seguintes funções JavaScript.

    ```js
    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          var event = asyncResult.asyncContext;
          var body = "";
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }
        
          var arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (var index = 0; index < arrayOfTerms.length; index++) {
            var term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }
        
          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result){
                var event = asyncResult.asyncContext;
                if (result.value.length <= 0) {
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (var i=0;i<result.value.length;i++) {
                    if(result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
                    
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }
    ```

1. Adicione o código JavaScript a seguir no final do arquivo.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. Salve suas alterações.

> [!IMPORTANT]
> Windows: no momento, as importações não são suportadas no arquivo JavaScript onde você implementa o tratamento para a ativação baseada em eventos.

## <a name="try-it-out"></a>Experimente

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executa este comando, o servidor web local será iniciado (se ainda não estiver em execução) e seu suplemento será transferido.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Se o seu add-in não foi automaticamente sideload, siga as instruções em [Sideload Outlook add-ins](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) for testing to manually sideload the add-in in Outlook.

1. Em Outlook no Windows, crie uma nova mensagem e de definir o assunto. No corpo, adicione texto como "Ei, confira esta imagem do meu cachorro!".
1. Envie a mensagem. Uma caixa de diálogo deve aparecer com uma recomendação para adicionar um anexo.
1. Adicione um anexo e envie a mensagem novamente. Não deve haver nenhum alerta desta vez.

> [!NOTE]
> Se você estiver executando o seu complemento no localhost e vir o erro "Lamentamos, não foi possível acessar *{your-add-in-name-here}*. Certifique-se de ter uma conexão de rede. Se o problema continuar, tente novamente mais tarde.", talvez seja necessário habilitar uma isenção de loopback.
>
> 1. Close Outlook.
> 1. Abra o **Gerenciador de Tarefas** e certifique-se de que o **msoadfsb.exe** não está em execução.
> 1. Execute o seguinte comando.
>
>    ```command&nbsp;line
>    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>    ```
>
> 1. Reinicie o Outlook.

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Configurar seu Outlook para ativação baseada em eventos](autolaunch.md)
- [Como depurar os complementos baseados em eventos](debug-autolaunch.md)
- [Opções de listagem do AppSource para seu Outlook de evento](autolaunch-store-options.md)