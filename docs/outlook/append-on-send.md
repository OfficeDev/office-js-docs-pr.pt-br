---
title: Implementar o append-on-send em seu suplemento do Outlook
description: Saiba como implementar o recurso append-on-send no suplemento do Outlook.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: c8239634b6c9ca281255caf89276fb1b454efc84
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767158"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Implementar o append-on-send em seu suplemento do Outlook

Ao final deste passo a passo, você terá um suplemento do Outlook que pode inserir um aviso de isenção de responsabilidade quando uma mensagem é enviada.

> [!NOTE]
> O suporte para esse recurso foi introduzido no conjunto de requisitos 1.9. Confira, [clientes e plataformas](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) que oferecem suporte a esse conjunto de requisitos.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua o [início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) que cria um projeto de suplemento com o gerador Yeoman para Suplementos do Office.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para configurar o manifesto, abra a guia para o tipo de manifesto que você está usando.

# <a name="xml-manifest"></a>[Manifesto XML](#tab/xmlmanifest)

Para habilitar o recurso append-on-send em seu suplemento, você deve incluir a `AppendOnSend` permissão na coleção de [ExtendedPermissions](/javascript/api/manifest/extendedpermissions).

Para esse cenário, em vez de executar a `action` função ao escolher o botão **Executar uma ação** , você executará a `appendOnSend` função.

1. No editor de código, abra o projeto de início rápido.

1. Abra o arquivo **manifest.xml** localizado na raiz do projeto.

1. Selecione o nó inteiro **\<VersionOverrides\>** (incluindo marcas abertas e fechadas) e substitua-o pelo XML a seguir.

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.9">
            <bt:Set Name="Mailbox" />
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
            <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

> [!IMPORTANT]
> Ainda não há suporte para o [manifesto do Teams para Suplementos do Office (versão prévia)](../develop/json-manifest-overview.md). Essa guia é para uso futuro.

1. Abra o arquivo manifest.json.

1. Adicione o objeto a seguir à matriz "extensions.runtimes". Observe o seguinte sobre este código.

   - A "minVersion" do conjunto de requisitos da caixa de correio é definida como "1.9" para que o suplemento não possa ser instalado em plataformas e versões do Office em que esse recurso não tem suporte. 
   - A "id" do runtime é definida como o nome descritivo "function_command_runtime".
   - A propriedade "code.page" é definida como a URL do arquivo HTML sem interface do usuário que carregará o comando da função.
   - A propriedade "lifetime" é definida como "curta", o que significa que o runtime é iniciado quando o botão de comando da função é selecionado e é desligado quando a função é concluída. (Em certos casos raros, o runtime é desligado antes da conclusão do manipulador. Consulte [Runtimes em Suplementos do Office](../testing/runtimes.md).)
   - Há uma ação para executar uma função chamada "appendDisclaimerOnSend". Você criará essa função em uma etapa posterior.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.9"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "function_command_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "appendDisclaimerOnSend",
                "type": "executeFunction",
                "displayName": "appendDisclaimerOnSend"
            }
        ]
    }
    ```

1. Na matriz "authorization.permissions.resourceSpecific", adicione o objeto a seguir. Certifique-se de que ele está separado de outros objetos na matriz com uma vírgula.

    ```json
    {
      "name": "Mailbox.AppendOnSend.User",
      "type": "Delegated"
    }
    ```

---

> [!TIP]
> Para saber mais sobre manifestos para suplementos do Outlook, confira [Manifestos de suplementos do Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Implementar o tratamento de apêndice em envio

Em seguida, implemente a anexação no evento de envio.

> [!IMPORTANT]
> Se o suplemento também implementar o [tratamento de eventos no envio usando `ItemSend`](outlook-on-send-addins.md), chamar `AppendOnSendAsync` o manipulador de envio retornará um erro, pois esse cenário não terá suporte.

Para esse cenário, você implementará a anexação de uma isenção de responsabilidade ao item quando o usuário enviar.

1. No mesmo projeto de início rápido, abra o arquivo **./src/commands/commands.js** no editor de código.

1. Após a `action` função, insira a função JavaScript a seguir.

    ```js
    function appendDisclaimerOnSend(event) {
      const appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. Imediatamente abaixo da função, adicione a linha a seguir para registrar a função.

    ```js
    Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
    ```

## <a name="try-it-out"></a>Experimente

1. Execute o seguinte comando no diretório raiz do seu projeto. Quando você executar esse comando, o servidor Web local será iniciado se ele ainda não estiver em execução e seu suplemento for sideload.

    ```command&nbsp;line
    npm start
    ```

1. Crie uma nova mensagem e adicione-se à linha **To** .

1. No menu faixa de opções ou estouro, escolha **Executar uma ação**.

1. Envie a mensagem e abra-a na **caixa de entrada** ou na pasta **Itens Enviados** para exibir o aviso de isenção de responsabilidade acrescentado.

    ![Uma mensagem de exemplo com o aviso de isenção de responsabilidade acrescentado ao enviar Outlook na Web.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>Confira também

[Manifestos de suplementos do Outlook](manifests.md)
