---
title: Ativar o suplemento do Outlook em várias mensagens (versão prévia)
description: Saiba como ativar o suplemento do Outlook quando várias mensagens são selecionadas.
ms.topic: article
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9d81d698facfc4535b3945d8cee4c97492fc8a88
ms.sourcegitcommit: 5544cf174d145e356e33866e2480bde999514ada
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/14/2022
ms.locfileid: "68574141"
---
# <a name="activate-your-outlook-add-in-on-multiple-messages-preview"></a>Ativar o suplemento do Outlook em várias mensagens (versão prévia)

Com o recurso de seleção múltipla de itens, o suplemento do Outlook agora pode ativar e executar operações em várias mensagens selecionadas de uma só vez. Determinadas operações, como carregar mensagens no sistema CRM (Gerenciamento de Relacionamento com o Cliente) ou categorizar vários itens, agora podem ser facilmente concluídas com um único clique.

As seções a seguir explicam como configurar seu suplemento para recuperar a linha de assunto de várias mensagens no modo de leitura.

> [!IMPORTANT]
> O recurso de seleção de vários itens só está disponível em versão prévia com uma assinatura do Microsoft 365 no Outlook no Windows. Os recursos em versão prévia não devem ser usados em suplementos de produção. Convidamos você a testar esse recurso em ambientes de teste ou desenvolvimento e receber comentários de boas-vindas sobre sua experiência por meio  do GitHub (consulte a seção Comentários no final desta página).

> [!NOTE]
> No momento, não há suporte para o recurso de seleção de itens no manifesto do [Teams (](../develop/json-manifest-overview.md)versão prévia), mas a equipe está trabalhando para disponibilizar isso.

## <a name="prerequisites-to-preview-item-multi-select"></a>Pré-requisitos para visualizar a seleção de vários itens

Para visualizar o recurso de seleção multissessão, instale o Outlook no Windows, começando com a versão 2209 (Build 15629.20110). Depois de instalado, ingresse no [programa Office Insider](https://insider.office.com/join/windows) e selecione a **opção Canal Beta** para acessar builds beta do Office.

## <a name="set-up-your-environment"></a>Configurar seu ambiente

Conclua [o início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) para criar um projeto de suplemento com o [gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Configurar o manifesto

Para habilitar o suplemento para ativar em várias mensagens selecionadas, você deve adicionar o elemento filho [SupportsMultiSelect](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsmultiselect-preview) **\<Action\>** ao elemento e definir seu valor como `true`. Como a seleção de vários itens dá suporte apenas a mensagens no momento, **\<ExtensionPoint\>** `xsi:type` o valor do atributo do elemento deve ser definido como `MessageReadCommandSurface` ou `MessageComposeCommandSurface`.

1. No editor de código preferido, abra o projeto de início rápido do Outlook que você criou.

1. Abra o **manifest.xml** arquivo localizado na raiz do projeto.

1. Atribua o **\<Permissions\>** valor ao `ReadWriteMailbox` elemento.

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. Selecione o nó **\<VersionOverrides\>** inteiro e substitua-o pelo XML a seguir.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.12">
                  <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                        <Label resid="TaskpaneButton.Label"/>
                                        <Supertip>
                                            <Title resid="TaskpaneButton.Label"/>
                                            <Description resid="TaskpaneButton.Tooltip"/>
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="Icon.16x16"/>
                                            <bt:Image size="32" resid="Icon.32x32"/>
                                            <bt:Image size="80" resid="Icon.80x80"/>
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="Taskpane.Url"/>
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
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
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. Salve suas alterações.

## <a name="configure-the-task-pane"></a>Configurar o painel de tarefas

A seleção de vários itens depende do [evento SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) para determinar quando as mensagens são selecionadas ou desmarcadas. Esse evento requer uma implementação do painel de tarefas.

1. Na pasta **./src/taskpane** , abra **taskpane.html**.

1. No elemento **\<script\>** , defina o `src` atributo como `"https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"`. Isso faz referência à biblioteca beta na CDN (rede de distribuição de conteúdo).

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
    ```

1. No elemento **\<body\>** , substitua todo o elemento **\<main\>** pela marcação a seguir.

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Salve suas alterações.

## <a name="implement-a-handler-for-the-selecteditemschanged-event"></a>Implementar um manipulador para o evento SelectedItemsChanged

Para alertar o suplemento quando o `SelectedItemsChanged` evento ocorrer, você deve registrar um manipulador de eventos usando o `addHandlerAsync` método.

1. Na pasta **./src/taskpane** , abra **taskpane.js**.

1. Na função `Office.onReady()` de retorno de chamada, substitua o código existente pelo seguinte:

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## <a name="retrieve-the-subject-line-of-selected-messages"></a>Recuperar a linha de assunto das mensagens selecionadas

Agora que você registrou um manipulador de eventos, chame o método [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) para recuperar a linha de assunto das mensagens selecionadas e registrá-las no painel de tarefas. O `getSelectedItemsAsync` método também pode ser usado para obter outras propriedades de mensagem, como a ID do item, o tipo de item (`Message` é o único tipo com suporte no momento) e o modo de item (`Read` ou `Compose`).

1. No **taskpane.js**, navegue até a função `run` e insira o código a seguir.

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. Salve suas alterações.

## <a name="try-it-out"></a>Experimente

1. Em um terminal, execute o código a seguir no diretório raiz do projeto. Isso inicia o servidor Web local e o sideload do suplemento.

    ```command line
    npm start
    ```

    > [!TIP]
    > Se o suplemento não realizar o sideload automaticamente, siga as instruções em [Fazer sideload de suplementos do Outlook](sideload-outlook-add-ins-for-testing.md?tabs=windows#outlook-on-the-desktop) para teste para realizar o sideload manual no Outlook.

1. No Outlook, verifique se o Painel de Leitura está habilitado. Para habilitar o Painel de Leitura, [consulte Usar e configurar o Painel de Leitura para visualizar mensagens](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).

1. Navegue até a caixa de entrada e escolha várias mensagens mantendo **a tecla Ctrl pressionada** enquanto seleciona mensagens.

1. Selecione **Mostrar Painel de Tarefas** na faixa de opções.

1. No painel de tarefas, selecione **Executar** para exibir uma lista das linhas de assunto das mensagens selecionadas.

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="Uma lista de exemplos de linhas de assunto recuperadas de várias mensagens selecionadas.":::

## <a name="item-multi-select-behavior-and-limitations"></a>Limitações e comportamento de seleção de vários itens

A seleção de vários itens só dá suporte a mensagens em uma caixa de correio do Exchange nos modos de leitura e redação. Um suplemento do Outlook só será ativado em várias mensagens se as condições a seguir forem atendidas.

- As mensagens devem ser selecionadas em uma caixa de correio do Exchange por vez. Não há suporte para caixas de correio que não são do Exchange.
- As mensagens devem ser selecionadas em uma pasta de caixa de correio por vez. Um suplemento não será ativado em várias mensagens se elas estiverem localizadas em pastas diferentes, a menos que o modo de exibição Conversas esteja habilitado. Para obter mais informações, consulte [Multisseleção em conversas](#multi-select-in-conversations).
- Um suplemento deve implementar um painel de tarefas para detectar o `SelectedItemsChanged` evento.
- O [Painel de Leitura](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) no Outlook deve estar habilitado.
- No máximo 100 mensagens podem ser selecionadas por vez.

> [!NOTE]
> Convites e respostas de reunião são considerados mensagens, não compromissos e, portanto, podem ser incluídos em uma seleção.

### <a name="multi-select-in-conversations"></a>Seleção multisseletor em conversas

A seleção de vários itens dá suporte [à](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) exibição Conversas, independentemente de estar habilitada em sua caixa de correio ou em pastas específicas. A tabela a seguir descreve os comportamentos esperados quando as conversas são expandidas ou recolhidas, quando o cabeçalho da conversa é selecionado e quando as mensagens de conversa estão localizadas em uma pasta diferente da que está atualmente em exibição.

|Selection|Exibição de conversa expandida|Exibição de conversa recolhida|
|------|------|------|
|**O cabeçalho da conversa está selecionado**|Se o cabeçalho da conversa for o único item selecionado, um suplemento que dá suporte a várias seleções não será ativado. No entanto, se outras mensagens que não são de cabeçalho também forem selecionadas, o suplemento só será ativado nessas mensagens e não no cabeçalho selecionado.|A mensagem mais recente (ou seja, a primeira mensagem na pilha de conversa) está incluída na seleção da mensagem.<br><br>Se a mensagem mais recente na conversa estiver localizada em outra pasta da que está atualmente em exibição, a mensagem subsequente na pilha localizada na pasta atual será incluída na seleção.|
|**As mensagens de conversa selecionadas estão localizadas na mesma pasta que a que está atualmente em exibição**|Todas as mensagens de conversa escolhidas são incluídas na seleção.|Não aplicável. Somente o cabeçalho da conversa está disponível para seleção no modo de exibição de conversa recolhido.|
|**As mensagens de conversa selecionadas estão localizadas em pastas diferentes das que estão atualmente em exibição** |Todas as mensagens de conversa escolhidas são incluídas na seleção.|Não aplicável. Somente o cabeçalho da conversa está disponível para seleção no modo de exibição de conversa recolhido.|

## <a name="next-steps"></a>Próximas etapas

Agora que você habilitou seu suplemento para operar em várias mensagens selecionadas, você pode estender os recursos do suplemento e aprimorar ainda mais a experiência do usuário. Explore a execução de operações mais complexas usando as IDs de item das mensagens selecionadas com serviços como [Os Serviços Web do Exchange (EWS)](web-services.md) e [o Microsoft Graph](/graph/overview).

## <a name="see-also"></a>Confira também

- [Manifestos de suplementos do Outlook](manifests.md)
- [Chamar serviços Web de um suplemento do Outlook](web-services.md)
- [Visão geral do Microsoft Graph](/graph/overview)
