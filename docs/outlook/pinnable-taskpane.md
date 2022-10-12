---
title: Implementar um painel de tarefas fixável em um suplemento do Outlook
description: A forma do painel de tarefas da experiência de usuário dos comandos do suplemento abre um painel de tarefas vertical à direita de uma solicitação de reunião ou de uma mensagem aberta, permitindo ao suplemento fornecer à interface do usuário interações mais detalhadas.
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 834d43a6046ddaa63a7c8899cfd5b07d0ea80ef6
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541119"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Implementar um painel de tarefas fixável no Outlook

The [task pane](add-in-commands-for-outlook.md#launch-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Embora o recurso de painéis de tarefas fixáveis tenha sido introduzido no conjunto de requisitos [1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5), ele está disponível apenas para assinantes do Microsoft 365 usando o seguinte:
>
> - O Outlook 2016 ou posterior para Windows (build 7668.2000 ou posterior para os usuários dos Canais Atuais ou do Office Insider, build 7900.xxxx ou posterior para os usuários em canais Adiados)
> - Outlook 2016 ou posterior no Mac (versão 16.13.503 ou posterior)
> - Outlook na Web moderno

> [!IMPORTANT]
> Os painéis de tarefas fixáveis não estão disponíveis para o seguinte:
>
> - Compromissos/Reuniões
> - Outlook.com

> [!TIP]
> Se você planeja publicar [](../publish/publish.md) seu suplemento do Outlook no [AppSource](https://appsource.microsoft.com) e ele está configurado para um painel de tarefas fixável, para passar na validação do [AppSource](/legal/marketplace/certification-policies), o conteúdo do suplemento não deve ser estático e deve exibir claramente os dados relacionados à mensagem que está aberta ou selecionada na caixa de correio.

## <a name="support-task-pane-pinning"></a>Suporte para fixação do painel de tarefas

A primeira etapa consiste em adicionar o suporte para fixação no manifesto do suplemento. A marcação varia dependendo do tipo de manifesto.

# <a name="xml-manifest"></a>[Manifesto XML](#tab/xmlmanifest)

Adicione o [elemento SupportsPinning](/javascript/api/manifest/action#supportspinning) ao elemento **\<Action\>** que descreve o botão do painel de tarefas. Apresentamos um exemplo a seguir.

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

**\<SupportsPinning\>** O elemento é definido no esquema VersionOverrides v1.1, portanto, você precisará incluir um elemento [VersionOverrides](/javascript/api/manifest/versionoverrides) para v1.0 e v1.1.

# <a name="teams-manifest-developer-preview"></a>[Manifesto do Teams (versão prévia do desenvolvedor)](#tab/jsonmanifest)

Adicione uma propriedade "fixável", `true`definida como , para o objeto na matriz "actions" que define o botão ou item de menu que abre o painel de tarefas. Apresentamos um exemplo a seguir.

```json
"actions": [
    {
        "id": "OpenTaskPane",
        "type": "openPage",
        "view": "TaskPaneView",
        "displayName": "OpenTaskPane",
        "pinnable": true
    }
]
```

---

Para ver um exemplo completo, confira o controle `msgReadOpenPaneButton` na [amostra de manifesto command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Atualizações de tratamento da interface do usuário com base na mensagem atualmente selecionada

Para atualizar a interface do usuário ou as variáveis internas do painel de tarefas com base no item atual, você deve registrar um manipulador de eventos para receber notificações das alterações.

### <a name="implement-the-event-handler"></a>Implementar o manipulador de eventos

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> A implementação de manipuladores de eventos para um evento ItemChanged deve verificar se o Office.content.mailbox.item é nulo.
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a>Registrar o manipulador de eventos

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a>Confira também

Para obter um exemplo de suplemento que implementa um painel de tarefas fixável, confira [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) no GitHub.
