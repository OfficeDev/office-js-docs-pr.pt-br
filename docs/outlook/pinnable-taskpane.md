---
title: Implementar um painel de tarefas fixável em um suplemento do Outlook
description: A forma do painel de tarefas da experiência de usuário dos comandos do suplemento abre um painel de tarefas vertical à direita de uma solicitação de reunião ou de uma mensagem aberta, permitindo ao suplemento fornecer à interface do usuário interações mais detalhadas.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 57a17a90fe565adb3ffb9d23e3b169bc83be2735
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937956"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Implementar um painel de tarefas fixável no Outlook

A forma do [painel de tarefas](add-in-commands-for-outlook.md#launching-a-task-pane) da experiência de usuário dos comandos do suplemento abre um painel de tarefas vertical à direita de uma solicitação de reunião ou de uma mensagem aberta, permitindo ao suplemento fornecer a interface do usuário a fim de obter interações mais detalhadas (preenchimento de vários campos etc.). Esse painel de tarefas pode ser exibido no painel de leitura durante a exibição de uma lista de mensagens, permitindo o processamento rápido de uma mensagem.

No entanto, se o usuário abrir um painel de tarefas do suplemento em uma mensagem no painel de leitura e selecionar uma nova mensagem, o painel de tarefas será fechado automaticamente, por padrão. Para um suplemento bastante usado, o usuário pode optar por manter esse painel aberto, eliminando a necessidade de reativar o suplemento em cada mensagem. Com os painéis de tarefas fixáveis, o suplemento pode fornecer essa opção aos usuários.

> [!NOTE]
> Embora o recurso de painéis de tarefas pinnable tenha sido introduzido no conjunto de requisitos [1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), ele está disponível apenas para assinantes Microsoft 365 usando o seguinte:
>
> - Outlook 2016 ou posterior no Windows (build 7668.2000 ou posterior para usuários nos canais Insider Current ou Office, build 7900.xxxx ou posterior para usuários em canais adiados)
> - Outlook 2016 ou posterior no Mac (versão 16.13.503 ou posterior)
> - Outlook na Web moderno

> [!IMPORTANT]
> Os painéis de tarefas pinnable não estão disponíveis para o seguinte:
>
> - Compromissos/Reuniões
> - Outlook.com

## <a name="support-task-pane-pinning"></a>Suporte para fixação do painel de tarefas

A primeira etapa consiste em adicionar o suporte para fixação no [manifesto](manifests.md) do suplemento. Para fazer isso, adicione o elemento [SupportsPinning](../reference/manifest/action.md#supportspinning) ao elemento `Action`, que descreve o botão do painel de tarefas.

O elemento `SupportsPinning` é definido no esquema VersionOverrides v1.1. Portanto, você deve incluir um elemento [VersionOverrides](../reference/manifest/versionoverrides.md) nas versões 1.0 e 1.1.

> [!NOTE]
> Se você pretende [publicar](../publish/publish.md) seu suplemento do Outlook em [AppSource](https://appsource.microsoft.com), quando usar o elemento **SupportsPinning**, para passar a [validação da AppSource](/legal/marketplace/certification-policies), o conteúdo do seu suplemento não deve ser estático e deve exibir claramente os dados relacionados à mensagem que está aberta ou selecionada na caixa de correio.

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

Para ver um exemplo completo, confira o controle `msgReadOpenPaneButton` na [amostra de manifesto command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Atualizações de tratamento da interface do usuário com base na mensagem atualmente selecionada

Para atualizar a interface do usuário ou as variáveis internas do painel de tarefas com base no item atual, você deve registrar um manipulador de eventos para receber notificações das alterações.

### <a name="implement-the-event-handler"></a>Implementar o manipulador de eventos

O manipulador de eventos deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` desse objeto será definida como `Office.EventType.ItemChanged`. Ao chamar o evento, o objeto `Office.context.mailbox.item` já estará atualizado para refletir o item atualmente selecionado.

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

Use o método [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para registrar o manipulador de eventos para o evento `Office.EventType.ItemChanged`. Você deve fazer isso na função `Office.initialize` do painel de tarefas.

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
