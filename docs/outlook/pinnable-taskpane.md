---
title: Implementar um painel de tarefas fixável em um suplemento do Outlook
description: A forma do painel de tarefas da experiência de usuário dos comandos do suplemento abre um painel de tarefas vertical à direita de uma solicitação de reunião ou de uma mensagem aberta, permitindo ao suplemento fornecer à interface do usuário interações mais detalhadas.
ms.date: 11/18/2019
localization_priority: Normal
ms.openlocfilehash: 94c136a74dfddac1af663aea06c3c6ca27f22dcd
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165706"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="3fe3f-103">Implementar um painel de tarefas fixável no Outlook</span><span class="sxs-lookup"><span data-stu-id="3fe3f-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="3fe3f-p101">A forma do [painel de tarefas](add-in-commands-for-outlook.md#launching-a-task-pane) da experiência de usuário dos comandos do suplemento abre um painel de tarefas vertical à direita de uma solicitação de reunião ou de uma mensagem aberta, permitindo ao suplemento fornecer a interface do usuário a fim de obter interações mais detalhadas (preenchimento de vários campos etc.). Esse painel de tarefas pode ser exibido no painel de leitura durante a exibição de uma lista de mensagens, permitindo o processamento rápido de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-p101">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="3fe3f-p102">No entanto, se o usuário abrir um painel de tarefas do suplemento em uma mensagem no painel de leitura e selecionar uma nova mensagem, o painel de tarefas será fechado automaticamente, por padrão. Para um suplemento bastante usado, o usuário pode optar por manter esse painel aberto, eliminando a necessidade de reativar o suplemento em cada mensagem. Com os painéis de tarefas fixáveis, o suplemento pode fornecer essa opção aos usuários.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-p102">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="3fe3f-109">Painéis de tarefa fixáveis estão atualmente disponíveis para assinantes do Office 365 usando o Outlook 2016 ou posterior no Windows (build 7668.2000 ou posterior para usuários nos canais atuais ou do Office Insider e build 7900.xxxx ou posterior para usuários em canais adiados), Outlook 2016 ou posterior para Mac (versão 16.13.503 ou posterior) e Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-109">Pinnable task panes are currently available to Office 365 subscribers using Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels), Outlook 2016 or later on Mac (version 16.13.503 or later), and Outlook on the web.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3fe3f-110">Painéis de tarefas fixáveis não estão disponíveis para o seguinte.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-110">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="3fe3f-111">Compromissos/Reuniões</span><span class="sxs-lookup"><span data-stu-id="3fe3f-111">Appointments/Meetings</span></span>
> - <span data-ttu-id="3fe3f-112">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="3fe3f-112">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="3fe3f-113">Suporte para fixação do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3fe3f-113">Support task pane pinning</span></span>

<span data-ttu-id="3fe3f-p103">A primeira etapa consiste em adicionar o suporte para fixação no [manifesto](manifests.md) do suplemento. Para fazer isso, adicione o elemento [SupportsPinning](../reference/manifest/action.md#supportspinning) ao elemento `Action`, que descreve o botão do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-p103">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="3fe3f-116">O elemento `SupportsPinning` é definido no esquema VersionOverrides v1.1. Portanto, você deve incluir um elemento [VersionOverrides](../reference/manifest/versionoverrides.md) nas versões 1.0 e 1.1.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-116">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="3fe3f-117">Se você pretende [publicar](../publish/publish.md) seu suplemento do Outlook em [AppSource](https://appsource.microsoft.com), quando usar o elemento **SupportsPinning**, para passar a [validação da AppSource](/office/dev/store/validation-policies), o conteúdo do seu suplemento não deve ser estático e deve exibir claramente os dados relacionados à mensagem que está aberta ou selecionada na caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-117">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/office/dev/store/validation-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

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

<span data-ttu-id="3fe3f-118">Para ver um exemplo completo, confira o controle `msgReadOpenPaneButton` na [amostra de manifesto command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="3fe3f-118">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="3fe3f-119">Atualizações de tratamento da interface do usuário com base na mensagem atualmente selecionada</span><span class="sxs-lookup"><span data-stu-id="3fe3f-119">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="3fe3f-120">Para atualizar a interface do usuário ou as variáveis internas do painel de tarefas com base no item atual, você deve registrar um manipulador de eventos para receber notificações das alterações.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-120">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="3fe3f-121">Implementar o manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="3fe3f-121">Implement the event handler</span></span>

<span data-ttu-id="3fe3f-p104">O manipulador de eventos deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` desse objeto será definida como `Office.EventType.ItemChanged`. Ao chamar o evento, o objeto `Office.context.mailbox.item` já estará atualizado para refletir o item atualmente selecionado.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-p104">The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="3fe3f-125">A implementação de manipuladores de eventos para um evento ItemChanged deve verificar se o Office.content.mailbox.item é nulo.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-125">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="3fe3f-126">Registrar o manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="3fe3f-126">Register the event handler</span></span>

<span data-ttu-id="3fe3f-p105">Use o método [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para registrar o manipulador de eventos para o evento `Office.EventType.ItemChanged`. Você deve fazer isso na função `Office.initialize` do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-p105">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="3fe3f-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="3fe3f-129">See also</span></span>

<span data-ttu-id="3fe3f-130">Para obter um exemplo de suplemento que implementa um painel de tarefas fixável, confira [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="3fe3f-130">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>