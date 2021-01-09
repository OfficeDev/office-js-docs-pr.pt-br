---
title: Executar código no seu Add-in do Office quando o documento for aberto
description: Saiba como executar código no seu add-in do Office quando o documento for aberto.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789211"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a><span data-ttu-id="10669-103">Executar código no seu Add-in do Office quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="10669-103">Run code in your Office Add-in when the document opens</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="10669-104">Você pode configurar seu Complemento do Office para carregar e executar o código assim que o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-104">You can configure your Office Add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="10669-105">Isso será útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar a interface do usuário ou executar outras tarefas antes que o complemento seja visível.</span><span class="sxs-lookup"><span data-stu-id="10669-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="10669-106">Configurar o seu complemento para carregar quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="10669-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="10669-107">O código a seguir configura o seu complemento para carregar e começar a ser executado quando o documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="10669-108">O `setStartupBehavior` método é assíncrono.</span><span class="sxs-lookup"><span data-stu-id="10669-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="10669-109">Configurar o seu add-in para nenhum comportamento de carregamento ao abrir o documento</span><span class="sxs-lookup"><span data-stu-id="10669-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="10669-110">O código a seguir configura o seu complemento para não iniciar quando o documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="10669-111">Em vez disso, ele iniciará quando o usuário a envolver de alguma forma, como escolher um botão da faixa de opções ou abrir o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="10669-111">Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="10669-112">Obter o comportamento de carregamento atual</span><span class="sxs-lookup"><span data-stu-id="10669-112">Get the current load behavior</span></span>

<span data-ttu-id="10669-113">Para determinar qual é o comportamento de inicialização atual, execute a função a seguir, que retorna um `Office.StartupBehavior` objeto.</span><span class="sxs-lookup"><span data-stu-id="10669-113">To determine what the current startup behavior is, run the following function, which returns an `Office.StartupBehavior` object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="10669-114">Como executar código quando o documento é aberto</span><span class="sxs-lookup"><span data-stu-id="10669-114">How to run code when the document opens</span></span>

<span data-ttu-id="10669-115">Quando o seu add-in estiver configurado para carregar no documento aberto, ele será executado imediatamente.</span><span class="sxs-lookup"><span data-stu-id="10669-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="10669-116">O `Office.initialize` manipulador de eventos será chamado.</span><span class="sxs-lookup"><span data-stu-id="10669-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="10669-117">Coloque o código de inicialização no `Office.initialize` manipulador de eventos ou no manipulador de `Office.onReady` eventos.</span><span class="sxs-lookup"><span data-stu-id="10669-117">Place your startup code in the `Office.initialize` or `Office.onReady` event handler.</span></span>

<span data-ttu-id="10669-118">O seguinte código de complemento do Excel mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="10669-118">The following Excel add-in code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="10669-119">Se você configurar seu complemento para carregar ao abrir o documento, esse código registrará o manipulador de eventos quando o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="10669-120">Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-120">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

<span data-ttu-id="10669-121">O código de complemento do PowerPoint a seguir mostra como registrar um manipulador de eventos para eventos de alteração de seleção do documento do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="10669-121">The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document.</span></span> <span data-ttu-id="10669-122">Se você configurar seu complemento para carregar ao abrir o documento, esse código registrará o manipulador de eventos quando o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-122">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="10669-123">Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.</span><span class="sxs-lookup"><span data-stu-id="10669-123">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="see-also"></a><span data-ttu-id="10669-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="10669-124">See also</span></span>

- [<span data-ttu-id="10669-125">Configurar o Seu Add-in do Office para usar um tempo de execução JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="10669-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="10669-126">Compartilhar dados e eventos entre funções personalizadas do Excel e tutorial do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="10669-126">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="10669-127">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="10669-127">Work with Events using the Excel JavaScript API</span></span>](../excel/excel-add-ins-events.md)
