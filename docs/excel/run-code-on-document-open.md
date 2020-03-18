---
title: Executar o código em seu suplemento do Excel quando o documento for aberto (visualização)
description: Executar o código em seu suplemento do Excel quando o documento for aberto.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: fba43fdc508245632da911acecbfa52e00847b3b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717030"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens-preview"></a><span data-ttu-id="dd731-103">Executar o código em seu suplemento do Excel quando o documento for aberto (visualização)</span><span class="sxs-lookup"><span data-stu-id="dd731-103">Run code in your Excel add-in when the document opens (preview)</span></span>

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="dd731-104">Você pode configurar seu suplemento do Excel para carregar e executar o código assim que o documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="dd731-104">You can configure your Excel add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="dd731-105">Isso é útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar interface do usuário ou executar outras tarefas antes de o suplemento ficar visível.</span><span class="sxs-lookup"><span data-stu-id="dd731-105">This is useful if you need to register event handlers, preload data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="dd731-106">Configurar seu suplemento para carregar quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="dd731-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="dd731-107">O código a seguir configura o suplemento para carregar e começar a ser executado quando o documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="dd731-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="dd731-108">O `setStartupBehavior` método é assíncrono.</span><span class="sxs-lookup"><span data-stu-id="dd731-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="dd731-109">Configurar seu suplemento para nenhum comportamento de carga no documento aberto</span><span class="sxs-lookup"><span data-stu-id="dd731-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="dd731-110">O código a seguir configura seu suplemento para não iniciar quando o documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="dd731-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="dd731-111">Em vez disso, ele será iniciado quando o usuário o envolver de alguma maneira (como a escolha de um botão de faixa de opções ou a abertura do painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="dd731-111">Instead it will start when the user engages it in some way (such as choosing a ribbon button, or opening the task pane.)</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="dd731-112">Obter o comportamento de carregamento atual</span><span class="sxs-lookup"><span data-stu-id="dd731-112">Get the current load behavior</span></span>

<span data-ttu-id="dd731-113">Para determinar qual é o comportamento de inicialização atual, execute a seguinte função, que retorna um objeto Office. StartupBehavior.</span><span class="sxs-lookup"><span data-stu-id="dd731-113">To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="dd731-114">Como executar o código quando o documento é aberto</span><span class="sxs-lookup"><span data-stu-id="dd731-114">How to run code when the document opens</span></span>

<span data-ttu-id="dd731-115">Quando o suplemento estiver configurado para carregar no documento aberto, ele será executado imediatamente.</span><span class="sxs-lookup"><span data-stu-id="dd731-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="dd731-116">O `Office.initialize` manipulador de eventos será chamado.</span><span class="sxs-lookup"><span data-stu-id="dd731-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="dd731-117">Coloque o código de inicialização no `Office.initialize` manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="dd731-117">Place your startup code in the `Office.initialize` event handler.</span></span>

<span data-ttu-id="dd731-118">O código a seguir mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="dd731-118">The following code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="dd731-119">Se você configurar seu suplemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="dd731-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="dd731-120">Você pode manipular eventos de alteração antes de abrir o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="dd731-120">You can handle change events before the task pane is opened.</span></span>


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a><span data-ttu-id="dd731-121">Também confira</span><span class="sxs-lookup"><span data-stu-id="dd731-121">See also</span></span>

- [<span data-ttu-id="dd731-122">Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="dd731-122">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)