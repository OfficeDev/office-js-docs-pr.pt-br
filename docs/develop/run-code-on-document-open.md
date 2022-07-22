---
title: Execute o código em seu Suplemento do Office quando o documento for aberto
description: Saiba como executar código em seu suplemento do Office quando o documento for aberto.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a1c3277a349dc4054da5f089c62331296590021
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958436"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Execute o código em seu Suplemento do Office quando o documento for aberto

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Você pode configurar seu Suplemento do Office para carregar e executar código assim que o documento for aberto. Isso será útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar a interface do usuário ou executar outras tarefas antes que o suplemento fique visível.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurar o suplemento para carregar quando o documento for aberto

O código a seguir configura o suplemento para carregar e começar a ser executado quando o documento é aberto.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> O `setStartupBehavior` método é assíncrono.

## <a name="place-startup-code-in-officeinitialize"></a>Colocar o código de inicialização em Office.initialize

Quando o suplemento estiver configurado para carregar no documento aberto, ele será executado imediatamente. O `Office.initialize` manipulador de eventos será chamado. Coloque o código de inicialização no manipulador `Office.initialize` de eventos `Office.onReady` ou no manipulador de eventos.

O código de suplemento do Excel a seguir mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa. Se você configurar o suplemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

O código de suplemento do PowerPoint a seguir mostra como registrar um manipulador de eventos para eventos de alteração de seleção do documento do PowerPoint. Se você configurar o suplemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurar seu suplemento para não ter nenhum comportamento de carregamento ao abrir o documento

O código a seguir configura o suplemento para não iniciar quando o documento é aberto. Em vez disso, ele será iniciado quando o usuário o envolver de alguma forma, como escolher um botão da faixa de opções ou abrir o painel de tarefas.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obter o comportamento de carga atual

Para determinar qual é o comportamento de inicialização atual, execute o método a seguir, que retorna um `Office.StartupBehavior` objeto.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>Confira também

- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](configure-your-add-in-to-use-a-shared-runtime.md)
- [Compartilhar dados e eventos entre funções personalizadas do Excel e o tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Trabalhar com eventos usando a API JavaScript do Excel](../excel/excel-add-ins-events.md)
