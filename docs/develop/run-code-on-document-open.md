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
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Executar código no seu Add-in do Office quando o documento for aberto

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Você pode configurar seu Complemento do Office para carregar e executar o código assim que o documento for aberto. Isso será útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar a interface do usuário ou executar outras tarefas antes que o complemento seja visível.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurar o seu complemento para carregar quando o documento for aberto

O código a seguir configura o seu complemento para carregar e começar a ser executado quando o documento é aberto.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> O `setStartupBehavior` método é assíncrono.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurar o seu add-in para nenhum comportamento de carregamento ao abrir o documento

O código a seguir configura o seu complemento para não iniciar quando o documento é aberto. Em vez disso, ele iniciará quando o usuário a envolver de alguma forma, como escolher um botão da faixa de opções ou abrir o painel de tarefas.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obter o comportamento de carregamento atual

Para determinar qual é o comportamento de inicialização atual, execute a função a seguir, que retorna um `Office.StartupBehavior` objeto.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Como executar código quando o documento é aberto

Quando o seu add-in estiver configurado para carregar no documento aberto, ele será executado imediatamente. O `Office.initialize` manipulador de eventos será chamado. Coloque o código de inicialização no `Office.initialize` manipulador de eventos ou no manipulador de `Office.onReady` eventos.

O seguinte código de complemento do Excel mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa. Se você configurar seu complemento para carregar ao abrir o documento, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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

O código de complemento do PowerPoint a seguir mostra como registrar um manipulador de eventos para eventos de alteração de seleção do documento do PowerPoint. Se você configurar seu complemento para carregar ao abrir o documento, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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

## <a name="see-also"></a>Confira também

- [Configurar o Seu Add-in do Office para usar um tempo de execução JavaScript compartilhado](configure-your-add-in-to-use-a-shared-runtime.md)
- [Compartilhar dados e eventos entre funções personalizadas do Excel e tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Trabalhar com eventos usando a API JavaScript do Excel](../excel/excel-add-ins-events.md)
