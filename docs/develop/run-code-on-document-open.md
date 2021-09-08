---
title: Execute o código em seu Suplemento do Office quando o documento for aberto
description: Saiba como executar código em seu Office de complemento quando o documento for aberto.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938886"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>Execute o código em seu Suplemento do Office quando o documento for aberto

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Você pode configurar seu Office de usuário para carregar e executar o código assim que o documento for aberto. Isso é útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar a interface do usuário ou executar outras tarefas antes que o complemento seja visível.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurar o seu complemento para carregar quando o documento for aberto

O código a seguir configura o seu complemento para carregar e começar a ser executado quando o documento é aberto.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> O `setStartupBehavior` método é assíncrono.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurar o seu add-in para nenhum comportamento de carga ao abrir o documento

O código a seguir configura o seu complemento para não ser aberto quando o documento é aberto. Em vez disso, ele começará quando o usuário o envolver de alguma forma, como escolher um botão de faixa de opções ou abrir o painel de tarefas.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obter o comportamento de carga atual

Para determinar qual é o comportamento atual de inicialização, execute a seguinte função, que retorna um `Office.StartupBehavior` objeto.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Como executar código quando o documento é aberto

Quando o seu add-in estiver configurado para carregar no documento aberto, ele será executado imediatamente. O `Office.initialize` manipulador de eventos será chamado. Coloque seu código de inicialização no `Office.initialize` manipulador `Office.onReady` de eventos ou.

O código Excel de complemento a seguir mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa. Se você configurar seu complemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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

O código PowerPoint de complemento a seguir mostra como registrar um manipulador de eventos para eventos de alteração de seleção do PowerPoint documento. Se você configurar seu complemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes que o painel de tarefas seja aberto.

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

- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](configure-your-add-in-to-use-a-shared-runtime.md)
- [Compartilhar dados e eventos entre Excel funções personalizadas e tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Trabalhar com eventos usando a API JavaScript do Excel](../excel/excel-add-ins-events.md)
