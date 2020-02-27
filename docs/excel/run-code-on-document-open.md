---
title: Executar o código em seu suplemento do Excel quando o documento for aberto (visualização)
description: Executar o código em seu suplemento do Excel quando o documento for aberto.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 5b8c646a1154540244b1f5e0ac47ad8eaec1801f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284108"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens-preview"></a>Executar o código em seu suplemento do Excel quando o documento for aberto (visualização)

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

Você pode configurar seu suplemento do Excel para carregar e executar o código assim que o documento é aberto. Isso é útil se você precisar registrar manipuladores de eventos, pré-carregar dados para o painel de tarefas, sincronizar interface do usuário ou executar outras tarefas antes de o suplemento ficar visível.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>Configurar seu suplemento para carregar quando o documento for aberto

O código a seguir configura o suplemento para carregar e começar a ser executado quando o documento é aberto.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> O `setStartupBehavior` método é assíncrono.

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>Configurar seu suplemento para nenhum comportamento de carga no documento aberto

O código a seguir configura seu suplemento para não iniciar quando o documento é aberto. Em vez disso, ele será iniciado quando o usuário o envolver de alguma maneira (como a escolha de um botão de faixa de opções ou a abertura do painel de tarefas).

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>Obter o comportamento de carregamento atual

Para determinar qual é o comportamento de inicialização atual, execute a seguinte função, que retorna um objeto Office. StartupBehavior.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>Como executar o código quando o documento é aberto

Quando o suplemento estiver configurado para carregar no documento aberto, ele será executado imediatamente. O `Office.initialize` manipulador de eventos será chamado. Coloque o código de inicialização no `Office.initialize` manipulador de eventos.

O código a seguir mostra como registrar um manipulador de eventos para eventos de alteração da planilha ativa. Se você configurar seu suplemento para carregar no documento aberto, esse código registrará o manipulador de eventos quando o documento for aberto. Você pode manipular eventos de alteração antes de abrir o painel de tarefas.


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

## <a name="see-also"></a>Confira também

- [Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)