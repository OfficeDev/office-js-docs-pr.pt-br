---
title: 'Mostre ou oculte o painel de tarefas de seu Suplemento do Office '
description: Saiba como ocultar ou mostrar programaticamente a interface do usuário de um suplemento enquanto ele é executado continuamente.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 95f8c716bf1a0331fe47bc74e5aad49c17b65437
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660127"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Mostre ou oculte o painel de tarefas de seu Suplemento do Office 

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Você pode mostrar o painel de tarefas do suplemento do Office chamando a `Office.addin.showAsTaskpane()` função.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

O código anterior pressupõe um cenário em que há uma planilha do Excel chamada **CurrentQuarterSales**. O suplemento tornará o painel de tarefas visível sempre que essa planilha for ativada. O método `onCurrentQuarter` é um manipulador para o [evento Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) que foi registrado para a planilha.

Você também pode ocultar o painel de tarefas chamando a `Office.addin.hide()` função.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

O código anterior é um manipulador registrado para o [evento Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) .

## <a name="additional-details-on-showing-the-task-pane"></a>Detalhes adicionais sobre como mostrar o painel de tarefas

Quando você chamar `Office.addin.showAsTaskpane()`, o Office exibirá em um painel de tarefas o arquivo atribuído como o valor da ID do recurso (`resid`) do painel de tarefas. Esse `resid` valor pode ser atribuído ou alterado abrindo seu arquivo **manifest.xml** e localizando **\<SourceLocation\>** dentro do `<Action xsi:type="ShowTaskpane">` elemento.
(Consulte [Configurar seu Suplemento do Office para usar um runtime compartilhado](configure-your-add-in-to-use-a-shared-runtime.md) para obter detalhes adicionais.)

Como `Office.addin.showAsTaskpane()` é um método assíncrono, seu código continuará em execução até que a função seja concluída. Aguarde essa conclusão com a palavra-chave `await` ou um `then()` método, dependendo da sintaxe JavaScript que você está usando.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Configurar o suplemento para usar o runtime compartilhado

Para usar os `showAsTaskpane()` métodos `hide()` e os métodos, o suplemento deve usar o runtime compartilhado. Para obter mais informações, [consulte Configurar seu Suplemento do Office para usar um runtime compartilhado](configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="preservation-of-state-and-event-listeners"></a>Preservação de ouvintes de estado e eventos

Os `hide()` métodos `showAsTaskpane()` e os métodos alteram *apenas a visibilidade* do painel de tarefas. Eles não descarregam nem recarregam (ou reinicializam seu estado).

Considere o seguinte cenário: um painel de tarefas é projetado com guias. A **guia** Página Inicial é aberta quando o suplemento é iniciado pela primeira vez. Suponha que um usuário abra **a guia Configurações** e, posteriormente, o código no painel de tarefas chame `hide()` em resposta a algum evento. Ainda mais tarde, o código `showAsTaskpane()` chama em resposta a outro evento. O painel de tarefas reaparecerá e a **guia Configurações** ainda está selecionada.

![Uma captura de tela do painel de tarefas que tem quatro guias rotuladas como Página Inicial, Configurações, Favoritos e Contas.](../images/TaskpaneWithTabs.png)

Além disso, todos os ouvintes de eventos registrados no painel de tarefas continuam a ser executados mesmo quando o painel de tarefas está oculto.

Considere o seguinte cenário: o painel de tarefas tem um manipulador registrado para o Excel `Worksheet.onActivated` `Worksheet.onDeactivated` e eventos para uma planilha chamada **Sheet1**. O manipulador ativado faz com que um ponto verde apareça no painel de tarefas. O manipulador desativado torna o ponto vermelho (que é seu estado padrão). Suponha que esse código chame quando `hide()` **Sheet1** não estiver ativado e o ponto estiver vermelho. Enquanto o painel de tarefas está oculto, **Sheet1** é ativado. Chamadas de código posteriores `showAsTaskpane()` em resposta a algum evento. Quando o painel de tarefas é aberto, o ponto fica verde porque os ouvintes e manipuladores de eventos foram executados mesmo que o painel de tarefas estivesse oculto.

## <a name="handle-the-visibility-changed-event"></a>Manipular o evento de visibilidade alterada

Quando seu código altera a visibilidade do painel de tarefas com `showAsTaskpane()` ou `hide()`, o Office dispara o `VisibilityModeChanged` evento. Pode ser útil lidar com esse evento. Por exemplo, suponha que o painel de tarefas exiba uma lista de todas as planilhas em uma pasta de trabalho. Se uma nova planilha for adicionada enquanto o painel de tarefas estiver oculto, tornar o painel de tarefas visível não adicionaria, por si só, o novo nome da planilha à lista. Mas seu código pode `VisibilityModeChanged` responder ao evento para recarregar [a propriedade Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) de todas as planilhas na coleção [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) , conforme mostrado no código de exemplo abaixo.

Para registrar um manipulador para o evento, você não usa um método "adicionar manipulador" como faria na maioria dos contextos javaScript do Office. Em vez disso, há uma função especial para a qual você passa o manipulador: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). Apresentamos um exemplo a seguir. Observe que a propriedade `args.visibilityMode` é do tipo [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

A função retorna outra função que *desregisia* o manipulador. Aqui está um exemplo simples, mas não robusto.

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

O `onVisibilityModeChanged` método é assíncrono e retorna uma promessa, o que significa que seu código precisa aguardar o cumprimento da promessa antes de chamar o manipulador **de cancelamento** de registro.

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

A função de cancelamento de registro também é assíncrona e retorna uma promessa. Portanto, se você tiver um código que não deve ser executado até que o cancelamento do registro seja concluído, aguarde a promessa retornada pela função de cancelamento de registro.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>Confira também

- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](configure-your-add-in-to-use-a-shared-runtime.md)
- [Execute o código em seu Suplemento do Office quando o documento for aberto](run-code-on-document-open.md)
