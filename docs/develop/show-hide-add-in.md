---
title: Mostrar ou ocultar o painel de tarefas do seu Add-in do Office
description: Saiba como ocultar ou mostrar programaticamente a interface do usuário de um complemento enquanto ele é executado continuamente.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789212"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Mostrar ou ocultar o painel de tarefas do seu Add-in do Office

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Você pode mostrar o painel de tarefas do seu Complemento do Office chamando a `Office.addin.showAsTaskpane()` função.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

O código anterior assume um cenário em que há uma planilha do Excel chamada **CurrentQuarterSales**. O complemento torna o painel de tarefas visível sempre que essa planilha é ativada. O método `onCurrentQuarter` é um manipulador para o evento [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) que foi registrado para a planilha.

Você também pode ocultar o painel de tarefas chamando a `Office.addin.hide()` função.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

O código anterior é um manipulador que está registrado para o [evento Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)

## <a name="additional-details-on-showing-the-task-pane"></a>Detalhes adicionais sobre como mostrar o painel de tarefas

Ao chamar, o Office exibirá em um painel de tarefas o arquivo atribuído como o valor da ID do recurso `Office.addin.showAsTaskpane()` ( ) do painel de `resid` tarefas. Esse valor pode ser atribuído ou alterado abrindo seu `resid` **arquivomanifest.xml** e localizando dentro `<SourceLocation>` do `<Action xsi:type="ShowTaskpane">` elemento.
(Confira [Configurar o seu Add-in do Office para usar um tempo de execução compartilhado](configure-your-add-in-to-use-a-shared-runtime.md) para obter detalhes adicionais.)

Como `Office.addin.showAsTaskpane()` é um método assíncrono, seu código continuará em execução até que a função seja concluída. Aguarde essa conclusão com a `await` palavra-chave ou um `then()` método, dependendo da sintaxe JavaScript que você está usando.

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>Configurar seu complemento para usar o tempo de execução compartilhado

Para usar os `showAsTaskpane()` métodos `hide()` e os métodos, o seu complemento deve usar o tempo de execução compartilhado. Para saber mais, confira [Configurar o seu Add-in do Office para usar um tempo de execução compartilhado.](configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="preservation-of-state-and-event-listeners"></a>Preservação de ouvintes de estado e eventos

The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane. Eles não descarregam ou recarregam (ou reinicializam seu estado).

Considere o seguinte cenário: um painel de tarefas foi projetado com guias. A **guia** Início é aberta quando o complemento é lançado pela primeira vez. Suponha que um usuário abra a guia **Configurações** e, mais tarde, o código no painel de tarefas chama `hide()` em resposta a algum evento. Ainda mais tarde, o `showAsTaskpane()` código chama em resposta a outro evento. O painel de tarefas reaparecerá e **a guia Configurações** ainda está selecionada.

![Uma captura de tela do painel de tarefas que tem quatro guias rotuladas como Página Inicial, Configurações, Favoritos e Contas.](../images/TaskpaneWithTabs.png)

Além disso, os ouvintes de eventos registrados no painel de tarefas continuam a ser executados mesmo quando o painel de tarefas está oculto.

Considere o seguinte cenário: O painel de tarefas tem um manipulador registrado para o Excel e eventos `Worksheet.onActivated` para uma planilha chamada `Worksheet.onDeactivated` **Sheet1**. O manipulador ativado faz com que um ponto verde apareça no painel de tarefas. O manipulador desativado ativa o ponto vermelho (que é seu estado padrão). Suponha então que o código `hide()` chama **quando Sheet1** não está ativado e o ponto é vermelho. Enquanto o painel de tarefas está oculto, **Sheet1** é ativado. Chamadas de código `showAsTaskpane()` posteriores em resposta a algum evento. Quando o painel de tarefas é aberto, o ponto fica verde porque os ouvintes e manipuladores de eventos foram embora o painel de tarefas tenha sido ocultado.

## <a name="handle-the-visibility-changed-event"></a>Manipular o evento de visibilidade alterada

Quando seu código altera a visibilidade do painel de tarefas com `showAsTaskpane()` `hide()` ou, o Office aciona o `VisibilityModeChanged` evento. Pode ser útil para manipular esse evento. Por exemplo, suponha que o painel de tarefas exibe uma lista de todas as planilhas em uma planilha. Se uma nova planilha for adicionada enquanto o painel de tarefas estiver oculto, tornar o painel de tarefas visível não adicionaria, por si só, o novo nome da planilha à lista. Mas seu código pode responder ao evento para recarregar a propriedade Worksheet.name de todas as planilhas na coleção `VisibilityModeChanged` [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) conforme mostrado no código de exemplo abaixo. [](/javascript/api/excel/excel.worksheet#name)

Para registrar um manipulador para o evento, você não usa um método "adicionar manipulador" como faria na maioria dos contextos javaScript do Office. Em vez disso, há uma função especial para a qual você passa seu manipulador: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-). Apresentamos um exemplo a seguir. Observe que a `args.visibilityMode` propriedade é do tipo [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

A função retorna outra função que *desregula o* manipulador. Veja um exemplo simples, mas não robusto:

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

O método é assíncrono e retorna uma promessa, o que significa que seu código precisa aguardar o cumprimento da promessa antes de chamar o manipulador `onVisibilityModeChanged` **de registro.**

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

A função de desregister também é assíncrona e retorna uma promessa. Portanto, se você tiver um código que não deve ser executado até que o desregistramento seja concluído, aguarde a promessa retornada pela função de desregister.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>Confira também

- [Configurar o Seu Add-in do Office para usar um tempo de execução JavaScript compartilhado](configure-your-add-in-to-use-a-shared-runtime.md)
- [Executar código no seu Add-in do Office quando o documento for aberto](run-code-on-document-open.md)
