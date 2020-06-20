---
title: Anexar um depurador do painel de tarefas
description: Saiba como anexar um depurador do painel de tarefas
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 53cfce211241dbdf3d16e8a126e059a2f2db3f23
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810839"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Anexar um depurador do painel de tarefas

In Office 2016 on Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, Node.js, Angular, or another tool. 

Para iniciar a ferramenta **Anexar Depurador**, escolha o canto superior direito do painel de tarefas para ativar o menu **Personalidade** (conforme mostrado no círculo vermelho na imagem a seguir).   

> [!NOTE]
> - Atualmente, a única ferramenta de depurador é o [Visual Studio 2015](https://www.visualstudio.com/downloads/) com a [Atualização 3](https://msdn.microsoft.com/library/mt752379.aspx) ou posterior. Se você não tiver o Visual Studio instalado, selecionar a opção **anexar depurador** não resultará em nenhuma ação.   
> - You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".

![Captura de tela do menu Anexar Depurador](../images/attach-debugger.png)

Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image. 

![Captura de tela da caixa de diálogo Depurador JIT do Visual Studio](../images/visual-studio-debugger.png)

In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.

> [!NOTE]
> Se você não vir o menu Personalidade, é possível depurar o suplemento com o Visual Studio. Certifique-se de que o suplemento do painel tarefas esteja aberto no Office e, em seguida, siga estas etapas:
>
> 1. No Visual Studio, escolha **DEPURAR** > **Anexar ao Processo**.
> 2. Em **Processos disponíveis**, selecione*qualquer um* dos `Iexplore.exe` processos disponíveis *ou* todos os `MicrosoftEdge*.exe` processos disponíveis, dependendo [ se seu suplemento usa Internet Explorer ou Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md), e depois clique no botão **Anexar**.

Veja mais informações sobre depuração no Visual Studio, em:

-    Para iniciar e usar o Explorador do DOM no Visual Studio, confira a Dica 4 na seção [Dicas e Truques](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) da publicação [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) (Criar aplicativos atraentes para o Office usando os novos modelos de projeto) do blog.
-    Para definir pontos de interrupção, confira [Usar Pontos de Interrupção](/visualstudio/debugger/using-breakpoints?view=vs-2015).
-    Para usar o F12, confira o artigo [Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).
-   Para usar as ferramentas de desenvolvedor do Microsoft Edge, confira [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

## <a name="see-also"></a>Confira também

- [Depurar Suplementos do Office no Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Extensão do depurador de suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)