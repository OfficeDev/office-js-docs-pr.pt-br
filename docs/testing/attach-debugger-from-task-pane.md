---
title: Anexar um depurador do painel de tarefas
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: f3d5b5596a69eed3404a0e37b7764c1e74d445c1
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639977"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Anexar um depurador do painel de tarefas

No Office 2016 para Windows, Build 77xx.xxxx ou posterior, é possível anexar o depurador do painel de tarefas. O recurso para anexar o depurador anexará diretamente o depurador ao processo correto do Internet Explorer. É possível anexar um depurador independentemente de você estar utilizando Yeoman Generator, Visual Studio Code, node.js, Angular ou outra ferramenta. 

Para iniciar a ferramenta **Anexar Depurador** , escolha o canto superior direito do painel de tarefas para ativar o menu **Personalidade** (conforme mostrado no círculo vermelho na imagem a seguir).   

> [!NOTE]
> - Atualmente, a única ferramenta suportada de depurador é o [Visual Studio 2015](https://www.visualstudio.com/downloads/) com a [Atualização 3](https://msdn.microsoft.com/library/mt752379.aspx) ou posterior. Se você não instalou o Visual Studio, selecionar a opção **Anexar Depurador** não resultará em nenhuma ação.   
> - Só é possível depurar JavaScript do lado do cliente com a ferramenta de **Anexar depurador** . Para depurar código do lado do servidor, como com um servidor Node.js, você tem várias opções. Para obter informações sobre como depurar com o código do Visual Studio, confira [Depuração de Node.js no código VS](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Se você não estiver usando o código do Visual Studio, pesquise "debug Node. js" ou "debug {nome-do-servidor}".

![Captura de tela do menu Anexar Depurador](../images/attach-debugger.png)

Selecione **Anexar Depurador**. Isso inicia a caixa de diálogo **Depurador Just-In-Time do Visual Studio**, conforme mostrado na imagem a seguir. 

![Captura de tela da caixa de diálogo Depurador JIT do Visual Studio](../images/visual-studio-debugger.png)

No Visual Studio, você verá os arquivos de código no **Solution Explorer**.   Você pode definir pontos de interrupção para a linha de código que você deseja depurar no Visual Studio.

> [!NOTE]
> Se você não vir o menu de Personalidade, você pode depurar seu suplemento usando o Visual Studio. Certifique-se de que seu suplemento de painel tarefa esteja aberto no Office e, em seguida, siga estas etapas:

> 1. No Visual Studio, escolha **DEPURAR** > **Anexar ao Processo**.
> 2. Na caixa de diálogo **Anexar ao Processo**, escolha todos os processos Iexplore.exe disponíveis e, em seguida, selecione o botão **Anexar**.

Confira mais informações sobre depuração no Visual Studio em:

-   Para iniciar e usar o Explorador do DOM no Visual Studio, confira a Dica 4 na seção [Dicas e Truques](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) da publicação [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) (Criar aplicativos atraentes para o Office usando os novos modelos de projeto) do blog.
-   Para definir pontos de interrupção, confira [Usar Pontos de Interrupção](https://docs.microsoft.com/visualstudio/debugger/using-breakpoints?view=vs-2015).
-   Para usar o F12, confira o artigo [Usando as ferramentas de desenvolvedor F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).

## <a name="see-also"></a>Confira também

- [Criar e depurar suplementos do Office no Visual Studio](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [Publicar seu suplemento do Office](../publish/publish.md)
