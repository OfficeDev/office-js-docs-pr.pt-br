---
title: Anexar um depurador do painel de tarefas
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 2bc3d44f1d554fb065dbb8004a744acac67ed06c
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944443"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Anexar um depurador do painel de tarefas

No Office 2016 para Windows, Build 77xx.xxxx ou posterior, é possível anexar o depurador do painel de tarefas. O recurso de anexar o depurador anexará diretamente o depurador ao processo correto do Internet Explorer. É possível anexar um depurador independentemente de você estar utilizando Yeoman Generator, Visual Studio Code, node.js, Angular ou outra ferramenta. 

Para iniciar a ferramenta **Anexar Depurador**, escolha o canto superior direito do painel de tarefas para ativar o menu **Personalidade** (conforme mostrado no círculo vermelho na imagem a seguir).   

> [!NOTE]
> - Atualmente, a única ferramenta de depurador é o [Visual Studio 2015](https://www.visualstudio.com/downloads/) com a [Atualização 3](https://msdn.microsoft.com/library/mt752379.aspx) ou posterior. Se você não instalou o Visual Studio, selecionar a opção **Anexar Depurador** não resultará em nenhuma ação.   
> - Só é possível depurar o JavaScript do lado do cliente com a ferramenta **Anexar Depurador**. Para depurar o código do lado do servidor, como com um servidor Node.js, há várias opções. Confira informações sobre como depurar com o Visual Studio Code em [Depuração do Node.js no VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Se você não estiver usando o Visual Studio Code, pesquise por "depurar Node.js" ou "depurar {nome do servidor}".

![Captura de tela do menu Anexar Depurador](../images/attach-debugger.png)

Selecione **Anexar Depurador**. Isso inicia a caixa de diálogo **Depurador Just-In-Time do Visual Studio**, conforme mostrado na imagem a seguir. 

![Captura de tela da caixa de diálogo Depurador JIT do Visual Studio](../images/visual-studio-debugger.png)

No Visual Studio, você verá os arquivos de código no **Gerenciador de Soluções**.   Você pode definir pontos de interrupção na linha de código que deseja depurar no Visual Studio.

Confira mais informações sobre depuração no Visual Studio em:

-   Para iniciar e usar o Explorador do DOM no Visual Studio, confira a Dica 4 na seção [Dicas e Truques](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) da publicação [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) (Criar aplicativos atraentes para o Office usando os novos modelos de projeto) do blog.
-   Para definir pontos de interrupção, confira [Usar Pontos de Interrupção](https://docs.microsoft.com/visualstudio/debugger/using-breakpoints?view=vs-2015).
-   Para usar o F12, confira o artigo [Usando as ferramentas de desenvolvedor F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).

## <a name="see-also"></a>Veja também

- [Criar e depurar suplementos do Office no Visual Studio](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [Publicar seu Suplemento do Office](../publish/publish.md)
