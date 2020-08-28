---
title: Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code
description: Use o depurador de suplemento do Visual Studio Code Extension para depurar seu suplemento do Office.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 1343014fa875509fd6f0c615c3504cc9ae50dc0d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293440"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code

O Microsoft Office Add-in Debugger Extension para o Visual Studio Code permite que você depure seu suplemento do Office em tempo de execução de borda.

Este modo de depuração é dinâmico, permitindo que você defina pontos de interrupção enquanto o código está sendo executado. Você pode ver alterações no seu código imediatamente enquanto o depurador é anexado, tudo sem perder a sessão de depuração. Suas alterações de código também persistim, portanto, você pode ver os resultados de várias alterações em seu código. A imagem a seguir mostra essa extensão em ação.

![Extensão do depurador de suplementos do Office depuração de uma seção de suplementos do Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como um administrador)
- [Node.js (versão 10 +)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Estas instruções pressupõem que você tenha experiência em usar a linha de comando, entenda o JavaScript básico e criou um projeto de suplemento do Office antes de usar o gerador do Office Yo. Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este [tutorial de suplemento do Office Excel](../tutorials/excel-tutorial.md).

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

1. Se você precisar criar um projeto de suplemento, [use o gerador de Yo do Office para criar um](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator). Siga os prompts dentro da linha de comando para configurar seu projeto. Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.

> [!NOTE]
> Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.

2. Abra um prompt de comando como administrador.
   ![Opções de prompt de comando, incluindo "executar como administrador" no Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Navegue até o diretório do projeto.

4. Execute o seguinte comando para abrir seu projeto no Visual Studio Code como um administrador.

```command&nbsp;line
code .
```

Depois que o Visual Studio code estiver aberto, navegue manualmente para a pasta do projeto.

> [!TIP]
> Para abrir o Visual Studio Code como um administrador, selecione a opção **Executar como administrador** ao abrir o código do Visual Studio após procurá-lo no Windows.

5. No VS Code, selecione **Ctrl + Shift + X** para abrir a barra de extensões. Procure a extensão "depurador de suplementos do Microsoft Office" e instale-a.

6. Na pasta. vscode do projeto, abra o **launch.jsem** arquivo. Adicione o seguinte código à `configurations` seção:

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. Na seção de JSON que você acabou de copiar, encontre a seção "URL". Nesta URL, será necessário substituir o texto de HOST em maiúsculas pelo aplicativo que está hospedando o suplemento do Office. Por exemplo, se o suplemento do Office for Excel, seu valor de URL será " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0".

8. Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto. Execute o comando `npm start` para iniciar o servidor de desenvolvimento. Quando o suplemento for carregado no cliente do Office, abra o painel de tarefas.

9. Retorne ao Visual Studio Code e escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.

10. Nas opções de depuração, escolha **anexar a suplementos do Office**. Selecione **F5** ou escolha **debug-> iniciar a depuração** no menu para iniciar a depuração.

11. Defina um ponto de interrupção no arquivo de painel de tarefas do projeto. É possível definir pontos de interrupção no VS Code focalizando ao lado de uma linha de código e selecionando o círculo vermelho que aparece.

![Um círculo vermelho aparece em uma linha de código no VS Code](../images/set-breakpoint.jpg)

12. Execute o suplemento. Você verá que os pontos de interrupção foram atingidos e pode inspecionar as variáveis locais.

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)

* [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
