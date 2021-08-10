---
title: Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code
description: Use o Visual Studio Code de Microsoft Office Depurador de Complementos para depurar seu Office Add-in.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: d027e5937fa3a58623ce9e798fc683e5459e73b8b72606c0a006e465c9c1360c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088462"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code

A extensão de depurador de Microsoft Office do Visual Studio Code permite depurar seu Office Add-in no Microsoft Edge com o tempo de execução do WebView (EdgeHTML) original. Para obter instruções sobre a depuração em Microsoft Edge WebView2 (Chromium baseado em Chromium), [consulte este artigo](./debug-desktop-using-edge-chromium.md)

Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução. Você pode ver alterações em seu código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração. As alterações de código também persistem, para que você possa ver os resultados de várias alterações no código. A imagem a seguir mostra essa extensão em ação.

![Office Extensão de depurador de add-in depurando uma seção de Excel de complementos.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Estas instruções pressuem que você tenha experiência usando a linha de comando, entenda JavaScript básico e tenha criado um projeto de Office de Office antes de usar o gerador Yo Office. Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este Excel Office [tutorial de complemento.](../tutorials/excel-tutorial.md)

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

1. Se você precisar criar um projeto de add-in, [use o gerador Yo Office para criar um](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator). Siga os prompts dentro da linha de comando para configurar seu projeto. Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.

    > [!NOTE]
    > Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.

1. Abra um prompt de comando como administrador.
   ![Opções de prompt de comando, incluindo "executar como administrador" no Windows 10.](../images/run-as-administrator-vs-code.jpg)

1. Navegue até o diretório do projeto.

1. Execute o seguinte comando para abrir seu projeto Visual Studio Code como administrador.

    ```command&nbsp;line
    code .
    ```

  Depois Visual Studio Code abrir, navegue manualmente até a pasta do projeto.

  > [!TIP]
  > Para abrir Visual Studio Code como administrador, selecione  a opção executar como administrador ao abrir Visual Studio Code depois de procurá-lo no Windows.

1. No VS Code, selecione **Ctrl+Shift+X** para abrir a barra Extensões. Procure a extensão "Microsoft Office Depurador de Complementos" e instale-a.

1. Na pasta .vscode do seu projeto, abra o arquivo **launch.json**. Adicione o código a seguir à `configurations` seção.

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

1. Na seção JSON que você acabou de copiar, encontre a seção "url". Nesta URL, você precisará substituir o texto HOST maiúscula pelo aplicativo que está hospedando seu Office Add-in. Por exemplo, se o Office de Office for para Excel, o valor da URL será " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16,01$en-US$ \$ \$ \$ 0".

1. Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto. Execute o comando `npm start` para iniciar o servidor de dev. Quando o seu complemento for carregado no cliente Office, abra o painel de tarefas.

1. Retorne ao Visual Studio Code e escolha **Exibir > Depurar** ou insira **CTRL + SHIFT + D** para alternar para o exibição de depuração.

1. Nas opções Depurar, escolha **Anexar a Office Depuração.** Selecione **F5** ou escolha **Debug -> Iniciar Depuração** no menu para começar a depuração.

1. De definir um ponto de interrupção no arquivo do painel de tarefas do seu projeto. Você pode definir pontos de interrupção Visual Studio Code ao passar o mouse ao lado de uma linha de código e selecionando o círculo vermelho que aparece.

    ![O círculo vermelho aparece em uma linha de código Visual Studio Code.](../images/set-breakpoint.jpg)

1. Execute o seu complemento. Você verá que os pontos de interrupção foram atingidos e você pode inspecionar variáveis locais.

## <a name="see-also"></a>Confira também

- [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)

- [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Depurar suplementos no Windows usando o WebView2 do Microsoft Edge (baseado em Chromium)](debug-desktop-using-edge-chromium.md)
