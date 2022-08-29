---
title: Depurar os suplementos no Windows usando Visual Studio Code e Microsoft Edge WebView herdado (EdgeHTML)
description: Saiba como depurar suplementos do Office que usam o Versão Prévia do Microsoft Edge WebView (EdgeHTML) usando a Extensão do Depurador de Suplementos do Office no VS Code.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 87e503d3a79b5fa4b797bb9c6ee657b7d8916109
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423234"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code

Os Suplementos do Office em execução no Windows podem usar a Extensão do Depurador de Suplementos do Office no Visual Studio Code para depurar em Versão Prévia do Microsoft Edge com o runtime original do WebView (EdgeHTML). 

> [!IMPORTANT]
> Este artigo só se aplica quando o Office executa suplementos no runtime original do WebView (EdgeHTML), conforme explicado em Navegadores usados pelos [Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Para obter instruções sobre a depuração no código do Visual Studio no Microsoft Edge WebView2 (baseado em Chromium), consulte a Extensão do Depurador de Suplementos do [Microsoft Office para Visual Studio Code](debug-desktop-using-edge-chromium.md).

> [!TIP]
> Se você não puder, ou não quiser, depurar usando ferramentas internas no Visual Studio Code; ou estiver encontrando um problema que só ocorre quando o suplemento é executado fora do Visual Studio Code, você pode depurar o runtime edge legacy (EdgeHTML) usando as ferramentas de desenvolvedor herdadas do Edge, conforme descrito nos [suplementos de depuração usando ferramentas de desenvolvedor no Versão Prévia do Microsoft Edge.](debug-add-ins-using-devtools-edge-legacy.md)

Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução. Você pode ver as alterações no código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração. As alterações de código também persistem, para que você possa ver os resultados de várias alterações em seu código. A imagem a seguir mostra essa extensão em ação.

![Extensão do Depurador de Suplemento do Office depurando uma seção de suplementos do Excel.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Pré-requisitos

- [Código do Visual Studio](https://code.visualstudio.com/)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10, 11
- [Microsoft Edge](https://www.microsoft.com/edge) Uma combinação de plataforma e aplicativo do Office que dá suporte Versão Prévia do Microsoft Edge com o modo de exibição da Web original (EdgeHTML), conforme explicado em Navegadores usados pelos [Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

Estas instruções pressupõem que você tenha experiência com o uso da linha de comando, entenda o JavaScript básico e tenha criado um projeto de suplemento do Office antes de usar o gerador [Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md). Se você ainda não fez isso, considere visitar um de nossos tutoriais, como este tutorial de Suplemento [do Office do Excel](../tutorials/excel-tutorial.md).

1. A primeira etapa depende do projeto e de como ele foi criado.

   - Se você quiser criar um projeto para experimentar a depuração no Visual Studio Code, use o [Gerador Yeoman para Suplemento do Office](../develop/yeoman-generator-overview.md). Use qualquer um de nossos guias de início rápido, como o [início rápido do suplemento do Outlook](../quickstarts/outlook-quickstart.md), para fazer isso. 
   - Se você quiser depurar um projeto existente que foi criado com Yo Office, pule para a próxima etapa.
   - Se você quiser depurar um projeto existente que não foi criado com o Yo Office, execute o procedimento no Apêndice e [](#appendix) retorne à próxima etapa deste procedimento.


1. Abra VS Code e abra seu projeto nele. 

1. Dentro do código VS, selecione **Ctrl+Shift+X** para abrir a Barra de extensões. Pesquise a extensão "Depurador de Suplementos do Microsoft Office" e instale-a.

1. Escolha  **Exibir > Executar** ou insira **Ctrl+Shift+D** para alternar para o exibição de depuração.

1. Nas opções **EXECUTAR E DEPURAR** , escolha a opção Edge Legacy para seu aplicativo host, como a Área de Trabalho do **Outlook (Versão Prévia do Edge)**. Selecione **F5** ou escolha **Executar > Iniciar Depuração** no menu para começar a depuração. Esta ação inicia automaticamente um servidor local em uma janela de Nó para hospedar seu suplemento e depois abre automaticamente o aplicativo host, como o Excel ou Word. Isso pode levar vários segundos.

1. No aplicativo host, seu suplemento agora está pronto para uso. Selecione **Mostrar Painel de Tarefas** ou execute qualquer outro comando de suplemento. Uma caixa de diálogo será semelhante à seguinte:

   > WebView Stop On Load.
   > Para depurar o WebView, anexe o VS Code à instância do WebView usando a extensão Microsoft Debugger para Edge e clique em **OK** para continuar. Para impedir que essa caixa de diálogo apareça no futuro, clique em **Cancelar**.

   Clique em **OK**.

   > [!NOTE]
   > Se você selecionar **Cancelar**, a caixa de diálogo não será mostrada novamente enquanto esta instância do suplemento estiver em execução. No entanto, se você reiniciar o suplemento, você verá a caixa de diálogo novamente.

1. Defina um ponto de interrupção no arquivo do painel de tarefas do projeto. Para definir pontos de interrupção Visual Studio Code, passe o mouse ao lado de uma linha de código e selecione o círculo vermelho que aparece.

    ![O círculo vermelho aparece em uma linha de código no Visual Studio Code.](../images/set-breakpoint.jpg)

1. Execute a funcionalidade no seu complemento que chama as linhas com pontos de interrupção. Você verá que os pontos de interrupção foram atingidos e você pode inspecionar variáveis locais.

   > [!NOTE]
   > Pontos de interrupção em chamadas de `Office.initialize` ou `Office.onReady` são ignorados. Para obter detalhes sobre esses métodos, consulte [Inicialize seu Suplemento do Office](../develop/initialize-add-in.md).

> [!IMPORTANT]
> A melhor maneira de interromper uma sessão de depuração é selecionar **Shift+F5** ou escolher **Executar > Interromper Depuração** no menu. Esta ação deve fechar a janela do servidor de Nó e tentar fechar o aplicativo host, mas haverá um aviso no aplicativo host perguntando se você deseja salvar o documento ou não. Faça uma escolha apropriada e deixe o aplicativo host fechar. Evite fechar manualmente a janela de Nó ou o aplicativo host. Fazer isso pode causar bugs, especialmente quando você interrompe e inicia sessões de depuração repetidamente.
>
> Se a depuração parar de funcionar; por exemplo, se os pontos de interrupção estão sendo ignorados; interrompa a depuração. Em seguida, se necessário, feche todas as janelas do aplicativo host e a janela de Nó. Finalmente, feche o Visual Studio Code e abra-o novamente.

### <a name="appendix"></a>Apêndice

Se seu projeto não tiver sido criado com o Yo Office, você precisará criar uma configuração de depuração para Visual Studio Code. 

1. Crie um arquivo nomeado `launch.json` na `\.vscode` pasta do projeto se ainda não houver um. 
1. Verifique se o arquivo tem uma `configurations` matriz. A seguir, um exemplo simples de um `launch.json`.

    ```json
    {
      // other properities may be here.

      "configurations": [

        // configuration objects may be here.

      ]

      //other properies may be here.
    }
    ```

1. Adicione o objeto a seguir à `configurations` matriz.

    ```json
    {
      "name": "HOST Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: HOST Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. Substitua o espaço reservado `HOST` em todos os três locais pelo nome do aplicativo do Office no qual o suplemento é executado; por exemplo, `Outlook` ou `Word`.
1. Salve e feche o arquivo.

## <a name="see-also"></a>Confira também

- [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
- [Depurar suplementos no Windows usando o Visual Studio Code e o Microsoft Edge WebView2 (baseados Chromium)](debug-desktop-using-edge-chromium.md).
- [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
- [Runtimes em Suplementos do Office](runtimes.md)