---
title: Depurar os complementos no Windows usando Visual Studio Code e Microsoft Edge WebView herdado (EdgeHTML)
description: Saiba como depurar Office Depuração de Versão Prévia do Microsoft Edge Que usam o WebView (EdgeHTML) usando Office Extensão de Depurador de Office no VS Code.
ms.date: 02/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 11b728f9b3f467017711c9d75cfd07767957deae
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467691"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code

Office Os Windows em execução no Windows podem usar Office Extensão de Depurador de Complementos no Visual Studio Code para depurar Versão Prévia do Microsoft Edge com o tempo de execução do WebView (EdgeHTML) original. 

> [!IMPORTANT]
> Este artigo só se aplica quando Office executa os complementos no tempo de execução do WebView (EdgeHTML) original, conforme explicado em [Navegadores usados por Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Para obter instruções sobre Visual Studio depuração em um código Microsoft Edge WebView2 (baseado em Chromium), consulte [Microsoft Office Extensão de Depurador](debug-desktop-using-edge-chromium.md) de Microsoft Office para Visual Studio Code.

> [!TIP]
> Se você não puder, ou não desejar, depurar usando ferramentas criadas no Visual Studio Code ou estiver encontrando um problema que só ocorre quando o complemento é executado fora do Visual Studio Code, você pode depurar o tempo de execução do Edge Legacy (EdgeHTML) usando as ferramentas de desenvolvedor herdadas de borda conforme descrito em [Depurar os complementos usando ferramentas de desenvolvedor em Versão Prévia do Microsoft Edge](debug-add-ins-using-devtools-edge-legacy.md).

Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução. Você pode ver alterações em seu código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração. As alterações de código também persistem, para que você possa ver os resultados de várias alterações no código. A imagem a seguir mostra essa extensão em ação.

![Office Extensão de Depurador de Complementos depurando uma seção de Excel de complementos.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Pré-requisitos

- [Código do Visual Studio](https://code.visualstudio.com/)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10, 11
- [Microsoft Edge](https://www.microsoft.com/edge) Uma combinação de plataforma e aplicativo Office que oferece suporte Versão Prévia do Microsoft Edge com o webview original (EdgeHTML), conforme explicado em [Navegadores usados por Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

Estas instruções pressuem que você tenha experiência usando a linha de comando, entenda JavaScript básico e tenha criado um projeto de Office de Office antes de usar o gerador Yo Office. Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este Excel Office [tutorial de complemento](../tutorials/excel-tutorial.md).

1. A primeira etapa depende do projeto e de como ele foi criado.

   - Se você quiser criar um projeto para experimentar a depuração no Visual Studio Code, use o gerador [Yeoman para Office Desempois](https://github.com/OfficeDev/generator-office). Use qualquer um de nossos guias de início rápido, como o Outlook de início rápido do Outlook de [complemento, para](../quickstarts/outlook-quickstart.md) fazer isso. 
   - Se você quiser depurar um projeto existente que foi criado com Yo Office, pule para a próxima etapa.
   - Se você quiser depurar um projeto existente que não foi criado com Yo Office, realize o procedimento no Apêndice e retorne para a próxima [](#appendix) etapa deste procedimento.


1. Abra VS Code e abra seu projeto nele. 

1. Dentro do código VS, selecione **Ctrl+Shift+X** para abrir a Barra de extensões. Procure a extensão "Microsoft Office Depurador de Complementos" e instale-a.

1. Escolha  **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o exibição de depuração.

1. Nas opções **EXECUTAR E DEPURar**, escolha a opção Legado de Borda para seu aplicativo host, como Outlook **Desktop (Edge Legacy)**. Selecione **F5** ou escolha **Executar > Iniciar Depuração** no menu para começar a depuração. Esta ação inicia automaticamente um servidor local em uma janela de Nó para hospedar seu suplemento e depois abre automaticamente o aplicativo host, como o Excel ou Word. Isso pode levar vários segundos.

1. No aplicativo host, seu suplemento agora está pronto para uso. Selecione **Mostrar Painel de Tarefas** ou execute qualquer outro comando de suplemento. Uma caixa de diálogo aparecerá semelhante à seguinte:

   > WebView Stop On Load.
   > Para depurar o WebView, anexe VS Code à instância WebView usando a extensão Microsoft Depurador para Borda e clique em **OK** para continuar. Para impedir que essa caixa de diálogo apareça no futuro, clique em **Cancelar**.

   Clique em **OK**.

   > [!NOTE]
   > Se você selecionar **Cancelar**, a caixa de diálogo não será mostrada novamente enquanto esta instância do suplemento estiver em execução. No entanto, se você reiniciar o suplemento, você verá a caixa de diálogo novamente.

1. De definir um ponto de interrupção no arquivo do painel de tarefas do seu projeto. Para definir pontos de interrupção Visual Studio Code, passe o mouse ao lado de uma linha de código e selecione o círculo vermelho que aparece.

    ![O círculo vermelho aparece em uma linha de código Visual Studio Code.](../images/set-breakpoint.jpg)

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
      "name": "$HOST$ Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=Excel$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: Excel Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. Substitua o espaço reservado `$HOST$` pelo nome do Office aplicativo em que o complemento é executado; por exemplo, `Outlook` ou `Word`.
1. Salve e feche o arquivo.

## <a name="see-also"></a>Confira também

- [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
- [Depurar os Windows usando Visual Studio Code e Microsoft Edge WebView2 (baseados em Chromium)](debug-desktop-using-edge-chromium.md).
- [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
