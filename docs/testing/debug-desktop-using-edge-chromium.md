---
title: Depurar suplementos no Windows usando o WebView2 do Microsoft Edge (baseado em Chromium)
description: Saiba como depurar Suplementos do Office que usam o WebView2 do Microsoft Edge (baseado em Chromium) usando o Depurador para a extensão do Microsoft Edge no VS Code.
ms.date: 10/05/2021
ms.localizationpriority: high
ms.openlocfilehash: 8ee266b3197a2b02dd4d072b6666cd68add6fec9
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138643"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Depurar suplementos no Windows usando o WebView2 do Edge Chromium

Os Suplementos do Office em execução no Windows podem usar o Depurador para a extensão do Microsoft Edge no VS Code para depurar em relação ao tempo de execução do WebView2 do Edge Chromium.

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10, 11
- Uma combinação de plataforma e aplicativo do Office que oferece suporte ao Microsoft Edge com WebView2 (baseado em Chromium), conforme explicado em [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Se a sua versão do Microsoft 365 for anterior a 2101, você precisará instalar o WebView2. Use as instruções para instalá-lo em [Microsoft Edge WebView2 / Embedar conteúdo da web ... com Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

1. Crie um projeto usando o [gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office). Para isso, você pode usar um dos nossos guias de início rápido, como o [Início rápido do suplemento do Outlook](../quickstarts/outlook-quickstart.md).

    > [!TIP]
    > Se você não estiver usando um suplemento baseado no gerador do Yeoman, pode ser solicitado que você ajuste uma chave do registro. Enquanto estiver na pasta raiz do seu projeto, execute o seguinte na linha de comando:  `office-add-in-debugging start <your manifest path>`

1. Abra o projeto no VS Code. Dentro do código VS, selecione **Ctrl+Shift+X** para abrir a Barra de extensões. Procure a extensão "Depurador do Microsoft Edge" e instale-a.

1. Em seguida, escolha  **Visualizar > Executar** ou digite **Ctrl+Shift+D** para alternar para a modo de depuração.

1. Nas opções **EXECUTAR E DEBUGAR**, escolha a opção Edge Chromium para seu aplicativo host, como **Excel Desktop (Edge Chromium)**. Selecione **F5** ou escolha **Executar > Iniciar Depuração** no menu para começar a depuração. Esta ação inicia automaticamente um servidor local em uma janela de Nó para hospedar seu suplemento e depois abre automaticamente o aplicativo host, como o Excel ou Word. Isso pode levar vários segundos.

1. No aplicativo host, seu suplemento agora está pronto para uso. Selecione **Mostrar Painel de Tarefas** ou execute qualquer outro comando de suplemento. Uma caixa de diálogo aparecerá, lendo:

   > WebView Stop On Load.
   > Para depurar o modo de exibição da Web, anexe o VS Code à instância de modo de exibição da Web usando o Depurador da Microsoft para extensão do Edge, e clique em OK para continuar. Para impedir que essa caixa de diálogo seja exibida no futuro, clique em Cancelar.

   Clique em **OK**.

   > [!NOTE]
   > Se você selecionar **Cancelar**, a caixa de diálogo não será mostrada novamente enquanto esta instância do suplemento estiver em execução. No entanto, se você reiniciar o suplemento, você verá a caixa de diálogo novamente.

1. Agora você pode definir pontos de interrupção no código e depuração do projeto.

   > [!NOTE]
   > Pontos de interrupção em chamadas de `Office.initialize` ou `Office.onReady` são ignorados. Para obter detalhes sobre esses métodos, consulte [Inicialize seu Suplemento do Office](../develop/initialize-add-in.md).

> [!IMPORTANT]
> A melhor maneira de interromper uma sessão de depuração é selecionar **Shift+F5** ou escolher **Executar > Interromper Depuração** no menu. Esta ação deve fechar a janela do servidor de Nó e tentar fechar o aplicativo host, mas haverá um aviso no aplicativo host perguntando se você deseja salvar o documento ou não. Faça uma escolha apropriada e deixe o aplicativo host fechar. Evite fechar manualmente a janela de Nó ou o aplicativo host. Fazer isso pode causar bugs, especialmente quando você interrompe e inicia sessões de depuração repetidamente.
>
> Se a depuração parar de funcionar; por exemplo, se os pontos de interrupção estão sendo ignorados; interrompa a depuração. Em seguida, se necessário, feche todas as janelas do aplicativo host e a janela de Nó. Finalmente, feche o Visual Studio Code e abra-o novamente.

## <a name="see-also"></a>Confira também

- [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
- [Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)
- [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)
