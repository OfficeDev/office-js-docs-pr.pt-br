---
title: Depurar suplementos no Windows usando o WebView2 do Microsoft Edge (baseado em Chromium)
description: Saiba como depurar Suplementos do Office que usam o WebView2 do Microsoft Edge (baseado em Chromium) usando o Depurador para a extensão do Microsoft Edge no VS Code.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 0908bb5040b49568006324600acacb5e36dbd1a5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238110"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Depurar suplementos no Windows usando o WebView2 do Edge Chromium

Os Suplementos do Office em execução no Windows podem usar o Depurador para a extensão do Microsoft Edge no VS Code para depurar em relação ao tempo de execução do WebView2 do Edge Chromium.

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)
- [Node.js (versão 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge Chromium disponível para Usuários do Windows Insider](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a>Instalar e usar o depurador

1. Crie um projeto usando o [gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office). Para isso, você pode usar um dos nossos guias de início rápido, como o [Início rápido do suplemento do Outlook](../quickstarts/outlook-quickstart.md).

> [!TIP]
> Se você não estiver usando um suplemento baseado em um gerador Yeoman, será necessário ajustar uma chave de registro. Enquanto estiver na pasta raiz do seu projeto, execute o seguinte na linha de comando: `office-add-in-debugging start <your manifest path>`.

2. Abra o projeto no VS Code. No VS Code, selecione **Ctrl+Shift+X** para abrir a barra Extensões. Procure a extensão "Depurador do Microsoft Edge" e instale-a.

3. Na pasta **.vscode** do seu projeto, abra o arquivo **launch.json**. Adicione o seguinte código à seção de configurações:

```JSON
  {
      "name": "Debug Office Add-in (Edge Chromium)",
      "type": "edge",
      "request": "attach",
      "useWebView": "advanced",
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
    },
```

4. Em seguida, escolha  **Exibir > Depurar** ou digite **Ctrl+Shift+D** para alternar para o modo de depuração.

5. Nas opções de Depuração, escolha a opção Edge Chromium para seu aplicativo host, como **Excel Desktop (Edge Chromium)**. Selecione **F5** ou escolha **Depurar > Iniciar Depuração** no menu para começar a depuração.

6. No aplicativo host, como o Excel, o seu suplemento está agora pronto para uso. Selecione **Mostrar Painel de Tarefas** ou execute qualquer outro comando de suplemento. Uma caixa de diálogo aparecerá, lendo:

> WebView Stop On Load. 
> Para depurar o modo de exibição da Web, anexe o VS Code à instância de modo de exibição da Web usando o Depurador da Microsoft para extensão do Edge, e clique em OK para continuar. Para impedir que essa caixa de diálogo seja exibida no futuro, clique em Cancelar."

Clique em **OK**.

> [!NOTE]
> Se você selecionar **Cancelar**, a caixa de diálogo não será mostrada novamente enquanto esta instância do suplemento estiver em execução. No entanto, se você reiniciar o suplemento, você verá a caixa de diálogo novamente.

7. Agora você pode definir pontos de interrupção no código e depuração do projeto.

## <a name="see-also"></a>Confira também

* [Testar e depurar Suplementos do Office](test-debug-office-add-ins.md)
* [Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)
* [Anexar um depurador do painel de tarefas](attach-debugger-from-task-pane.md)