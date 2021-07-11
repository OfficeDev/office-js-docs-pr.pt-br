---
title: Depurar suplementos no Windows usando o WebView2 do Microsoft Edge (baseado em Chromium)
description: Saiba como depurar Suplementos do Office que usam o WebView2 do Microsoft Edge (baseado em Chromium) usando o Depurador para a extensão do Microsoft Edge no VS Code.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 6a62718147fbb5d2e8a6819066425737d853cbf0
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350173"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a><span data-ttu-id="effe0-103">Depurar suplementos no Windows usando o WebView2 do Edge Chromium</span><span class="sxs-lookup"><span data-stu-id="effe0-103">Debug add-ins on Windows using Edge Chromium WebView2</span></span>

<span data-ttu-id="effe0-104">Os Suplementos do Office em execução no Windows podem usar o Depurador para a extensão do Microsoft Edge no VS Code para depurar em relação ao tempo de execução do WebView2 do Edge Chromium.</span><span class="sxs-lookup"><span data-stu-id="effe0-104">Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="effe0-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="effe0-105">Prerequisites</span></span>

- <span data-ttu-id="effe0-106">[Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)</span><span class="sxs-lookup"><span data-stu-id="effe0-106">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="effe0-107">Node.js (versão 10+)</span><span class="sxs-lookup"><span data-stu-id="effe0-107">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="effe0-108">Windows 10</span><span class="sxs-lookup"><span data-stu-id="effe0-108">Windows 10</span></span>
- [<span data-ttu-id="effe0-109">Microsoft Edge Chromium disponível para Usuários do Windows Insider</span><span class="sxs-lookup"><span data-stu-id="effe0-109">Microsoft Edge Chromium available to Windows Insiders</span></span>](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="effe0-110">Instalar e usar o depurador</span><span class="sxs-lookup"><span data-stu-id="effe0-110">Install and use the debugger</span></span>

1. <span data-ttu-id="effe0-111">Crie um projeto usando o [gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office). Para isso, você pode usar um dos nossos guias de início rápido, como o [Início rápido do suplemento do Outlook](../quickstarts/outlook-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="effe0-111">Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.</span></span>

    > [!TIP]
    > <span data-ttu-id="effe0-112">Se você não estiver usando um suplemento baseado em um gerador Yeoman, será necessário ajustar uma chave de registro.</span><span class="sxs-lookup"><span data-stu-id="effe0-112">If you aren't using a Yeoman generator based add-in, you need to adjust a registry key.</span></span> <span data-ttu-id="effe0-113">Enquanto estiver na pasta raiz do seu projeto, execute o seguinte na linha de comando: `office-add-in-debugging start <your manifest path>`.</span><span class="sxs-lookup"><span data-stu-id="effe0-113">While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.</span></span>

1. <span data-ttu-id="effe0-114">Abra o projeto no VS Code.</span><span class="sxs-lookup"><span data-stu-id="effe0-114">Open your project in VS Code.</span></span> <span data-ttu-id="effe0-115">No VS Code, selecione **Ctrl+Shift+X** para abrir a barra Extensões.</span><span class="sxs-lookup"><span data-stu-id="effe0-115">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="effe0-116">Procure a extensão "Depurador do Microsoft Edge" e instale-a.</span><span class="sxs-lookup"><span data-stu-id="effe0-116">Search for the "Debugger for Microsoft Edge" extension and install it.</span></span>

1. <span data-ttu-id="effe0-117">Na pasta **.vscode** do seu projeto, abra o arquivo **launch.json**.</span><span class="sxs-lookup"><span data-stu-id="effe0-117">In the **.vscode** folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="effe0-118">Adicione o código a seguir à seção de configurações.</span><span class="sxs-lookup"><span data-stu-id="effe0-118">Add the following code to the configurations section.</span></span>

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

1. <span data-ttu-id="effe0-119">Em seguida, escolha  **Exibir > Depurar** ou digite **Ctrl+Shift+D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="effe0-119">Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

1. <span data-ttu-id="effe0-120">Nas opções de Depuração, escolha a opção Edge Chromium para seu aplicativo host, como **Excel Desktop (Edge Chromium)**.</span><span class="sxs-lookup"><span data-stu-id="effe0-120">From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**.</span></span> <span data-ttu-id="effe0-121">Selecione **F5** ou escolha **Depurar > Iniciar Depuração** no menu para começar a depuração.</span><span class="sxs-lookup"><span data-stu-id="effe0-121">Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.</span></span>

1. <span data-ttu-id="effe0-122">No aplicativo host, como o Excel, o seu suplemento está agora pronto para uso.</span><span class="sxs-lookup"><span data-stu-id="effe0-122">In the host application, such as Excel, your add-in is now ready to use.</span></span> <span data-ttu-id="effe0-123">Selecione **Mostrar Painel de Tarefas** ou execute qualquer outro comando de suplemento.</span><span class="sxs-lookup"><span data-stu-id="effe0-123">Select **Show Taskpane** or run any other add-in command.</span></span> <span data-ttu-id="effe0-124">Uma caixa de diálogo aparecerá, lendo:</span><span class="sxs-lookup"><span data-stu-id="effe0-124">A dialog box will appear, reading:</span></span>

    > <span data-ttu-id="effe0-125">WebView Stop On Load.</span><span class="sxs-lookup"><span data-stu-id="effe0-125">WebView Stop On Load.</span></span>
    > <span data-ttu-id="effe0-126">Para depurar o modo de exibição da Web, anexe o VS Code à instância de modo de exibição da Web usando o Depurador da Microsoft para extensão do Edge, e clique em OK para continuar.</span><span class="sxs-lookup"><span data-stu-id="effe0-126">To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue.</span></span> <span data-ttu-id="effe0-127">Para impedir que essa caixa de diálogo seja exibida no futuro, clique em Cancelar."</span><span class="sxs-lookup"><span data-stu-id="effe0-127">To prevent this dialog from appearing in the future, click Cancel."</span></span>

    <span data-ttu-id="effe0-128">Clique em **OK**.</span><span class="sxs-lookup"><span data-stu-id="effe0-128">Select **OK**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="effe0-129">Se você selecionar **Cancelar**, a caixa de diálogo não será mostrada novamente enquanto esta instância do suplemento estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="effe0-129">If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running.</span></span> <span data-ttu-id="effe0-130">No entanto, se você reiniciar o suplemento, você verá a caixa de diálogo novamente.</span><span class="sxs-lookup"><span data-stu-id="effe0-130">However, if you restart your add-in, you'll see the dialog again.</span></span>

1. <span data-ttu-id="effe0-131">Agora você pode definir pontos de interrupção no código e depuração do projeto.</span><span class="sxs-lookup"><span data-stu-id="effe0-131">You're now able to set breakpoints in your project's code and debug.</span></span>

## <a name="see-also"></a><span data-ttu-id="effe0-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="effe0-132">See also</span></span>

- [<span data-ttu-id="effe0-133">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="effe0-133">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
- [<span data-ttu-id="effe0-134">Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="effe0-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="effe0-135">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="effe0-135">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)