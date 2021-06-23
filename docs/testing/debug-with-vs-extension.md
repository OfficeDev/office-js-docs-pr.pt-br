---
title: Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code
description: Use o Visual Studio Code de Microsoft Office Depurador de Complementos para depurar seu Office Add-in.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 264a5d43a8b4f0faf7d6216664d30d7c8b64cccc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077117"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="722af-103">Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="722af-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="722af-104">A extensão de depurador de Microsoft Office do Visual Studio Code permite depurar seu Office Add-in no Microsoft Edge com o tempo de execução do WebView (EdgeHTML) original.</span><span class="sxs-lookup"><span data-stu-id="722af-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="722af-105">Para obter instruções sobre a depuração em Microsoft Edge WebView2 (Chromium baseado em Chromium), [consulte este artigo](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="722af-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="722af-106">Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução.</span><span class="sxs-lookup"><span data-stu-id="722af-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="722af-107">Você pode ver alterações em seu código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração.</span><span class="sxs-lookup"><span data-stu-id="722af-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="722af-108">As alterações de código também persistem, para que você possa ver os resultados de várias alterações no código.</span><span class="sxs-lookup"><span data-stu-id="722af-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="722af-109">A imagem a seguir mostra essa extensão em ação.</span><span class="sxs-lookup"><span data-stu-id="722af-109">The following image shows this extension in action.</span></span>

![Office Extensão de depurador de add-in depurando uma seção de Excel de complementos.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="722af-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="722af-111">Prerequisites</span></span>

- <span data-ttu-id="722af-112">[Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)</span><span class="sxs-lookup"><span data-stu-id="722af-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="722af-113">Node.js (versão 10+)</span><span class="sxs-lookup"><span data-stu-id="722af-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="722af-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="722af-114">Windows 10</span></span>
- [<span data-ttu-id="722af-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="722af-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="722af-116">Estas instruções pressuem que você tenha experiência usando a linha de comando, entenda JavaScript básico e tenha criado um projeto de Office de Office antes de usar o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="722af-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="722af-117">Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este Excel Office [tutorial de complemento.](../tutorials/excel-tutorial.md)</span><span class="sxs-lookup"><span data-stu-id="722af-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="722af-118">Instalar e usar o depurador</span><span class="sxs-lookup"><span data-stu-id="722af-118">Install and use the debugger</span></span>

1. <span data-ttu-id="722af-119">Se você precisar criar um projeto de add-in, [use o gerador Yo Office para criar um](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="722af-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="722af-120">Siga os prompts dentro da linha de comando para configurar seu projeto.</span><span class="sxs-lookup"><span data-stu-id="722af-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="722af-121">Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="722af-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="722af-122">Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.</span><span class="sxs-lookup"><span data-stu-id="722af-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="722af-123">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="722af-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="722af-124">![Opções de prompt de comando, incluindo "executar como administrador" no Windows 10.](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="722af-124">![Command prompt options, including "run as administrator" in Windows 10.](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="722af-125">Navegue até o diretório do projeto.</span><span class="sxs-lookup"><span data-stu-id="722af-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="722af-126">Execute o seguinte comando para abrir seu projeto Visual Studio Code como administrador.</span><span class="sxs-lookup"><span data-stu-id="722af-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="722af-127">Depois Visual Studio Code abrir, navegue manualmente até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="722af-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="722af-128">Para abrir Visual Studio Code como administrador, selecione  a opção executar como administrador ao abrir Visual Studio Code depois de procurá-lo no Windows.</span><span class="sxs-lookup"><span data-stu-id="722af-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="722af-129">No VS Code, selecione **Ctrl+Shift+X** para abrir a barra Extensões.</span><span class="sxs-lookup"><span data-stu-id="722af-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="722af-130">Procure a extensão "Microsoft Office Depurador de Complementos" e instale-a.</span><span class="sxs-lookup"><span data-stu-id="722af-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="722af-131">Na pasta .vscode do seu projeto, abra o arquivo **launch.json**.</span><span class="sxs-lookup"><span data-stu-id="722af-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="722af-132">Adicione o seguinte código à `configurations` seção:</span><span class="sxs-lookup"><span data-stu-id="722af-132">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="722af-133">Na seção JSON que você acabou de copiar, encontre a seção "url".</span><span class="sxs-lookup"><span data-stu-id="722af-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="722af-134">Nesta URL, você precisará substituir o texto HOST maiúscula pelo aplicativo que está hospedando seu Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="722af-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="722af-135">Por exemplo, se o Office de Office for para Excel, o valor da URL será " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16,01$en-US$ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="722af-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="722af-136">Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="722af-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="722af-137">Execute o comando `npm start` para iniciar o servidor de dev.</span><span class="sxs-lookup"><span data-stu-id="722af-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="722af-138">Quando o seu complemento for carregado no cliente Office, abra o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="722af-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="722af-139">Retorne ao Visual Studio Code e escolha **Exibir > Depurar** ou insira **CTRL + SHIFT + D** para alternar para o exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="722af-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="722af-140">Nas opções Depurar, escolha **Anexar a Office Depuração.** Selecione **F5** ou escolha **Debug -> Iniciar Depuração** no menu para começar a depuração.</span><span class="sxs-lookup"><span data-stu-id="722af-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="722af-141">De definir um ponto de interrupção no arquivo do painel de tarefas do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="722af-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="722af-142">Você pode definir pontos de interrupção Visual Studio Code ao passar o mouse ao lado de uma linha de código e selecionando o círculo vermelho que aparece.</span><span class="sxs-lookup"><span data-stu-id="722af-142">You can set breakpoints in Visual Studio Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![O círculo vermelho aparece em uma linha de código Visual Studio Code.](../images/set-breakpoint.jpg)

12. <span data-ttu-id="722af-144">Execute o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="722af-144">Run your add-in.</span></span> <span data-ttu-id="722af-145">Você verá que os pontos de interrupção foram atingidos e você pode inspecionar variáveis locais.</span><span class="sxs-lookup"><span data-stu-id="722af-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="722af-146">Confira também</span><span class="sxs-lookup"><span data-stu-id="722af-146">See also</span></span>

* [<span data-ttu-id="722af-147">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="722af-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="722af-148">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="722af-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="722af-149">Depurar suplementos no Windows usando o WebView2 do Microsoft Edge (baseado em Chromium)</span><span class="sxs-lookup"><span data-stu-id="722af-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
