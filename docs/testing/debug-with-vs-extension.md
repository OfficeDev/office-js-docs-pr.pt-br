---
title: Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code
description: Use a extensão do Visual Studio Code do Depurador do Microsoft Office Add-in para depurar seu Complemento do Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237928"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="5c57e-103">Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="5c57e-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="5c57e-104">A Extensão de Depurador de Add-in do Microsoft Office para Visual Studio Code permite que você depure seu Complemento do Office em relação ao Microsoft Edge com o tempo de execução original do WebView (EdgeHTML).</span><span class="sxs-lookup"><span data-stu-id="5c57e-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="5c57e-105">Para obter instruções sobre a depuração no Microsoft Edge WebView2 (baseado no Chromium), [consulte este artigo](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="5c57e-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="5c57e-106">Esse modo de depuração é dinâmico, permitindo definir pontos de interrupção enquanto o código está em execução.</span><span class="sxs-lookup"><span data-stu-id="5c57e-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="5c57e-107">Você pode ver as alterações em seu código imediatamente enquanto o depurador está anexado, tudo sem perder sua sessão de depuração.</span><span class="sxs-lookup"><span data-stu-id="5c57e-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="5c57e-108">As alterações de código também persistem, para que você possa ver os resultados de várias alterações em seu código.</span><span class="sxs-lookup"><span data-stu-id="5c57e-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="5c57e-109">A imagem a seguir mostra essa extensão em ação.</span><span class="sxs-lookup"><span data-stu-id="5c57e-109">The following image shows this extension in action.</span></span>

![Extensão do Depurador de Addin do Office depurando uma seção de Complementos do Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="5c57e-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="5c57e-111">Prerequisites</span></span>

- <span data-ttu-id="5c57e-112">[Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como administrador)</span><span class="sxs-lookup"><span data-stu-id="5c57e-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="5c57e-113">Node.js (versão 10+)</span><span class="sxs-lookup"><span data-stu-id="5c57e-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="5c57e-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="5c57e-114">Windows 10</span></span>
- [<span data-ttu-id="5c57e-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="5c57e-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="5c57e-116">Estas instruções presumem que você tenha experiência com o uso da linha de comando, compreenda o JavaScript básico e tenha criado um projeto de Complemento do Office antes de usar o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="5c57e-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="5c57e-117">Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este tutorial de Complemento [do Office do Excel.](../tutorials/excel-tutorial.md)</span><span class="sxs-lookup"><span data-stu-id="5c57e-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="5c57e-118">Instalar e usar o depurador</span><span class="sxs-lookup"><span data-stu-id="5c57e-118">Install and use the debugger</span></span>

1. <span data-ttu-id="5c57e-119">Se você precisar criar um projeto de complemento, [use o gerador Yo Office para criar um.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)</span><span class="sxs-lookup"><span data-stu-id="5c57e-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="5c57e-120">Siga os prompts dentro da linha de comando para configurar seu projeto.</span><span class="sxs-lookup"><span data-stu-id="5c57e-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="5c57e-121">Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="5c57e-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="5c57e-122">Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.</span><span class="sxs-lookup"><span data-stu-id="5c57e-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="5c57e-123">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="5c57e-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="5c57e-124">![Opções do prompt de comando, incluindo "executar como administrador" no Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="5c57e-124">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="5c57e-125">Navegue até o diretório do projeto.</span><span class="sxs-lookup"><span data-stu-id="5c57e-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="5c57e-126">Execute o seguinte comando para abrir seu projeto no Visual Studio Code como administrador.</span><span class="sxs-lookup"><span data-stu-id="5c57e-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="5c57e-127">Depois que o Visual Studio Code for aberto, navegue manualmente até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="5c57e-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="5c57e-128">Para abrir o Visual Studio Code como administrador, selecione a opção **executar** como administrador ao abrir o Visual Studio Code depois de procurar no Windows.</span><span class="sxs-lookup"><span data-stu-id="5c57e-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="5c57e-129">No VS Code, selecione **CTRL + SHIFT + X** para abrir a barra extensões.</span><span class="sxs-lookup"><span data-stu-id="5c57e-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="5c57e-130">Procure a extensão "Depurador de Complementos do Microsoft Office" e instale-a.</span><span class="sxs-lookup"><span data-stu-id="5c57e-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="5c57e-131">Na pasta .vscode do seu projeto, abra o **launch.jsno** arquivo.</span><span class="sxs-lookup"><span data-stu-id="5c57e-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="5c57e-132">Adicione o seguinte código à `configurations` seção:</span><span class="sxs-lookup"><span data-stu-id="5c57e-132">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="5c57e-133">Na seção do JSON que você acabou de copiar, encontre a seção "url".</span><span class="sxs-lookup"><span data-stu-id="5c57e-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="5c57e-134">Nesta URL, você precisará substituir o texto HOST em maiúsculas pelo aplicativo que está hospedando o Seu Complemento do Office.</span><span class="sxs-lookup"><span data-stu-id="5c57e-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="5c57e-135">Por exemplo, se o seu Complemento do Office for para Excel, seu valor de URL seria " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="5c57e-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="5c57e-136">Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="5c57e-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="5c57e-137">Execute o comando `npm start` para iniciar o servidor dev.</span><span class="sxs-lookup"><span data-stu-id="5c57e-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="5c57e-138">Quando o seu complemento for carregado no cliente do Office, abra o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="5c57e-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="5c57e-139">Retorne ao Visual Studio Code e escolha **Exibir > Depurar** ou insira **CTRL + SHIFT + D** para alternar para o exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="5c57e-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="5c57e-140">Nas opções de Depuração, escolha **Anexar aos Complementos do Office.** Selecione **F5** ou **Depurar -> Iniciar Depuração** no menu para começar a depuração.</span><span class="sxs-lookup"><span data-stu-id="5c57e-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="5c57e-141">Definir um ponto de interrupção no arquivo do painel de tarefas do projeto.</span><span class="sxs-lookup"><span data-stu-id="5c57e-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="5c57e-142">Você pode definir pontos de interrupção no VS Code ao passar o mouse ao lado de uma linha de código e selecionando o círculo vermelho que aparece.</span><span class="sxs-lookup"><span data-stu-id="5c57e-142">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![Um círculo vermelho aparece em uma linha de código no VS Code](../images/set-breakpoint.jpg)

12. <span data-ttu-id="5c57e-144">Execute o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="5c57e-144">Run your add-in.</span></span> <span data-ttu-id="5c57e-145">Você verá que pontos de interrupção foram atingidos e poderá inspecionar variáveis locais.</span><span class="sxs-lookup"><span data-stu-id="5c57e-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="5c57e-146">Confira também</span><span class="sxs-lookup"><span data-stu-id="5c57e-146">See also</span></span>

* [<span data-ttu-id="5c57e-147">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5c57e-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="5c57e-148">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="5c57e-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="5c57e-149">Depurar complementos no Windows usando o Microsoft Edge WebView2 (baseado no Chromium)</span><span class="sxs-lookup"><span data-stu-id="5c57e-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
