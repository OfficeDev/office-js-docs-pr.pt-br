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
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="84fa0-103">Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="84fa0-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="84fa0-104">O Microsoft Office Add-in Debugger Extension para o Visual Studio Code permite que você depure seu suplemento do Office em tempo de execução de borda.</span><span class="sxs-lookup"><span data-stu-id="84fa0-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="84fa0-105">Este modo de depuração é dinâmico, permitindo que você defina pontos de interrupção enquanto o código está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="84fa0-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="84fa0-106">Você pode ver alterações no seu código imediatamente enquanto o depurador é anexado, tudo sem perder a sessão de depuração.</span><span class="sxs-lookup"><span data-stu-id="84fa0-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="84fa0-107">Suas alterações de código também persistim, portanto, você pode ver os resultados de várias alterações em seu código.</span><span class="sxs-lookup"><span data-stu-id="84fa0-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="84fa0-108">A imagem a seguir mostra essa extensão em ação.</span><span class="sxs-lookup"><span data-stu-id="84fa0-108">The following image shows this extension in action.</span></span>

![Extensão do depurador de suplementos do Office depuração de uma seção de suplementos do Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="84fa0-110">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="84fa0-110">Prerequisites</span></span>

- <span data-ttu-id="84fa0-111">[Visual Studio Code](https://code.visualstudio.com/) (deve ser executado como um administrador)</span><span class="sxs-lookup"><span data-stu-id="84fa0-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="84fa0-112">Node.js (versão 10 +)</span><span class="sxs-lookup"><span data-stu-id="84fa0-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="84fa0-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="84fa0-113">Windows 10</span></span>
- [<span data-ttu-id="84fa0-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="84fa0-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="84fa0-115">Estas instruções pressupõem que você tenha experiência em usar a linha de comando, entenda o JavaScript básico e criou um projeto de suplemento do Office antes de usar o gerador do Office Yo.</span><span class="sxs-lookup"><span data-stu-id="84fa0-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="84fa0-116">Se você ainda não fez isso antes, considere visitar um de nossos tutoriais, como este [tutorial de suplemento do Office Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="84fa0-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="84fa0-117">Instalar e usar o depurador</span><span class="sxs-lookup"><span data-stu-id="84fa0-117">Install and use the debugger</span></span>

1. <span data-ttu-id="84fa0-118">Se você precisar criar um projeto de suplemento, [use o gerador de Yo do Office para criar um](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="84fa0-118">If you need to create an add-in project, [use the Yo Office generator to create one](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span></span> <span data-ttu-id="84fa0-119">Siga os prompts dentro da linha de comando para configurar seu projeto.</span><span class="sxs-lookup"><span data-stu-id="84fa0-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="84fa0-120">Você pode escolher qualquer idioma ou tipo de projeto para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="84fa0-120">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="84fa0-121">Se você já tiver um projeto, pule a etapa 1 e vá para a etapa 2.</span><span class="sxs-lookup"><span data-stu-id="84fa0-121">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="84fa0-122">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="84fa0-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="84fa0-123">![Opções de prompt de comando, incluindo "executar como administrador" no Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="84fa0-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="84fa0-124">Navegue até o diretório do projeto.</span><span class="sxs-lookup"><span data-stu-id="84fa0-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="84fa0-125">Execute o seguinte comando para abrir seu projeto no Visual Studio Code como um administrador.</span><span class="sxs-lookup"><span data-stu-id="84fa0-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="84fa0-126">Depois que o Visual Studio code estiver aberto, navegue manualmente para a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="84fa0-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="84fa0-127">Para abrir o Visual Studio Code como um administrador, selecione a opção **Executar como administrador** ao abrir o código do Visual Studio após procurá-lo no Windows.</span><span class="sxs-lookup"><span data-stu-id="84fa0-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="84fa0-128">No VS Code, selecione **Ctrl + Shift + X** para abrir a barra de extensões.</span><span class="sxs-lookup"><span data-stu-id="84fa0-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="84fa0-129">Procure a extensão "depurador de suplementos do Microsoft Office" e instale-a.</span><span class="sxs-lookup"><span data-stu-id="84fa0-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="84fa0-130">Na pasta. vscode do projeto, abra o **launch.jsem** arquivo.</span><span class="sxs-lookup"><span data-stu-id="84fa0-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="84fa0-131">Adicione o seguinte código à `configurations` seção:</span><span class="sxs-lookup"><span data-stu-id="84fa0-131">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="84fa0-132">Na seção de JSON que você acabou de copiar, encontre a seção "URL".</span><span class="sxs-lookup"><span data-stu-id="84fa0-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="84fa0-133">Nesta URL, será necessário substituir o texto de HOST em maiúsculas pelo aplicativo que está hospedando o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="84fa0-133">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office add-in.</span></span> <span data-ttu-id="84fa0-134">Por exemplo, se o suplemento do Office for Excel, seu valor de URL será " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="84fa0-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="84fa0-135">Abra o prompt de comando e verifique se você está na pasta raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="84fa0-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="84fa0-136">Execute o comando `npm start` para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="84fa0-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="84fa0-137">Quando o suplemento for carregado no cliente do Office, abra o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="84fa0-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="84fa0-138">Retorne ao Visual Studio Code e escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="84fa0-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="84fa0-139">Nas opções de depuração, escolha **anexar a suplementos do Office**. Selecione **F5** ou escolha **debug-> iniciar a depuração** no menu para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="84fa0-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="84fa0-140">Defina um ponto de interrupção no arquivo de painel de tarefas do projeto.</span><span class="sxs-lookup"><span data-stu-id="84fa0-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="84fa0-141">É possível definir pontos de interrupção no VS Code focalizando ao lado de uma linha de código e selecionando o círculo vermelho que aparece.</span><span class="sxs-lookup"><span data-stu-id="84fa0-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![Um círculo vermelho aparece em uma linha de código no VS Code](../images/set-breakpoint.jpg)

12. <span data-ttu-id="84fa0-143">Execute o suplemento.</span><span class="sxs-lookup"><span data-stu-id="84fa0-143">Run your add-in.</span></span> <span data-ttu-id="84fa0-144">Você verá que os pontos de interrupção foram atingidos e pode inspecionar as variáveis locais.</span><span class="sxs-lookup"><span data-stu-id="84fa0-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="84fa0-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="84fa0-145">See also</span></span>

* [<span data-ttu-id="84fa0-146">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="84fa0-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="84fa0-147">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="84fa0-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="84fa0-148">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="84fa0-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
