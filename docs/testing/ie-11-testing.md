---
title: Teste do Internet Explorer 11
description: Teste seu Office no Internet Explorer 11.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727931"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="6aa97-103">Testar seu Office de usuário no Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="6aa97-103">Test your Office Add-in on Internet Explorer 11</span></span>

<span data-ttu-id="6aa97-104">Se você planeja comercializar seu complemento por meio do AppSource ou planeja dar suporte a versões mais antigas do Windows e Office, o seu complemento deve funcionar no controle de navegador in-loca que se baseia no Internet Explorer 11 (IE11).</span><span class="sxs-lookup"><span data-stu-id="6aa97-104">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="6aa97-105">Você pode usar uma linha de comando para alternar de tempos de execução mais modernos usados pelos complementos para o tempo de execução do Internet Explorer 11 para esse teste.</span><span class="sxs-lookup"><span data-stu-id="6aa97-105">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="6aa97-106">Para obter informações sobre quais versões do Windows e Office usam o controle de exibição da Web do Internet Explorer 11, consulte Navegadores usados por [Office Dep.](../concepts/browsers-used-by-office-web-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="6aa97-106">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6aa97-107">O Internet Explorer 11 não oferece suporte às versões do JavaScript posteriores a ES5.</span><span class="sxs-lookup"><span data-stu-id="6aa97-107">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="6aa97-108">Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, você tem duas opções:</span><span class="sxs-lookup"><span data-stu-id="6aa97-108">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="6aa97-109">Escreva seu código no ECMAScript 2015 (também chamado de ES6) ou javaScript posterior ou em TypeScript e compile seu código para JavaScript do ES5 usando um compilador como [o babel](https://babeljs.io/) ou [o tsc](https://www.typescriptlang.org/index.html).</span><span class="sxs-lookup"><span data-stu-id="6aa97-109">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="6aa97-110">Escreva em ECMAScript 2015 ou posterior JavaScript, mas também carregue uma biblioteca de [polifilamento,](https://en.wikipedia.org/wiki/Polyfill_(programming)) como [core-js,](https://github.com/zloirock/core-js) que permite ao IE executar seu código.</span><span class="sxs-lookup"><span data-stu-id="6aa97-110">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="6aa97-111">Para obter mais informações sobre essas opções, consulte [Support Internet Explorer 11](../develop/support-ie-11.md).</span><span class="sxs-lookup"><span data-stu-id="6aa97-111">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="6aa97-112">Além disso, o Internet Explorer 11 não oferece suporte a alguns recursos do HTML5, como mídia, gravação e localização.</span><span class="sxs-lookup"><span data-stu-id="6aa97-112">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="6aa97-113">Para testar seu complemento no navegador do Internet Explorer 11, abra o Office na Web no Internet Explorer e coloque o [sideload do add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6aa97-113">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6aa97-114">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6aa97-114">Prerequisites</span></span>

- <span data-ttu-id="6aa97-115">[Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="6aa97-115">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="6aa97-116">Estas instruções pressuem que você tenha criado um projeto de gerador Yo Office antes.</span><span class="sxs-lookup"><span data-stu-id="6aa97-116">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="6aa97-117">Se você não tiver feito isso antes, considere ler um início rápido, como este para Excel [de Excel.](../quickstarts/excel-quickstart-jquery.md)</span><span class="sxs-lookup"><span data-stu-id="6aa97-117">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="6aa97-118">Alternando para o webview do Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="6aa97-118">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="6aa97-119">Crie um projeto yo Office gerador.</span><span class="sxs-lookup"><span data-stu-id="6aa97-119">Create a Yo Office generator project.</span></span> <span data-ttu-id="6aa97-120">Não importa o tipo de projeto selecionado, essa ferramenta funcionará com todos os tipos de projeto.</span><span class="sxs-lookup"><span data-stu-id="6aa97-120">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6aa97-121">Se você tiver um projeto existente e quiser adicionar essa ferramenta sem criar um novo projeto, pule esta etapa e vá para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="6aa97-121">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="6aa97-122">Na pasta raiz do seu projeto, execute o seguinte na linha de comando.</span><span class="sxs-lookup"><span data-stu-id="6aa97-122">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="6aa97-123">Este exemplo supõe que o arquivo de manifesto do seu projeto está na raiz.</span><span class="sxs-lookup"><span data-stu-id="6aa97-123">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="6aa97-124">Se não estiver, especifique o caminho relativo para o arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="6aa97-124">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="6aa97-125">Você deve ver uma mensagem na linha de comando que o tipo de exibição da Web agora está definido como IE.</span><span class="sxs-lookup"><span data-stu-id="6aa97-125">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="6aa97-126">Não é necessário usar esse comando, mas deve ajudar a depurar a maioria dos problemas relacionados ao tempo de execução do Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="6aa97-126">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="6aa97-127">Para uma robustez completa, você deve testar o uso de computadores com várias combinações de Windows 7, 8.1 e 10 e várias versões de Office.</span><span class="sxs-lookup"><span data-stu-id="6aa97-127">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="6aa97-128">Para obter mais informações, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span><span class="sxs-lookup"><span data-stu-id="6aa97-128">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="6aa97-129">Opções de comando</span><span class="sxs-lookup"><span data-stu-id="6aa97-129">Command options</span></span>

<span data-ttu-id="6aa97-130">O comando também pode ter vários tempos de `office-addin-dev-settings webview` execução como argumentos:</span><span class="sxs-lookup"><span data-stu-id="6aa97-130">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="6aa97-131">ie</span><span class="sxs-lookup"><span data-stu-id="6aa97-131">ie</span></span>
- <span data-ttu-id="6aa97-132">edge</span><span class="sxs-lookup"><span data-stu-id="6aa97-132">edge</span></span>
- <span data-ttu-id="6aa97-133">Padrão.</span><span class="sxs-lookup"><span data-stu-id="6aa97-133">default</span></span>

## <a name="see-also"></a><span data-ttu-id="6aa97-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="6aa97-134">See also</span></span>

* [<span data-ttu-id="6aa97-135">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6aa97-135">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="6aa97-136">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="6aa97-136">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="6aa97-137">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="6aa97-137">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="6aa97-138">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="6aa97-138">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
