---
ms.date: 05/16/2020
description: Teste seu suplemento do Office usando o Internet Explorer 11.
title: Testes do Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 1d6852d08308088a020e86ce7f5ab9cfdb9ab978
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006434"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a><span data-ttu-id="8abc8-103">Testar o suplemento do Office usando o Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="8abc8-103">Test your Office Add-in using Internet Explorer 11</span></span>

<span data-ttu-id="8abc8-104">Dependendo das especificações do seu suplemento, você pode planejar o suporte a versões mais antigas do Windows e do Office, que precisam de testes no Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="8abc8-104">Depending on the specifications of your add-in, you may plan to support older versions of Windows and Office, which require testing on Internet Explorer 11.</span></span> <span data-ttu-id="8abc8-105">Isso geralmente é necessário como parte do envio do suplemento para o AppSource.</span><span class="sxs-lookup"><span data-stu-id="8abc8-105">This is often necessary as part of submitting your add-in to AppSource.</span></span> <span data-ttu-id="8abc8-106">Você pode usar a seguinte ferramenta de linha de comando para mudar de tempos de execução mais modernos usados pelos suplementos para o tempo de execução do Internet Explorer 11 para este teste.</span><span class="sxs-lookup"><span data-stu-id="8abc8-106">You can use the following command line tooling to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span>

## <a name="pre-requisites"></a><span data-ttu-id="8abc8-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="8abc8-107">Pre-requisites</span></span>

- <span data-ttu-id="8abc8-108">[Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="8abc8-108">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>
- <span data-ttu-id="8abc8-109">Um editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="8abc8-109">A code editor.</span></span> <span data-ttu-id="8abc8-110">Recomendamos o [Visual Studio Code](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="8abc8-110">We recommend [Visual Studio Code](https://code.visualstudio.com/)</span></span>
- [<span data-ttu-id="8abc8-111">Fazer parte do programa Office Insider</span><span class="sxs-lookup"><span data-stu-id="8abc8-111">Be part of the Office Insider program</span></span>](https://insider.office.com)

<span data-ttu-id="8abc8-112">Estas instruções pressupõem que você tenha configurado um projeto de gerador do Office Yo antes.</span><span class="sxs-lookup"><span data-stu-id="8abc8-112">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="8abc8-113">Se você ainda não fez isso antes, considere ler um início rápido, como [este para suplementos do Excel](../quickstarts/excel-quickstart-jquery.md).</span><span class="sxs-lookup"><span data-stu-id="8abc8-113">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="using-ie11-tooling"></a><span data-ttu-id="8abc8-114">Usando a ferramenta de IE11</span><span class="sxs-lookup"><span data-stu-id="8abc8-114">Using IE11 tooling</span></span>

1. <span data-ttu-id="8abc8-115">Criar um projeto de gerador do Office Yo.</span><span class="sxs-lookup"><span data-stu-id="8abc8-115">Create a Yo Office generator project.</span></span> <span data-ttu-id="8abc8-116">Não importa o tipo de projeto selecionado, esta ferramenta funcionará com todos os tipos de projeto.</span><span class="sxs-lookup"><span data-stu-id="8abc8-116">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

> <span data-ttu-id="8abc8-117">! Observação Se você tiver um projeto existente e quiser adicionar essa ferramenta sem criar um novo projeto, pule esta etapa e vá para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="8abc8-117">![NOTE] If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

2. <span data-ttu-id="8abc8-118">Na pasta raiz do seu novo projeto, execute o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="8abc8-118">In the root folder of your new project, run the following in the command line:</span></span>

```command&nbsp;line
npx office-addin-dev-settings webview manifest.xml ie
```
<span data-ttu-id="8abc8-119">Você verá uma observação na linha de comando que o tipo de modo de exibição da Web agora está definido como IE.</span><span class="sxs-lookup"><span data-stu-id="8abc8-119">You should see a note in the command line that the web view type is now set to IE.</span></span>

> <span data-ttu-id="8abc8-120">! Tip Não é necessário usar essa ferramenta, mas ela deve ajudar a depurar a maioria dos problemas relacionados ao tempo de execução do Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="8abc8-120">![TIP] It isn't necessary to use this tooling, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="8abc8-121">Para uma robustez completa, você deve testar usando um computador com uma cópia do Windows 7 e do Office 2013 instalados.</span><span class="sxs-lookup"><span data-stu-id="8abc8-121">For complete robustness, you should test using a computer with a copy of Windows 7 and Office 2013 installed.</span></span>

## <a name="command-settings"></a><span data-ttu-id="8abc8-122">Configurações de comando</span><span class="sxs-lookup"><span data-stu-id="8abc8-122">Command settings</span></span>

<span data-ttu-id="8abc8-123">Se você tiver um caminho de manifesto diferente, especifique-o no comando, conforme mostrado a seguir:</span><span class="sxs-lookup"><span data-stu-id="8abc8-123">Should you have a different manifest path, specify this in the command, as shown in the following:</span></span>

`npx office-addin-dev-settings webview [path to your manifest] ie`

<span data-ttu-id="8abc8-124">O `office-addin-dev-settings webview` comando também pode ter vários tempos de execução como argumentos:</span><span class="sxs-lookup"><span data-stu-id="8abc8-124">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="8abc8-125">i</span><span class="sxs-lookup"><span data-stu-id="8abc8-125">ie</span></span>
- <span data-ttu-id="8abc8-126">vertical</span><span class="sxs-lookup"><span data-stu-id="8abc8-126">edge</span></span>
- <span data-ttu-id="8abc8-127">Padrão.</span><span class="sxs-lookup"><span data-stu-id="8abc8-127">default</span></span>

## <a name="see-also"></a><span data-ttu-id="8abc8-128">Também consulte</span><span class="sxs-lookup"><span data-stu-id="8abc8-128">See also</span></span>
* [<span data-ttu-id="8abc8-129">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8abc8-129">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="8abc8-130">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="8abc8-130">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="8abc8-131">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="8abc8-131">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="8abc8-132">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="8abc8-132">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
