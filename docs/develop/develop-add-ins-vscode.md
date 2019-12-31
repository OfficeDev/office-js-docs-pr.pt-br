---
title: Desenvolver Suplementos do Office com o Código do Visual Studio
description: Como desenvolver Suplementos do Office com o Código do Visual Studio
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: d6c6cafb28c12b2beb07f0b0cc30d8159e1df1b2
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40914858"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="b8028-103">Desenvolver Suplementos do Office com o Código do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b8028-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="b8028-104">Este artigo descreve como usar [o Código do Visual Studio (VS Code)](https://code.visualstudio.com) para desenvolver um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b8028-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b8028-105">Para saber mais sobre como usar o Visual Studio para criar um suplemento do Office, confira [Desenvolver suplementos do Office com o Visual Studio](develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="b8028-105">For information about using Visual Studio to create an Office Add-in, see [Create and debug Office Add-ins in Visual Studio](develop-add-ins-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b8028-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="b8028-106">Prerequisites</span></span>

- [<span data-ttu-id="b8028-107">Código do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b8028-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="b8028-108">Criar um projeto de suplemento usando o gerador Yeoman</span><span class="sxs-lookup"><span data-stu-id="b8028-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="b8028-109">Se você estiver usando o VS Code como o seu ambiente de desenvolvimento integrado (IDE), crie o projeto do Suplemento do Office com o [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office). O gerador Yeoman cria um projeto Node.js que pode ser gerenciado com o VS Code ou qualquer outro editor.</span><span class="sxs-lookup"><span data-stu-id="b8028-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="b8028-110">Para criar um Suplemento do Office com o gerador Yeoman, siga as instruções em[início rápido em 5 minutos](../index.md) que corresponda ao tipo de suplemento que você deseja criar.</span><span class="sxs-lookup"><span data-stu-id="b8028-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="b8028-111">Desenvolver o suplemento usando o VS Code</span><span class="sxs-lookup"><span data-stu-id="b8028-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="b8028-112">Quando o gerador Yeoman terminar de criar o projeto do suplemento, abra a pasta raiz do projeto com o VS Code.</span><span class="sxs-lookup"><span data-stu-id="b8028-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="b8028-113">No Windows, navegue até o diretório raiz do projeto por meio da linha de comando e, em seguida, insira `code .` para abrir essa pasta no VS Code.</span><span class="sxs-lookup"><span data-stu-id="b8028-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="b8028-114">No Mac, você precisará [adicionar o comando `code` ao caminho](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) antes de poder usá-lo para abrir a pasta do projeto no VS Code.</span><span class="sxs-lookup"><span data-stu-id="b8028-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="b8028-115">O gerador Yeoman cria um suplemento básico com funcionalidade limitada.</span><span class="sxs-lookup"><span data-stu-id="b8028-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="b8028-116">Você pode personalizar o suplemento editando o [manifesto](add-in-manifests.md), HTML, JavaScript ou TypeScript e arquivos CSS no VS Code.</span><span class="sxs-lookup"><span data-stu-id="b8028-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="b8028-117">Para obter uma descrição de alto nível sobre a estrutura e os arquivos do projeto no projeto de suplemento que o gerador de Yeoman cria, confira o tópico diretrizes do gerador Yeoman dentro em [Início rápido em 5 minutos](../index.md) que corresponda ao tipo de suplemento que você criou.</span><span class="sxs-lookup"><span data-stu-id="b8028-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="b8028-118">Testar e depurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="b8028-118">Test and debug the add-in</span></span>

<span data-ttu-id="b8028-119">Os métodos para testar, depurar e solucionar problemas de Suplementos do Office variam de acordo com a plataforma.</span><span class="sxs-lookup"><span data-stu-id="b8028-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="b8028-120">Para mais informações, confira [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b8028-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="b8028-121">Publique o suplemento</span><span class="sxs-lookup"><span data-stu-id="b8028-121">Publish the add-in</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="b8028-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="b8028-122">See also</span></span>

- [<span data-ttu-id="b8028-123">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="b8028-123">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="b8028-124">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8028-124">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="b8028-125">Desenvolver Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8028-125">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="b8028-126">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8028-126">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="b8028-127">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8028-127">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="b8028-128">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8028-128">Publish Office Add-ins</span></span>](../publish/publish.md)