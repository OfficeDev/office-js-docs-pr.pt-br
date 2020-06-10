---
title: Defina seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar suplementos do Office
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: f44f8e48aec402f0ffa6327732613a902ea0cfe6
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679350"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="dc9b1-103">Defina seu ambiente de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="dc9b1-103">Set up your development environment</span></span>

<span data-ttu-id="dc9b1-104">Este guia ajuda você a configurar ferramentas para que você possa criar suplementos do Office seguindo nosso início rápido ou tutoriais.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="dc9b1-105">Você precisará instalar as ferramentas na lista abaixo.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="dc9b1-106">Se você já tiver estes instalados, você está pronto para iniciar um início rápido, como este [início rápido reagir do Excel](../quickstarts/excel-quickstart-react.md).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="dc9b1-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="dc9b1-107">Node.js</span></span>
- <span data-ttu-id="dc9b1-108">npm</span><span class="sxs-lookup"><span data-stu-id="dc9b1-108">npm</span></span>
- <span data-ttu-id="dc9b1-109">Uma conta do Office 365 (a versão de assinatura do Office)</span><span class="sxs-lookup"><span data-stu-id="dc9b1-109">An Office 365 (the subscription version of Office) account</span></span>
- <span data-ttu-id="dc9b1-110">Um editor de código de sua escolha</span><span class="sxs-lookup"><span data-stu-id="dc9b1-110">A code editor of your choice</span></span>

<span data-ttu-id="dc9b1-111">Este guia pressupõe que você saiba como usar uma ferramenta de linha de comando.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="dc9b1-112">Instalar o Node. js</span><span class="sxs-lookup"><span data-stu-id="dc9b1-112">Install Node.js</span></span>

<span data-ttu-id="dc9b1-113">Node. js é um tempo de execução de JavaScript, você precisará desenvolver suplementos do Office modernos.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="dc9b1-114">Instale o Node. js [baixando a versão mais recente recomendada do site](https://nodejs.org).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="dc9b1-115">Siga as instruções de instalação do seu sistema operacional.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="dc9b1-116">Instalar o NPM</span><span class="sxs-lookup"><span data-stu-id="dc9b1-116">Install npm</span></span>

<span data-ttu-id="dc9b1-117">o NPM é um registro de software de código aberto do qual baixar os pacotes usados no desenvolvimento de suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="dc9b1-118">Para instalar o NPM, execute o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="dc9b1-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="dc9b1-119">Para verificar se você já tem o NPM instalado e veja a versão instalada, execute o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="dc9b1-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="dc9b1-120">Você pode querer usar um Gerenciador de versão do nó para permitir que você alterne entre várias versões do node. js e do NPM, mas isso não é estritamente necessário.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="dc9b1-121">Para obter detalhes sobre como fazer isso, [consulte as instruções do NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="dc9b1-122">Obter o Office 365</span><span class="sxs-lookup"><span data-stu-id="dc9b1-122">Get Office 365</span></span>

<span data-ttu-id="dc9b1-123">Caso ainda não tenha uma conta do Office 365, obtenha uma assinatura gratuita renovável por 90 dias do Office 365 ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-123">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="dc9b1-124">Instalar um editor de códigos</span><span class="sxs-lookup"><span data-stu-id="dc9b1-124">Install a code editor</span></span>

<span data-ttu-id="dc9b1-125">Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:</span><span class="sxs-lookup"><span data-stu-id="dc9b1-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="dc9b1-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="dc9b1-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="dc9b1-127">Atom</span><span class="sxs-lookup"><span data-stu-id="dc9b1-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="dc9b1-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="dc9b1-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="dc9b1-129">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="dc9b1-129">Next steps</span></span>

<span data-ttu-id="dc9b1-130">Tente criar seu próprio suplemento ou use o script Lab para experimentar exemplos internos.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="dc9b1-131">Criar um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="dc9b1-131">Create an Office add-in</span></span>

<span data-ttu-id="dc9b1-132">Você pode criar rapidamente um suplemento básico para o Excel, o OneNote, o Outlook, o PowerPoint, o Project ou o Word realizando um [início rápido de 5 minutos](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="dc9b1-133">Se você já concluiu um início rápido e deseja criar um suplemento um pouco mais complexo, experiente o [tutorial](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="dc9b1-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="dc9b1-134">Explorar as APIs com o Script Lab</span><span class="sxs-lookup"><span data-stu-id="dc9b1-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="dc9b1-135">Explore a biblioteca de amostras internas no [Script Lab](explore-with-script-lab.md) para ter uma ideia dos recursos das APIs JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="dc9b1-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="dc9b1-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="dc9b1-136">See also</span></span>

- [<span data-ttu-id="dc9b1-137">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="dc9b1-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="dc9b1-138">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="dc9b1-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="dc9b1-139">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="dc9b1-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="dc9b1-140">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="dc9b1-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="dc9b1-141">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="dc9b1-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="dc9b1-142">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="dc9b1-142">Publish Office Add-ins</span></span>](../publish/publish.md)
