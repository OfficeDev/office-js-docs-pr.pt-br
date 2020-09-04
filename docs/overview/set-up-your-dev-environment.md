---
title: Defina seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar suplementos do Office
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 5e7d91d81ef3d124e9582e74151626b9fd65991a
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47363692"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="cb09a-103">Defina seu ambiente de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="cb09a-103">Set up your development environment</span></span>

<span data-ttu-id="cb09a-104">Este guia ajuda você a configurar ferramentas para que você possa criar suplementos do Office seguindo nosso início rápido ou tutoriais.</span><span class="sxs-lookup"><span data-stu-id="cb09a-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="cb09a-105">Você precisará instalar as ferramentas na lista abaixo.</span><span class="sxs-lookup"><span data-stu-id="cb09a-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="cb09a-106">Se você já tiver estes instalados, você está pronto para iniciar um início rápido, como este [início rápido reagir do Excel](../quickstarts/excel-quickstart-react.md).</span><span class="sxs-lookup"><span data-stu-id="cb09a-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="cb09a-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="cb09a-107">Node.js</span></span>
- <span data-ttu-id="cb09a-108">npm</span><span class="sxs-lookup"><span data-stu-id="cb09a-108">npm</span></span>
- <span data-ttu-id="cb09a-109">Uma conta do Microsoft 365 que inclui a versão de assinatura do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="cb09a-110">Um editor de código de sua escolha</span><span class="sxs-lookup"><span data-stu-id="cb09a-110">A code editor of your choice</span></span>

<span data-ttu-id="cb09a-111">Este guia pressupõe que você saiba como usar uma ferramenta de linha de comando.</span><span class="sxs-lookup"><span data-stu-id="cb09a-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="cb09a-112">Instale o Node.js.</span><span class="sxs-lookup"><span data-stu-id="cb09a-112">Install Node.js</span></span>

<span data-ttu-id="cb09a-113">Node.js é um tempo de execução de JavaScript, você precisará desenvolver suplementos do Office modernos.</span><span class="sxs-lookup"><span data-stu-id="cb09a-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="cb09a-114">Instale o Node.js [baixando a versão mais recente recomendada do site](https://nodejs.org).</span><span class="sxs-lookup"><span data-stu-id="cb09a-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="cb09a-115">Siga as instruções de instalação do seu sistema operacional.</span><span class="sxs-lookup"><span data-stu-id="cb09a-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="cb09a-116">Instalar o NPM</span><span class="sxs-lookup"><span data-stu-id="cb09a-116">Install npm</span></span>

<span data-ttu-id="cb09a-117">o NPM é um registro de software de código aberto do qual baixar os pacotes usados no desenvolvimento de suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="cb09a-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="cb09a-118">Para instalar o NPM, execute o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="cb09a-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="cb09a-119">Para verificar se você já tem o NPM instalado e veja a versão instalada, execute o seguinte na linha de comando:</span><span class="sxs-lookup"><span data-stu-id="cb09a-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="cb09a-120">Você pode querer usar um Gerenciador de versão do nó para permitir que você alterne entre várias versões do Node.js e do NPM, mas isso não é estritamente necessário.</span><span class="sxs-lookup"><span data-stu-id="cb09a-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="cb09a-121">Para obter detalhes sobre como fazer isso, [consulte as instruções do NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span><span class="sxs-lookup"><span data-stu-id="cb09a-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="cb09a-122">Obter o Office 365</span><span class="sxs-lookup"><span data-stu-id="cb09a-122">Get Office 365</span></span>

<span data-ttu-id="cb09a-123">Se você ainda não tem uma conta no Microsoft 365, é possível obter uma assinatura gratuita do Microsoft 365 renovável por 90 dias ingressando no [programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="cb09a-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="cb09a-124">Instalar um editor de códigos</span><span class="sxs-lookup"><span data-stu-id="cb09a-124">Install a code editor</span></span>

<span data-ttu-id="cb09a-125">Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:</span><span class="sxs-lookup"><span data-stu-id="cb09a-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="cb09a-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="cb09a-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="cb09a-127">Atom</span><span class="sxs-lookup"><span data-stu-id="cb09a-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="cb09a-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="cb09a-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="cb09a-129">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="cb09a-129">Next steps</span></span>

<span data-ttu-id="cb09a-130">Tente criar seu próprio suplemento ou use o script Lab para experimentar exemplos internos.</span><span class="sxs-lookup"><span data-stu-id="cb09a-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="cb09a-131">Criar um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-131">Create an Office add-in</span></span>

<span data-ttu-id="cb09a-132">Você pode criar rapidamente um suplemento básico para o Excel, o OneNote, o Outlook, o PowerPoint, o Project ou o Word realizando um [início rápido de 5 minutos](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="cb09a-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="cb09a-133">Se você já concluiu um início rápido e deseja criar um suplemento um pouco mais complexo, experiente o [tutorial](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="cb09a-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="cb09a-134">Explorar as APIs com o Script Lab</span><span class="sxs-lookup"><span data-stu-id="cb09a-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="cb09a-135">Explore a biblioteca de amostras internas no [Script Lab](explore-with-script-lab.md) para ter uma ideia dos recursos das APIs JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="cb09a-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="cb09a-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="cb09a-136">See also</span></span>

- [<span data-ttu-id="cb09a-137">Desenvolver suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-137">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="cb09a-138">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="cb09a-139">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="cb09a-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="cb09a-140">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="cb09a-141">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="cb09a-142">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb09a-142">Publish Office Add-ins</span></span>](../publish/publish.md)
