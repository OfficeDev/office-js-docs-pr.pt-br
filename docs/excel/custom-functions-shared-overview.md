---
ms.date: 05/17/2020
description: Aprenda a executar funções personalizadas, botões da faixa de opções e código do painel de tarefas no mesmo tempo de execução do JavaScript para coordenar cenários em seu suplemento.
title: Executar o código do suplemento em um tempo de execução de JavaScript compartilhado
localization_priority: Priority
ms.openlocfilehash: e179d8cdff06b397c9998a6a18a506a0a6df1b42
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609223"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a><span data-ttu-id="d614e-103">Visão geral: executar o código do suplemento em um ambiente de tempo de execução JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="d614e-103">Overview: Run your add-in code in a shared JavaScript runtimes</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="d614e-104">Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados.</span><span class="sxs-lookup"><span data-stu-id="d614e-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="d614e-105">Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.</span><span class="sxs-lookup"><span data-stu-id="d614e-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="d614e-106">No entanto, você pode configurar o suplemento do Excel para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado).</span><span class="sxs-lookup"><span data-stu-id="d614e-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="d614e-107">Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d614e-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="d614e-108">A configuração de um tempo de execução compartilhado permite os seguintes cenários:</span><span class="sxs-lookup"><span data-stu-id="d614e-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="d614e-109">Seu suplemento terá um DOM compartilhado que a faixa de opções, o painel de tarefas e as funções personalizadas podem acessar.</span><span class="sxs-lookup"><span data-stu-id="d614e-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="d614e-110">Suas funções personalizadas terão suporte completo ao CORS.</span><span class="sxs-lookup"><span data-stu-id="d614e-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="d614e-111">Suas funções personalizadas podem chamar as APIs do Office.js para ler os dados do documento da planilha.</span><span class="sxs-lookup"><span data-stu-id="d614e-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="d614e-112">Seu suplemento pode executar o código assim que o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="d614e-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="d614e-113">Seu suplemento pode continuar executando o código após o fechamento do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="d614e-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="d614e-114">Quando você executa funções personalizadas em um tempo de execução compartilhado com o painel de tarefas, ele será executado em uma instância do navegador em plataformas diferentes, conforme explicado em [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, todos os botões exibidos pelo suplemento do Excel na faixa de opções serão executados no mesmo tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="d614e-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="d614e-115">A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d614e-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Funções personalizadas em execução em um tempo de execução compartilhado com botões de faixa de opções e o painel de tarefas no Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a><span data-ttu-id="d614e-117">Configurar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="d614e-117">Set up a shared runtime</span></span>

<span data-ttu-id="d614e-118">Consulte o [artigo Configurando um tempo de execução compartilhado](./configure-your-add-in-to-use-a-shared-runtime.md) para saber como configurar suas funções personalizadas para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="d614e-118">See the [configuring a shared runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.</span></span>

### <a name="debugging"></a><span data-ttu-id="d614e-119">Depuração</span><span class="sxs-lookup"><span data-stu-id="d614e-119">Debugging</span></span>

<span data-ttu-id="d614e-120">Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento.</span><span class="sxs-lookup"><span data-stu-id="d614e-120">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="d614e-121">Em vez disso, você precisará usar as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="d614e-121">You'll need to use developer tools instead.</span></span> <span data-ttu-id="d614e-122">Para obter mais informações, consulte [Depurar suplementos usando ferramentas de desenvolvedor no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="d614e-122">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="d614e-123">Envie-nos seus comentários</span><span class="sxs-lookup"><span data-stu-id="d614e-123">Give us feedback</span></span>

<span data-ttu-id="d614e-124">Adoraríamos ouvir seus comentários sobre esse recurso.</span><span class="sxs-lookup"><span data-stu-id="d614e-124">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="d614e-125">Se você encontrar algum bug ou problema, ou tiver solicitações sobre esse recurso, informe-nos criando um problema do GitHub no [repositório office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="d614e-125">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="d614e-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="d614e-126">See also</span></span>

- [<span data-ttu-id="d614e-127">Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d614e-127">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="d614e-128">Chamar as APIs do Excel a partir de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="d614e-128">Call Excel APIs from your custom function</span></span>](call-excel-apis-from-custom-function.md)
