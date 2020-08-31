---
title: 'Desenvolver Suplementos do Office '
description: Uma introdução ao desenvolvimento de Suplementos do Office.
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 419880e8872df20be5a3de40f480f70be2b18859
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292775"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="de4dc-103">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="de4dc-103">Develop Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="de4dc-104">Examine [Criação de Suplementos do Office](../overview/office-add-ins-fundamentals.md) antes de ler este artigo.</span><span class="sxs-lookup"><span data-stu-id="de4dc-104">Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.</span></span>

<span data-ttu-id="de4dc-105">Todos os Suplementos do Office são criados com base na plataforma de Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="de4dc-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="de4dc-106">Eles compartilham uma estrutura comum por meio da qual certas funcionalidades podem ser implementadas.</span><span class="sxs-lookup"><span data-stu-id="de4dc-106">They share a common framework through which certain capabilities can be implemented.</span></span> <span data-ttu-id="de4dc-107">Para qualquer suplemento que você crie, você precisará entender conceitos importantes como a disponibilidade do aplicativo e da plataforma, os padrões de programação da API do Office JavaScript, como especificar as configurações e os recursos do suplemento no arquivo de manifesto e muito mais.</span><span class="sxs-lookup"><span data-stu-id="de4dc-107">For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, and more.</span></span> <span data-ttu-id="de4dc-108">Os principais conceitos de desenvolvimento, como estes mencionados acima, são abordados aqui na seção **Conceitos básicos** > **Desenvolver** dessa documentação.</span><span class="sxs-lookup"><span data-stu-id="de4dc-108">Core development concepts like these are covered here in the **Core concepts** > **Develop** section of the documentation.</span></span> <span data-ttu-id="de4dc-109">Releia as informações contidas aqui antes de explorar a documentação específica do aplicativo que corresponde ao suplemento que você está criando (por exemplo, [Excel](../excel/index.yml)).</span><span class="sxs-lookup"><span data-stu-id="de4dc-109">Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

> [!NOTE]
> <span data-ttu-id="de4dc-110">A seção **Conceitos básicos** > **Desenvolver** > **Como** desta documentação contém artigos voltados para tarefas ou conceitos específicos de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="de4dc-110">The **Core concepts** > **Develop** > **How to** section of this documentation contains articles focused on specific development concepts or tasks.</span></span> <span data-ttu-id="de4dc-111">Por exemplo, você encontrará informações sobre tarefas como [desenvolvendo suplementos com o código do Visual Studio](develop-add-ins-vscode.md), [abrir automaticamente um painel de tarefas com um documento](automatically-open-a-task-pane-with-a-document.md), [criar comandos de suplemento](create-addin-commands.md)e [abrir uma caixa de diálogo](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="de4dc-111">For example, you'll find information there about tasks like [developing add-ins with Visual Studio Code](develop-add-ins-vscode.md), [automatically opening a task pane with a document](automatically-open-a-task-pane-with-a-document.md), [creating add-in commands](create-addin-commands.md), and [opening a dialog box](dialog-api-in-office-add-ins.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="de4dc-112">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="de4dc-112">Next steps</span></span>

<span data-ttu-id="de4dc-113">Depois de se familiarizar com os conceitos básicos abordados aqui, explore a documentação específica do aplicativo que corresponde ao suplemento que você está criando (por exemplo, [Excel](../excel/index.yml)).</span><span class="sxs-lookup"><span data-stu-id="de4dc-113">After you're familiar with the core concepts covered here, explore the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span> <span data-ttu-id="de4dc-114">Cada seção específica da documentação sobre o aplicativo contém informações específicas sobre a criação de suplementos para um determinado aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="de4dc-114">Each application-specific section of the documentation contains information specifically about building add-ins for a certain Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="de4dc-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="de4dc-115">See also</span></span>

- [<span data-ttu-id="de4dc-116">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de4dc-116">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="de4dc-117">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="de4dc-117">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="de4dc-118">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de4dc-118">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="de4dc-119">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de4dc-119">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="de4dc-120">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de4dc-120">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="de4dc-121">Publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="de4dc-121">Publish Office Add-ins</span></span>](../publish/publish.md)
