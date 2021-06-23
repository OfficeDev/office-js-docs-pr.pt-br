---
title: Elementos da interface do usuário do Office para suplementos do Office
description: Obter uma visão geral dos diferentes tipos de elementos da interface do usuário em um Office Add-in.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5d0a1576d850f2291c28e6bb39554cbb0403f50b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076326"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="bb76b-103">Elementos da interface do usuário do Office para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bb76b-103">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="bb76b-p101">Você pode usar vários tipos de elementos para estender a interface do usuário do Office, incluindo comandos de suplemento e contêineres HTML. Esses elementos de interface do usuário parecem uma extensão natural do Office e funcionam entre plataformas. Você pode inserir um código personalizado baseado na Web em qualquer um desses elementos.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="bb76b-107">A imagem a seguir mostra os tipos de elementos de interface do usuário do Office que você pode criar.</span><span class="sxs-lookup"><span data-stu-id="bb76b-107">The following image shows the types of Office UI elements that you can create.</span></span>

![Diagrama mostrando comandos de add-in na faixa de opções, um painel de tarefas e uma caixa de diálogo/um Office de conteúdo.](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="bb76b-109">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="bb76b-109">Add-in commands</span></span>

<span data-ttu-id="bb76b-110">Use [comandos de add-in](add-in-commands.md) para adicionar pontos de entrada ao seu add-in à faixa Aplicativo do Office faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="bb76b-110">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office app ribbon.</span></span> <span data-ttu-id="bb76b-111">Comandos iniciam ações no suplemento executando código JavaScript ou iniciando um contêiner HTML.</span><span class="sxs-lookup"><span data-stu-id="bb76b-111">Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container.</span></span> <span data-ttu-id="bb76b-112">Você pode criar dois tipos de comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="bb76b-112">You can create two types of add-in commands.</span></span>

|<span data-ttu-id="bb76b-113">Tipo de comando</span><span class="sxs-lookup"><span data-stu-id="bb76b-113">Command type</span></span>|<span data-ttu-id="bb76b-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="bb76b-114">Description</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="bb76b-115">Botões, menus e guias da faixa de opções</span><span class="sxs-lookup"><span data-stu-id="bb76b-115">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="bb76b-p103">Use para adicionar botões personalizados, menus (menus suspensos) ou guias à faixa de opções padrão no Office. Use botões e menus para disparar uma ação no Office. Use guias para agrupar e organizar botões e menus.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="bb76b-119">Menus de contexto</span><span class="sxs-lookup"><span data-stu-id="bb76b-119">Context menus</span></span>| <span data-ttu-id="bb76b-p104">Use para estender o menu de contexto padrão. Menus de contexto são exibidos quando os usuários clicam com o botão direito do mouse no texto em um documento do Office ou uma tabela no Excel.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>|

## <a name="html-containers"></a><span data-ttu-id="bb76b-122">Contêineres HTML</span><span class="sxs-lookup"><span data-stu-id="bb76b-122">HTML containers</span></span>

<span data-ttu-id="bb76b-p105">Use contêineres HTML para inserir código de interface do usuário baseado em HTML em clientes Office. Essas páginas da Web podem fazer referência à API do JavaScript do Office para interagir com conteúdo no documento. Você pode criar três tipos de contêineres HTML.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="bb76b-126">Contêiner HTML</span><span class="sxs-lookup"><span data-stu-id="bb76b-126">HTML container</span></span>|<span data-ttu-id="bb76b-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="bb76b-127">Description</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="bb76b-128">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="bb76b-128">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="bb76b-p106">Exibir a interface do usuário personalizada no painel à direita do documento do Office. Use os painéis de tarefas para permitir que os usuários interajam com o suplemento lado a lado com o documento do Office.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="bb76b-131">Suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="bb76b-131">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="bb76b-p107">Exibir a interface do usuário personalizada inserida em documentos do Office. Use os suplementos de conteúdo para permitir que os usuários interajam com o suplemento diretamente no documento do Office. Por exemplo, talvez você queira mostrar conteúdo externo, como vídeos ou visualizações de dados de outras fontes.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="bb76b-135">Caixas de diálogo</span><span class="sxs-lookup"><span data-stu-id="bb76b-135">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="bb76b-p108">Exibir uma interface do usuário personalizada em uma caixa de diálogo que se sobrepõe ao documento do Office. Use uma caixa de diálogo para interações que requerem foco e mais espaço, e não exigem uma interação lado a lado com o documento.</span><span class="sxs-lookup"><span data-stu-id="bb76b-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="bb76b-138">Confira também</span><span class="sxs-lookup"><span data-stu-id="bb76b-138">See also</span></span>

- [<span data-ttu-id="bb76b-139">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bb76b-139">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="bb76b-140">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="bb76b-140">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="bb76b-141">Suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="bb76b-141">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="bb76b-142">Caixas de diálogo</span><span class="sxs-lookup"><span data-stu-id="bb76b-142">Dialog boxes</span></span>](dialog-boxes.md)
