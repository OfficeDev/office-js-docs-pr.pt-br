---
title: Elementos da interface do usuário do Office para suplementos do Office
description: ''
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: f805926894a33f74914fbd315245c93c826d5fc0
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950422"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="f689f-102">Elementos da interface do usuário do Office para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f689f-102">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="f689f-p101">Você pode usar vários tipos de elementos para estender a interface do usuário do Office, incluindo comandos de suplemento e contêineres HTML. Esses elementos de interface do usuário parecem uma extensão natural do Office e funcionam entre plataformas. Você pode inserir um código personalizado baseado na Web em qualquer um desses elementos.</span><span class="sxs-lookup"><span data-stu-id="f689f-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="f689f-106">A imagem a seguir mostra os tipos de elementos de interface do usuário do Office que você pode criar.</span><span class="sxs-lookup"><span data-stu-id="f689f-106">The following image shows the types of Office UI elements that you can create.</span></span>

![Uma imagem que mostra comandos de suplemento na faixa de opções, um painel de tarefas e uma caixa de diálogo em um documento do Office](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="f689f-108">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="f689f-108">Add-in commands</span></span>

<span data-ttu-id="f689f-p102">Use [comandos de suplemento](add-in-commands.md) para adicionar pontos de entrada ao suplemento na faixa de opções do Office. Comandos iniciam ações no suplemento executando código JavaScript ou iniciando um contêiner HTML. Você pode criar dois tipos de comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="f689f-p102">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office ribbon. Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container. You can create two types of add-in commands.</span></span>

|<span data-ttu-id="f689f-112">**Tipo de comando**</span><span class="sxs-lookup"><span data-stu-id="f689f-112">**Command type**</span></span>|<span data-ttu-id="f689f-113">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="f689f-113">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="f689f-114">Botões, menus e guias da faixa de opções</span><span class="sxs-lookup"><span data-stu-id="f689f-114">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="f689f-p103">Use para adicionar botões personalizados, menus (menus suspensos) ou guias à faixa de opções padrão no Office. Use botões e menus para disparar uma ação no Office. Use guias para agrupar e organizar botões e menus.</span><span class="sxs-lookup"><span data-stu-id="f689f-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="f689f-118">Menus de contexto</span><span class="sxs-lookup"><span data-stu-id="f689f-118">Context menus</span></span>| <span data-ttu-id="f689f-p104">Use para estender o menu de contexto padrão. Menus de contexto são exibidos quando os usuários clicam com o botão direito do mouse no texto em um documento do Office ou uma tabela no Excel.</span><span class="sxs-lookup"><span data-stu-id="f689f-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="f689f-121">Contêineres HTML</span><span class="sxs-lookup"><span data-stu-id="f689f-121">HTML containers</span></span>

<span data-ttu-id="f689f-p105">Use contêineres HTML para inserir código de interface do usuário baseado em HTML em clientes Office. Essas páginas da Web podem fazer referência à API do JavaScript do Office para interagir com conteúdo no documento. Você pode criar três tipos de contêineres HTML.</span><span class="sxs-lookup"><span data-stu-id="f689f-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="f689f-125">**Contêiner HTML**</span><span class="sxs-lookup"><span data-stu-id="f689f-125">**HTML container**</span></span>|<span data-ttu-id="f689f-126">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="f689f-126">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="f689f-127">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="f689f-127">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="f689f-p106">Exibir a interface do usuário personalizada no painel à direita do documento do Office. Use os painéis de tarefas para permitir que os usuários interajam com o suplemento lado a lado com o documento do Office.</span><span class="sxs-lookup"><span data-stu-id="f689f-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="f689f-130">Suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="f689f-130">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="f689f-p107">Exibir a interface do usuário personalizada inserida em documentos do Office. Use os suplementos de conteúdo para permitir que os usuários interajam com o suplemento diretamente no documento do Office. Por exemplo, talvez você queira mostrar conteúdo externo, como vídeos ou visualizações de dados de outras fontes.</span><span class="sxs-lookup"><span data-stu-id="f689f-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="f689f-134">Caixas de diálogo</span><span class="sxs-lookup"><span data-stu-id="f689f-134">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="f689f-p108">Exibir uma interface do usuário personalizada em uma caixa de diálogo que se sobrepõe ao documento do Office. Use uma caixa de diálogo para interações que requerem foco e mais espaço, e não exigem uma interação lado a lado com o documento.</span><span class="sxs-lookup"><span data-stu-id="f689f-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="f689f-137">Confira também</span><span class="sxs-lookup"><span data-stu-id="f689f-137">See also</span></span>

- [<span data-ttu-id="f689f-138">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f689f-138">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="f689f-139">Painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="f689f-139">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="f689f-140">Suplementos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="f689f-140">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="f689f-141">Caixas de diálogo</span><span class="sxs-lookup"><span data-stu-id="f689f-141">Dialog boxes</span></span>](dialog-boxes.md)
