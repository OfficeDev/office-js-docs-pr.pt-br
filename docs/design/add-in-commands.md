---
title: Comandos de suplemento para Excel, Word e PowerPoint
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 698fd4b77ea90430a141db1c791856f4f57fa29b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533662"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="033d9-102">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="033d9-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="033d9-p101">Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. Os comandos de suplemento ajudam os usuários a localizar e usar o suplemento, o que pode ajudá-lo a aumentar a adoção e a reutilização do suplemento, além de melhorar a retenção de clientes.</span><span class="sxs-lookup"><span data-stu-id="033d9-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="033d9-107">Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="033d9-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="033d9-p102">Os catálogos do SharePoint não são compatíveis com os comandos de suplemento. É possível implantar comandos de suplemento pela [Implantação centralizada](../publish/centralized-deployment.md) ou pelo [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store) ou usar [sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para implantar seu comando de suplemento para testes.</span><span class="sxs-lookup"><span data-stu-id="033d9-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="033d9-110">*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="033d9-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Captura de tela de um comando de suplemento no Excel](../images/add-in-commands-1.png)

<span data-ttu-id="033d9-112">*Figura 2. Suplemento com comandos em execução no Excel Online*</span><span class="sxs-lookup"><span data-stu-id="033d9-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Captura de tela de um comando de suplemento no Excel Online](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="033d9-114">Recursos de comandos</span><span class="sxs-lookup"><span data-stu-id="033d9-114">Command capabilities</span></span>
<span data-ttu-id="033d9-115">Os seguintes recursos de comando são compatíveis no momento.</span><span class="sxs-lookup"><span data-stu-id="033d9-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="033d9-116">Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="033d9-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="033d9-117">**Pontos de extensão**</span><span class="sxs-lookup"><span data-stu-id="033d9-117">**Extension points**</span></span>

- <span data-ttu-id="033d9-118">Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="033d9-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="033d9-119">Menus de contexto: estender os menus de contexto selecionados.</span><span class="sxs-lookup"><span data-stu-id="033d9-119">Context menus - Extend selected context menus.</span></span>

<span data-ttu-id="033d9-120">**Tipos de controle**</span><span class="sxs-lookup"><span data-stu-id="033d9-120">**Control types**</span></span>

- <span data-ttu-id="033d9-121">Botões simples: disparar ações específicas.</span><span class="sxs-lookup"><span data-stu-id="033d9-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="033d9-122">Menus – menu suspenso simples com botões que disparam ações.</span><span class="sxs-lookup"><span data-stu-id="033d9-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="033d9-123">**Ações**</span><span class="sxs-lookup"><span data-stu-id="033d9-123">**Actions**</span></span>

- <span data-ttu-id="033d9-124">ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.</span><span class="sxs-lookup"><span data-stu-id="033d9-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="033d9-p103">ExecuteFunction: carrega uma página HTML invisível e executa uma função JavaScript dentro dela. Para mostrar a interface do usuário dentro de sua função (como erros, progresso ou entrada adicional), você pode usar a API [displayDialog](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="033d9-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="033d9-127">Plataformas com suporte</span><span class="sxs-lookup"><span data-stu-id="033d9-127">Supported platforms</span></span>

<span data-ttu-id="033d9-128">Os comandos de suplemento atualmente têm suporte nas seguintes plataformas:</span><span class="sxs-lookup"><span data-stu-id="033d9-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="033d9-129">Outlook 2016 ou posterior para Windows (build 16.0.6769+)</span><span class="sxs-lookup"><span data-stu-id="033d9-129">Office for Windows Desktop 2016 (build 16.0.6769+)</span></span>
- <span data-ttu-id="033d9-130">Office para Mac (build 15.33+)</span><span class="sxs-lookup"><span data-stu-id="033d9-130">Office for Mac (build 15.33+)</span></span>
- <span data-ttu-id="033d9-131">Office Online</span><span class="sxs-lookup"><span data-stu-id="033d9-131">Office Online</span></span>

<span data-ttu-id="033d9-132">Mais plataformas serão incluídas em breve.</span><span class="sxs-lookup"><span data-stu-id="033d9-132">More platforms are coming soon.</span></span>

## <a name="best-practices"></a><span data-ttu-id="033d9-133">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="033d9-133">Best practices</span></span>

<span data-ttu-id="033d9-134">Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="033d9-134">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="033d9-p104">Use os comandos para representar uma ação específica com um resultado claro e específico para os usuários. Não combine várias ações em um único botão.</span><span class="sxs-lookup"><span data-stu-id="033d9-p104">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="033d9-p105">Forneça ações granulares que tornam a realização de tarefas comuns no seu suplemento mais eficiente. Minimize o número de etapas necessárias para concluir uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="033d9-p105">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="033d9-139">Para o posicionamento dos comandos na faixa de opções do Office:</span><span class="sxs-lookup"><span data-stu-id="033d9-139">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="033d9-p106">Insira os comandos em uma guia existente (Inserir, Revisar e assim por diante) se a funcionalidade fornecida se encaixar ali. Por exemplo, se seu suplemento permitir que os usuários insiram mídia, adicione um grupo à guia Inserir. Observe que nem todas as guias estão disponíveis em todas as versões do Office. Para saber mais, confira o [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="033d9-p106">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> 
    - <span data-ttu-id="033d9-p107">Insira comandos na guia Página Inicial se a funcionalidade não se encaixar em outra guia e você menos de seis comandos de nível superior. Você também pode adicionar comandos à guia Página Inicial se seu suplemento precisar funcionar em diferentes versões do Office (como o Office para área de trabalho e o Office Online) e uma guia não estiver disponível em todas as versões (por exemplo, a guia Design não existe no Office Online).</span><span class="sxs-lookup"><span data-stu-id="033d9-p107">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="033d9-145">Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.</span><span class="sxs-lookup"><span data-stu-id="033d9-145">Place commands on a custom tab if you have more than six top-level commands.</span></span> 
    - <span data-ttu-id="033d9-p108">Nomeie seu grupo de acordo com o nome do seu suplemento. Se você tiver vários grupos, nomeie cada grupo com base na funcionalidade que os comandos nesse grupo fornecem.</span><span class="sxs-lookup"><span data-stu-id="033d9-p108">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="033d9-148">Não adicione botões supérfluos para aumentar o estado real do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="033d9-148">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="033d9-149">Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="033d9-149">Add-ins that take up too much space might not pass [AppSource validation](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="033d9-150">Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="033d9-150">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="033d9-151">Forneça uma versão do seu suplemento que também funcione em hosts que não tenham suporte para comandos.</span><span class="sxs-lookup"><span data-stu-id="033d9-151">Provide a version of your add-in that also works on hosts that do not support commands. A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a taskpane) hosts.</span></span> <span data-ttu-id="033d9-152">Um manifesto de suplemento único pode funcionar tanto em hosts cientes do comando (com os comandos) quanto em hosts não cientes do comando (como um painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="033d9-152">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="033d9-153">*Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*</span><span class="sxs-lookup"><span data-stu-id="033d9-153">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Uma captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="033d9-155">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="033d9-155">Next steps</span></span>

<span data-ttu-id="033d9-156">A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="033d9-156">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="033d9-157">Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="033d9-157">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js) reference content.</span></span>
