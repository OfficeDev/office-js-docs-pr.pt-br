---
title: Comandos de suplemento para Excel, Word e PowerPoint
description: ''
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 7b85d3016b195b353b1e7f314aceb761cf4e31b3
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952177"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="52a41-102">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="52a41-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="52a41-p101">Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. Os comandos de suplemento ajudam os usuários a localizar e usar o suplemento, o que pode ajudá-lo a aumentar a adoção e a reutilização do suplemento, além de melhorar a retenção de clientes.</span><span class="sxs-lookup"><span data-stu-id="52a41-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="52a41-107">Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="52a41-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="52a41-p102">Os catálogos do SharePoint não são compatíveis com os comandos de suplemento. É possível implantar comandos de suplemento pela [Implantação centralizada](../publish/centralized-deployment.md) ou pelo [AppSource](/office/dev/store/submit-to-the-office-store) ou usar [sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para implantar seu comando de suplemento para testes.</span><span class="sxs-lookup"><span data-stu-id="52a41-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="52a41-110">*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="52a41-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Captura de tela de um comando de suplemento no Excel](../images/add-in-commands-1.png)

<span data-ttu-id="52a41-112">*Figura 2. Suplemento com comandos em execução no Excel Online*</span><span class="sxs-lookup"><span data-stu-id="52a41-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Captura de tela de um comando de suplemento no Excel Online](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="52a41-114">Recursos de comandos</span><span class="sxs-lookup"><span data-stu-id="52a41-114">Command capabilities</span></span>

<span data-ttu-id="52a41-115">Os seguintes recursos de comando são compatíveis no momento.</span><span class="sxs-lookup"><span data-stu-id="52a41-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="52a41-116">Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="52a41-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="52a41-117">**Pontos de extensão**</span><span class="sxs-lookup"><span data-stu-id="52a41-117">**Extension points**</span></span>

- <span data-ttu-id="52a41-118">Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="52a41-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="52a41-119">Menus de contexto: estender os menus de contexto selecionados.</span><span class="sxs-lookup"><span data-stu-id="52a41-119">Context menus - Extend selected context menus.</span></span>

<span data-ttu-id="52a41-120">**Tipos de controle**</span><span class="sxs-lookup"><span data-stu-id="52a41-120">**Control types**</span></span>

- <span data-ttu-id="52a41-121">Botões simples: disparar ações específicas.</span><span class="sxs-lookup"><span data-stu-id="52a41-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="52a41-122">Menus – menu suspenso simples com botões que disparam ações.</span><span class="sxs-lookup"><span data-stu-id="52a41-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="52a41-123">**Ações**</span><span class="sxs-lookup"><span data-stu-id="52a41-123">**Actions**</span></span>

- <span data-ttu-id="52a41-124">ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.</span><span class="sxs-lookup"><span data-stu-id="52a41-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="52a41-p103">ExecuteFunction: carrega uma página HTML invisível e executa uma função JavaScript dentro dela. Para mostrar a interface do usuário dentro de sua função (como erros, progresso ou entrada adicional), você pode usar a API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="52a41-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="52a41-127">Plataformas com suporte</span><span class="sxs-lookup"><span data-stu-id="52a41-127">Supported platforms</span></span>

<span data-ttu-id="52a41-128">Os comandos de suplemento atualmente têm suporte nas seguintes plataformas:</span><span class="sxs-lookup"><span data-stu-id="52a41-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="52a41-129">Outlook 2016 no Windows (build 16.0.4678.1000+)</span><span class="sxs-lookup"><span data-stu-id="52a41-129">Outlook 2016 on Windows (build 16.0.4678.1000+)</span></span>
- <span data-ttu-id="52a41-130">Office no Windows conectado ao Office 365 (build 16.0.6769+)</span><span class="sxs-lookup"><span data-stu-id="52a41-130">Office on Windows connected to Office 365 (build 16.0.6769+)</span></span>
- <span data-ttu-id="52a41-131">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="52a41-131">Office 2019 for Windows</span></span>
- <span data-ttu-id="52a41-132">Office para Mac conectado ao Office 365 (build 15.33+)</span><span class="sxs-lookup"><span data-stu-id="52a41-132">Office for Mac connected to Office 365 (build 15.33+)</span></span>
- <span data-ttu-id="52a41-133">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="52a41-133">Office 2019 for Mac</span></span>
- <span data-ttu-id="52a41-134">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a41-134">Office Online</span></span>

<span data-ttu-id="52a41-135">Mais plataformas serão incluídas em breve.</span><span class="sxs-lookup"><span data-stu-id="52a41-135">More platforms are coming soon.</span></span>

## <a name="debugging"></a><span data-ttu-id="52a41-136">Depuração</span><span class="sxs-lookup"><span data-stu-id="52a41-136">Debugging</span></span>

<span data-ttu-id="52a41-137">Para depurar um comando de Suplemento, você deve executá-lo no Office Online.</span><span class="sxs-lookup"><span data-stu-id="52a41-137">To debug an Add-in Command, you must run it in Office Online.</span></span> <span data-ttu-id="52a41-138">Para obter detalhes, consulte [Depurar suplementos no Office Online](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="52a41-138">For details, see [Debug add-ins in Office Online](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="52a41-139">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="52a41-139">Best practices</span></span>

<span data-ttu-id="52a41-140">Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="52a41-140">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="52a41-p105">Use os comandos para representar uma ação específica com um resultado claro e específico para os usuários. Não combine várias ações em um único botão.</span><span class="sxs-lookup"><span data-stu-id="52a41-p105">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="52a41-p106">Forneça ações granulares que tornam a realização de tarefas comuns no seu suplemento mais eficiente. Minimize o número de etapas necessárias para concluir uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="52a41-p106">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="52a41-145">Para o posicionamento dos comandos na faixa de opções do Office:</span><span class="sxs-lookup"><span data-stu-id="52a41-145">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="52a41-p107">Insira os comandos em uma guia existente (Inserir, Revisar e assim por diante) se a funcionalidade fornecida se encaixar ali. Por exemplo, se seu suplemento permitir que os usuários insiram mídia, adicione um grupo à guia Inserir. Observe que nem todas as guias estão disponíveis em todas as versões do Office. Para saber mais, confira o [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="52a41-p107">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="52a41-p108">Insira comandos na guia Página Inicial se a funcionalidade não se encaixar em outra guia e você menos de seis comandos de nível superior. Você também pode adicionar comandos à guia Página Inicial se seu suplemento precisar funcionar em diferentes versões do Office (como o Office para área de trabalho e o Office Online) e uma guia não estiver disponível em todas as versões (por exemplo, a guia Design não existe no Office Online).</span><span class="sxs-lookup"><span data-stu-id="52a41-p108">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="52a41-151">Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.</span><span class="sxs-lookup"><span data-stu-id="52a41-151">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="52a41-p109">Nomeie seu grupo de acordo com o nome do seu suplemento. Se você tiver vários grupos, nomeie cada grupo com base na funcionalidade que os comandos nesse grupo fornecem.</span><span class="sxs-lookup"><span data-stu-id="52a41-p109">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="52a41-154">Não adicione botões supérfluos para aumentar o estado real do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="52a41-154">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="52a41-155">Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="52a41-155">Add-ins that take up too much space might not pass [AppSource validation](/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="52a41-156">Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="52a41-156">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="52a41-157">Forneça uma versão do seu suplemento que também funcione em hosts que não tenham suporte para comandos.</span><span class="sxs-lookup"><span data-stu-id="52a41-157">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="52a41-158">Um manifesto de suplemento único pode funcionar tanto em hosts cientes do comando (com os comandos) quanto em hosts não cientes do comando (como um painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="52a41-158">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="52a41-159">*Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*</span><span class="sxs-lookup"><span data-stu-id="52a41-159">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Uma captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="52a41-161">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="52a41-161">Next steps</span></span>

<span data-ttu-id="52a41-162">A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="52a41-162">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="52a41-163">Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides).</span><span class="sxs-lookup"><span data-stu-id="52a41-163">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) reference content.</span></span>
