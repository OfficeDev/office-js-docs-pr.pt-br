---
title: Conceitos básicos para comandos de suplemento
description: Aprenda a adicionar botões e itens de menu personalizados da faixa de opções ao Office como parte de um suplemento do Office.
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093872"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="698c9-103">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="698c9-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="698c9-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span><span class="sxs-lookup"><span data-stu-id="698c9-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span></span> <span data-ttu-id="698c9-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span><span class="sxs-lookup"><span data-stu-id="698c9-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span></span> <span data-ttu-id="698c9-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span><span class="sxs-lookup"><span data-stu-id="698c9-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> <span data-ttu-id="698c9-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span><span class="sxs-lookup"><span data-stu-id="698c9-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="698c9-108">Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Aplicativo do Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="698c9-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="698c9-109">SharePoint catalogs do not support add-in commands.</span><span class="sxs-lookup"><span data-stu-id="698c9-109">SharePoint catalogs do not support add-in commands.</span></span> <span data-ttu-id="698c9-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span><span class="sxs-lookup"><span data-stu-id="698c9-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="698c9-111">Os comandos de suplemento também são compatíveis com o Outlook.</span><span class="sxs-lookup"><span data-stu-id="698c9-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="698c9-112">Para saber mais, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="698c9-113">*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="698c9-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Captura de tela de um comando de suplemento no Excel](../images/add-in-commands-1.png)

<span data-ttu-id="698c9-115">*Figura 2. Suplemento com comandos em execução no Excel na Web*</span><span class="sxs-lookup"><span data-stu-id="698c9-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Captura de tela de um comando de suplemento no Excel na Web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="698c9-117">Recursos de comandos</span><span class="sxs-lookup"><span data-stu-id="698c9-117">Command capabilities</span></span>

<span data-ttu-id="698c9-118">Os seguintes recursos de comando são compatíveis no momento.</span><span class="sxs-lookup"><span data-stu-id="698c9-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="698c9-119">Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="698c9-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="698c9-120">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="698c9-120">Extension points</span></span>

- <span data-ttu-id="698c9-121">Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="698c9-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="698c9-122">Menus de contexto: estender menus de contexto selecionados.</span><span class="sxs-lookup"><span data-stu-id="698c9-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="698c9-123">Tipos de controle</span><span class="sxs-lookup"><span data-stu-id="698c9-123">Control types</span></span>

- <span data-ttu-id="698c9-124">Botões simples: disparar ações específicas.</span><span class="sxs-lookup"><span data-stu-id="698c9-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="698c9-125">Menus – menu suspenso simples com botões que disparam ações.</span><span class="sxs-lookup"><span data-stu-id="698c9-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="698c9-126">Ações</span><span class="sxs-lookup"><span data-stu-id="698c9-126">Actions</span></span>

- <span data-ttu-id="698c9-127">ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.</span><span class="sxs-lookup"><span data-stu-id="698c9-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="698c9-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span><span class="sxs-lookup"><span data-stu-id="698c9-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span></span> <span data-ttu-id="698c9-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span><span class="sxs-lookup"><span data-stu-id="698c9-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="698c9-130">Status padrão Habilitado ou Desabilitado (visualização)</span><span class="sxs-lookup"><span data-stu-id="698c9-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="698c9-131">Você pode especificar se o comando está ativado ou desativado quando o suplemento é iniciado e alterar programaticamente a configuração.</span><span class="sxs-lookup"><span data-stu-id="698c9-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="698c9-132">Esse recurso está em visualização e não tem suporte em todos os hosts ou cenários.</span><span class="sxs-lookup"><span data-stu-id="698c9-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="698c9-133">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="698c9-134">Plataformas compatíveis</span><span class="sxs-lookup"><span data-stu-id="698c9-134">Supported platforms</span></span>

<span data-ttu-id="698c9-135">Os comandos de suplemento atualmente têm suporte nas seguintes plataformas.</span><span class="sxs-lookup"><span data-stu-id="698c9-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="698c9-136">Office no Windows (Build 16.0.6769+, conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="698c9-136">Office on Windows (build 16.0.6769+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="698c9-137">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="698c9-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="698c9-138">Office no Windows (Build 15.33+, conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="698c9-138">Office on Mac (build 15.33+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="698c9-139">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="698c9-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="698c9-140">Office na Web</span><span class="sxs-lookup"><span data-stu-id="698c9-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="698c9-141">Para saber mais sobre o suporte do Outlook, confira [comandos de suplemento do Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="698c9-142">Depuração</span><span class="sxs-lookup"><span data-stu-id="698c9-142">Debugging</span></span>

<span data-ttu-id="698c9-143">Para depurar um comando de Suplemento, você deve executá-lo no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="698c9-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="698c9-144">Para obter detalhes, confira [Depurar suplementos no Office na Web](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="698c9-145">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="698c9-145">Best practices</span></span>

<span data-ttu-id="698c9-146">Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="698c9-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="698c9-147">Use commands to represent a specific action with a clear and specific outcome for users.</span><span class="sxs-lookup"><span data-stu-id="698c9-147">Use commands to represent a specific action with a clear and specific outcome for users.</span></span> <span data-ttu-id="698c9-148">Do not combine multiple actions in a single button.</span><span class="sxs-lookup"><span data-stu-id="698c9-148">Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="698c9-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span><span class="sxs-lookup"><span data-stu-id="698c9-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span></span> <span data-ttu-id="698c9-150">Minimize the number of steps an action takes to complete.</span><span class="sxs-lookup"><span data-stu-id="698c9-150">Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="698c9-151">Para o posicionamento dos comandos na faixa de opções do Aplicativo do Office:</span><span class="sxs-lookup"><span data-stu-id="698c9-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="698c9-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span><span class="sxs-lookup"><span data-stu-id="698c9-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span></span> <span data-ttu-id="698c9-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span><span class="sxs-lookup"><span data-stu-id="698c9-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span></span> <span data-ttu-id="698c9-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="698c9-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span><span class="sxs-lookup"><span data-stu-id="698c9-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span></span> <span data-ttu-id="698c9-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span><span class="sxs-lookup"><span data-stu-id="698c9-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="698c9-157">Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.</span><span class="sxs-lookup"><span data-stu-id="698c9-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="698c9-158">Name your group to match the name of your add-in.</span><span class="sxs-lookup"><span data-stu-id="698c9-158">Name your group to match the name of your add-in.</span></span> <span data-ttu-id="698c9-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span><span class="sxs-lookup"><span data-stu-id="698c9-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="698c9-160">Não adicione botões supérfluos para aumentar o estado real do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="698c9-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="698c9-161">Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="698c9-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="698c9-162">Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="698c9-163">Forneça uma versão do seu suplemento que também funcione em hosts que não tenham suporte para comandos.</span><span class="sxs-lookup"><span data-stu-id="698c9-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="698c9-164">Um manifesto de suplemento único pode funcionar tanto em hosts cientes do comando (com os comandos) quanto em hosts não cientes do comando (como um painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="698c9-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="698c9-165">*Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*</span><span class="sxs-lookup"><span data-stu-id="698c9-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Uma captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="698c9-167">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="698c9-167">Next steps</span></span>

<span data-ttu-id="698c9-168">A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="698c9-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="698c9-169">Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="698c9-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
