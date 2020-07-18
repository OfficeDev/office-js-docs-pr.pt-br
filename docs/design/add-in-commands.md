---
title: Conceitos básicos para comandos de suplemento
description: Aprenda a adicionar botões e itens de menu personalizados da faixa de opções ao Office como parte de um suplemento do Office.
ms.date: 07/10/2020
localization_priority: Priority
ms.openlocfilehash: 2c4731b773a20c666ed78eba7e10f59bf9404bfe
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159623"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="7f5f1-103">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7f5f1-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="7f5f1-p101">Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. Os comandos de suplemento ajudam os usuários a localizar e usar o suplemento, o que pode ajudá-lo a aumentar a adoção e a reutilização do suplemento, além de melhorar a retenção de clientes.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="7f5f1-108">Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Aplicativo do Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="7f5f1-p102">Os catálogos do SharePoint não são compatíveis com os comandos de suplemento. É possível implantar comandos de suplemento pela [Implantação centralizada](../publish/centralized-deployment.md) ou pelo [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ou usar [sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para implantar seu comando de suplemento para testes.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7f5f1-111">Os comandos de suplemento também são compatíveis com o Outlook.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="7f5f1-112">Para saber mais, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="7f5f1-113">*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="7f5f1-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Captura de tela de um comando de suplemento no Excel](../images/add-in-commands-1.png)

<span data-ttu-id="7f5f1-115">*Figura 2. Suplemento com comandos em execução no Excel na Web*</span><span class="sxs-lookup"><span data-stu-id="7f5f1-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Captura de tela de um comando de suplemento no Excel na Web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="7f5f1-117">Recursos de comandos</span><span class="sxs-lookup"><span data-stu-id="7f5f1-117">Command capabilities</span></span>

<span data-ttu-id="7f5f1-118">Os seguintes recursos de comando são compatíveis no momento.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="7f5f1-119">Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="7f5f1-120">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="7f5f1-120">Extension points</span></span>

- <span data-ttu-id="7f5f1-121">Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="7f5f1-122">Menus de contexto: estender menus de contexto selecionados.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="7f5f1-123">Tipos de controle</span><span class="sxs-lookup"><span data-stu-id="7f5f1-123">Control types</span></span>

- <span data-ttu-id="7f5f1-124">Botões simples: disparar ações específicas.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="7f5f1-125">Menus – menu suspenso simples com botões que disparam ações.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="7f5f1-126">Ações</span><span class="sxs-lookup"><span data-stu-id="7f5f1-126">Actions</span></span>

- <span data-ttu-id="7f5f1-127">ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="7f5f1-p104">ExecuteFunction: carrega uma página HTML invisível e executa uma função JavaScript dentro dela. Para mostrar a interface do usuário dentro de sua função (como erros, progresso ou entrada adicional), você pode usar a API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="7f5f1-130">Status padrão Habilitado ou Desabilitado (visualização)</span><span class="sxs-lookup"><span data-stu-id="7f5f1-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="7f5f1-131">Você pode especificar se o comando está ativado ou desativado quando o suplemento é iniciado e alterar programaticamente a configuração.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="7f5f1-132">Esse recurso está em visualização e não tem suporte em todos os hosts ou cenários.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="7f5f1-133">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="7f5f1-134">Plataformas compatíveis</span><span class="sxs-lookup"><span data-stu-id="7f5f1-134">Supported platforms</span></span>

<span data-ttu-id="7f5f1-135">Os comandos de suplemento atualmente têm suporte nas seguintes plataformas.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="7f5f1-136">Office no Windows (Build 16.0.6769 ou superior, conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="7f5f1-136">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="7f5f1-137">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="7f5f1-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="7f5f1-138">Office no Mac (build 15.33 ou superior, conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="7f5f1-138">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="7f5f1-139">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="7f5f1-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="7f5f1-140">Office na Web</span><span class="sxs-lookup"><span data-stu-id="7f5f1-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="7f5f1-141">Para saber mais sobre o suporte do Outlook, confira [comandos de suplemento do Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="7f5f1-142">Depuração</span><span class="sxs-lookup"><span data-stu-id="7f5f1-142">Debugging</span></span>

<span data-ttu-id="7f5f1-143">Para depurar um comando de Suplemento, você deve executá-lo no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="7f5f1-144">Para obter detalhes, confira [Depurar suplementos no Office na Web](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="7f5f1-145">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="7f5f1-145">Best practices</span></span>

<span data-ttu-id="7f5f1-146">Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="7f5f1-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="7f5f1-p107">Use os comandos para representar uma ação específica com um resultado claro e específico para os usuários. Não combine várias ações em um único botão.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p107">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="7f5f1-p108">Forneça ações granulares que tornam a realização de tarefas comuns no seu suplemento mais eficiente. Minimize o número de etapas necessárias para concluir uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p108">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="7f5f1-151">Para o posicionamento dos comandos na faixa de opções do Aplicativo do Office:</span><span class="sxs-lookup"><span data-stu-id="7f5f1-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="7f5f1-p109">Insira os comandos em uma guia existente (Inserir, Revisar e assim por diante) se a funcionalidade fornecida se encaixar ali. Por exemplo, se seu suplemento permitir que os usuários insiram mídia, adicione um grupo à guia Inserir. Observe que nem todas as guias estão disponíveis em todas as versões do Office. Para saber mais, confira o [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p109">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="7f5f1-p110">Insira comandos na guia Página Inicial se a funcionalidade não se encaixar em outra guia e você menos de seis comandos de nível superior. Você também pode adicionar comandos à guia Página Inicial se seu suplemento precisar funcionar em diferentes versões do Office (como o Office para área de trabalho e o Office na Web) e uma guia não está disponível em todas as versões (por exemplo, a guia Design não existe no Office na Web).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p110">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="7f5f1-157">Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="7f5f1-p111">Nomeie seu grupo de acordo com o nome do seu suplemento. Se você tiver vários grupos, nomeie cada grupo com base na funcionalidade que os comandos nesse grupo fornecem.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-p111">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="7f5f1-160">Não adicione botões supérfluos para aumentar o estado real do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="7f5f1-161">Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="7f5f1-162">Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="7f5f1-163">Forneça uma versão do seu suplemento que também funcione em hosts que não tenham suporte para comandos.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="7f5f1-164">Um manifesto de suplemento único pode funcionar tanto em hosts cientes do comando (com os comandos) quanto em hosts não cientes do comando (como um painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="7f5f1-165">*Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*</span><span class="sxs-lookup"><span data-stu-id="7f5f1-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Uma captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="7f5f1-167">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="7f5f1-167">Next steps</span></span>

<span data-ttu-id="7f5f1-168">A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="7f5f1-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="7f5f1-169">Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="7f5f1-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
