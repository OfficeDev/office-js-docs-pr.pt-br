---
title: Conceitos básicos para comandos de suplemento
description: Aprenda a adicionar botões e itens de menu personalizados da faixa de opções ao Office como parte de um suplemento do Office.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 922148130aa5384050813155819d2db4c75fe284
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076136"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="b9def-103">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b9def-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="b9def-p101">Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. Os comandos de suplemento ajudam os usuários a localizar e usar o suplemento, o que pode ajudá-lo a aumentar a adoção e a reutilização do suplemento, além de melhorar a retenção de clientes.</span><span class="sxs-lookup"><span data-stu-id="b9def-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="b9def-108">Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Aplicativo do Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="b9def-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-p102">Os catálogos do SharePoint não são compatíveis com os comandos de suplemento. É possível implantar comandos de suplemento pela [Implantação centralizada](../publish/centralized-deployment.md) ou pelo [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ou usar [sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para implantar seu comando de suplemento para testes.</span><span class="sxs-lookup"><span data-stu-id="b9def-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9def-111">Os comandos de suplemento também são compatíveis com o Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9def-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="b9def-112">Para saber mais, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="b9def-113">*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*</span><span class="sxs-lookup"><span data-stu-id="b9def-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Captura de tela mostrando comandos de suplemento realçados na faixa de opções do Excel.](../images/add-in-commands-1.png)

<span data-ttu-id="b9def-115">*Figura 2. Suplemento com comandos em execução no Excel na Web*</span><span class="sxs-lookup"><span data-stu-id="b9def-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Captura de tela de um comando de suplemento no Excel na Web.](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="b9def-117">Recursos de comandos</span><span class="sxs-lookup"><span data-stu-id="b9def-117">Command capabilities</span></span>

<span data-ttu-id="b9def-118">Os seguintes recursos de comando são compatíveis no momento.</span><span class="sxs-lookup"><span data-stu-id="b9def-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-119">Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="b9def-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="b9def-120">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b9def-120">Extension points</span></span>

- <span data-ttu-id="b9def-121">Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="b9def-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="b9def-122">Menus de contexto: estender menus de contexto selecionados.</span><span class="sxs-lookup"><span data-stu-id="b9def-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="b9def-123">Tipos de controle</span><span class="sxs-lookup"><span data-stu-id="b9def-123">Control types</span></span>

- <span data-ttu-id="b9def-124">Botões simples: disparar ações específicas.</span><span class="sxs-lookup"><span data-stu-id="b9def-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="b9def-125">Menus – menu suspenso simples com botões que disparam ações.</span><span class="sxs-lookup"><span data-stu-id="b9def-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="b9def-126">Ações</span><span class="sxs-lookup"><span data-stu-id="b9def-126">Actions</span></span>

- <span data-ttu-id="b9def-127">ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.</span><span class="sxs-lookup"><span data-stu-id="b9def-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="b9def-p104">ExecuteFunction: carrega uma página HTML invisível e executa uma função JavaScript dentro dela. Para mostrar a interface do usuário dentro de sua função (como erros, progresso ou entrada adicional), você pode usar a API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="b9def-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status"></a><span data-ttu-id="b9def-130">Status padrão Habilitado ou Desabilitado</span><span class="sxs-lookup"><span data-stu-id="b9def-130">Default Enabled or Disabled Status</span></span>

<span data-ttu-id="b9def-131">Você pode especificar se o comando está ativado ou desativado quando o suplemento é iniciado e alterar programaticamente a configuração.</span><span class="sxs-lookup"><span data-stu-id="b9def-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-132">Esse recurso não tem suporte em todos os aplicativos ou cenários do Office.</span><span class="sxs-lookup"><span data-stu-id="b9def-132">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b9def-133">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

### <a name="position-on-the-ribbon-preview"></a><span data-ttu-id="b9def-134">Posição na faixa de opções (visualização)</span><span class="sxs-lookup"><span data-stu-id="b9def-134">Position on the ribbon (preview)</span></span>

<span data-ttu-id="b9def-135">Você pode especificar onde uma guia personalizada é exibida na faixa de opções do aplicativo do Office, como "à direita da guia Página inicial".</span><span class="sxs-lookup"><span data-stu-id="b9def-135">You can specify where a custom tab appears on the Office application's ribbon, such as "just to the right of the Home tab".</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-136">Esse recurso não tem suporte em todos os aplicativos ou cenários do Office.</span><span class="sxs-lookup"><span data-stu-id="b9def-136">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b9def-137">Para saber mais, confira [Posicionar uma guia personalizada na faixa de opções](custom-tab-placement.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-137">For more information, see [Position a custom tab on the ribbon](custom-tab-placement.md).</span></span>

### <a name="integration-of-built-in-office-buttons-preview"></a><span data-ttu-id="b9def-138">Integração de botões internos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="b9def-138">Integration of built-in Office buttons (preview)</span></span>

<span data-ttu-id="b9def-139">Você pode inserir os botões internos da faixa de opções do Office em seus grupos de comandos personalizados e nas guias personalizadas da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b9def-139">You can insert the built-in Office ribbon buttons into your custom command groups and custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-140">Esse recurso não tem suporte em todos os aplicativos ou cenários do Office.</span><span class="sxs-lookup"><span data-stu-id="b9def-140">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b9def-141">Para saber mais, confira [Integrar os botões internos do Office em guias personalizadas](built-in-button-integration.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-141">For more information, see [Integrate built-in Office buttons into custom tabs](built-in-button-integration.md).</span></span>

### <a name="contextual-tabs-preview"></a><span data-ttu-id="b9def-142">Guias contextuais (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="b9def-142">Contextual tabs (preview)</span></span>

<span data-ttu-id="b9def-143">Você pode especificar que uma guia só seja visível na faixa de opções em determinados contextos, como quando um gráfico é selecionado no Excel.</span><span class="sxs-lookup"><span data-stu-id="b9def-143">You can specify that a tab is only visible on the ribbon in certain contexts, such as when a chart is selected in Excel.</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-144">Esse recurso não tem suporte em todos os aplicativos ou cenários do Office.</span><span class="sxs-lookup"><span data-stu-id="b9def-144">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b9def-145">Para obter mais informações, confira [Criar guias contextuais personalizadas em Suplementos do Office](contextual-tabs.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-145">For more information, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="b9def-146">Plataformas compatíveis</span><span class="sxs-lookup"><span data-stu-id="b9def-146">Supported platforms</span></span>

<span data-ttu-id="b9def-147">Os comandos de suplemento são atualmente suportados nas plataformas a seguir, exceto para limitações especificadas nas subseções de [Recursos de comandos](#command-capabilities) anteriores.</span><span class="sxs-lookup"><span data-stu-id="b9def-147">Add-in commands are currently supported on the following platforms, except for limitations specified in the subsections of [Command capabilities](#command-capabilities) earlier.</span></span>

- <span data-ttu-id="b9def-148">Office no Windows (Build 16.0.6769 ou superior, conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b9def-148">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b9def-149">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b9def-149">Office 2019 on Windows</span></span>
- <span data-ttu-id="b9def-150">Office no Mac (build 15.33 ou superior, conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b9def-150">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b9def-151">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b9def-151">Office 2019 on Mac</span></span>
- <span data-ttu-id="b9def-152">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b9def-152">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="b9def-153">Para saber mais sobre o suporte do Outlook, confira [comandos de suplemento do Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-153">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="b9def-154">Depuração</span><span class="sxs-lookup"><span data-stu-id="b9def-154">Debugging</span></span>

<span data-ttu-id="b9def-155">Para depurar um comando de Suplemento, você deve executá-lo no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="b9def-155">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="b9def-156">Para obter detalhes, confira [Depurar suplementos no Office na Web](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-156">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="b9def-157">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="b9def-157">Best practices</span></span>

<span data-ttu-id="b9def-158">Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="b9def-158">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="b9def-p110">Use os comandos para representar uma ação específica com um resultado claro e específico para os usuários. Não combine várias ações em um único botão.</span><span class="sxs-lookup"><span data-stu-id="b9def-p110">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="b9def-p111">Forneça ações granulares que tornam a realização de tarefas comuns no seu suplemento mais eficiente. Minimize o número de etapas necessárias para concluir uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="b9def-p111">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="b9def-163">Para o posicionamento dos comandos na faixa de opções do Aplicativo do Office:</span><span class="sxs-lookup"><span data-stu-id="b9def-163">For the placement of your commands in the Office app ribbon:</span></span>
  - <span data-ttu-id="b9def-p112">Insira os comandos em uma guia existente (Inserir, Revisar e assim por diante) se a funcionalidade fornecida se encaixar ali. Por exemplo, se seu suplemento permitir que os usuários insiram mídia, adicione um grupo à guia Inserir. Observe que nem todas as guias estão disponíveis em todas as versões do Office. Para saber mais, confira o [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-p112">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
  - <span data-ttu-id="b9def-p113">Insira comandos na guia Página Inicial se a funcionalidade não se encaixar em outra guia e você menos de seis comandos de nível superior. Você também pode adicionar comandos à guia Página Inicial se seu suplemento precisar funcionar em diferentes versões do Office (como o Office para área de trabalho e o Office na Web) e uma guia não está disponível em todas as versões (por exemplo, a guia Design não existe no Office na Web).</span><span class="sxs-lookup"><span data-stu-id="b9def-p113">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
  - <span data-ttu-id="b9def-169">Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.</span><span class="sxs-lookup"><span data-stu-id="b9def-169">Place commands on a custom tab if you have more than six top-level commands.</span></span>
  - <span data-ttu-id="b9def-p114">Nomeie seu grupo de acordo com o nome do seu suplemento. Se você tiver vários grupos, nomeie cada grupo com base na funcionalidade que os comandos nesse grupo fornecem.</span><span class="sxs-lookup"><span data-stu-id="b9def-p114">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
  - <span data-ttu-id="b9def-172">Não adicione botões supérfluos para aumentar o estado real do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b9def-172">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>
  - <span data-ttu-id="b9def-173">Não posicione uma guia personalizada à esquerda da guia Página inicial ou dê a ela o foco por padrão quando o documento for aberto, a menos que seu suplemento seja a principal maneira como os usuários vão interagir com o documento.</span><span class="sxs-lookup"><span data-stu-id="b9def-173">Do not position a custom tab to the left of the Home tab, or give it focus by default when the document opens, unless your add-in is the primary way users will interact with the document.</span></span> <span data-ttu-id="b9def-174">Dar destaque excessivo as inconveniências do seu suplemento e incomodar os usuários e os administradores.</span><span class="sxs-lookup"><span data-stu-id="b9def-174">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span>
  - <span data-ttu-id="b9def-175">Se o seu suplemento for a principal maneira como os usuários interagem com o documento e você tiver uma guia personalizada na faixa de opções, considere integrar na guia os botões das funções do Office que os usuários frequentemente precisarão.</span><span class="sxs-lookup"><span data-stu-id="b9def-175">If your add-in is the primary way users interact with the document and you have a custom ribbon tab, consider integrating into the tab the buttons for the Office functions that users will frequently need.</span></span>
  - <span data-ttu-id="b9def-176">Se a funcionalidade fornecida com uma guia personalizada deve estar disponível apenas em determinados contextos, use [guias contextuais personalizadas](contextual-tabs.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-176">If the functionality that is provided with a custom tab should only be available in certain contexts, use [custom contextual tabs](contextual-tabs.md).</span></span> <span data-ttu-id="b9def-177">Se você usar guias contextuais personalizadas, certifique-se de implementar uma experiência de [fallback para quando o suplemento for executado em plataformas que não oferecem suporte a guias contextuais personalizadas](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span><span class="sxs-lookup"><span data-stu-id="b9def-177">If you use custom contextual tabs, make sure to implement a [fallback experience for when your add-in runs on platforms that don't support custom contextual tabs](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

  > [!NOTE]
  > <span data-ttu-id="b9def-178">Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="b9def-178">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="b9def-179">Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-179">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="b9def-p117">Fornece uma versão do seu suplemento que também funciona em aplicativos do Office que não oferecem suporte a comandos. Um único manifesto de suplemento pode funcionar em aplicativos com reconhecimento de comando (com comandos) e sem reconhecimento de comando (como um painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="b9def-p117">Provide a version of your add-in that also works on Office applications that do not support commands. A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) applications.</span></span>

   <span data-ttu-id="b9def-182">*Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*</span><span class="sxs-lookup"><span data-stu-id="b9def-182">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016.](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a><span data-ttu-id="b9def-185">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b9def-185">Next steps</span></span>

<span data-ttu-id="b9def-186">A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="b9def-186">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="b9def-187">Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="b9def-187">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
