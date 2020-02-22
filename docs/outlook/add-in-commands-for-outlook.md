---
title: Comandos de suplementos do Outlook
description: Os comandos de suplementos do Outlook oferecem maneiras de iniciar ações específicas do suplemento na faixa de opções ao adicionar botões ou menus suspensos.
ms.date: 12/05/2019
localization_priority: Priority
ms.openlocfilehash: 4b7249aaaad10f8ddef02540dcd6a3e08524c4db
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165772"
---
# <a name="add-in-commands-for-outlook"></a><span data-ttu-id="991b0-103">Comandos de suplemento para o Outlook</span><span class="sxs-lookup"><span data-stu-id="991b0-103">Add-in commands for Outlook</span></span>

<span data-ttu-id="991b0-p101">Os comandos de suplemento do Outlook oferecem maneiras de iniciar ações específicas do suplemento na faixa de opções adicionando botões ou menus suspensos. Isso permite que os usuários acessem suplementos de maneira simples, intuitiva e discreta. Como eles oferecem maior funcionalidade de forma simplificada, você pode usar comandos de suplemento para criar soluções mais atraentes.</span><span class="sxs-lookup"><span data-stu-id="991b0-p101">Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.</span></span>

> [!NOTE]
> <span data-ttu-id="991b0-107">Comandos de suplemento estão disponíveis somente no Outlook 2013 ou posterior no Windows, Outlook 2016 ou posterior no Mac, Outlook no iPhone, Outlook no Android, Outlook na Web para o Exchange 2016 ou posterior e Outlook na Web para Office 365 e Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="991b0-107">Add-in commands are available only in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on iPhone, Outlook on Android, Outlook on the web for Exchange 2016 or later, and Outlook on the web for Office 365 and Outlook.com.</span></span>
>
> <span data-ttu-id="991b0-108">O suporte para comandos de suplementos no Outlook 2013 requer três atualizações:</span><span class="sxs-lookup"><span data-stu-id="991b0-108">Support for add-in commands in Outlook 2013 requires three updates:</span></span>
> - [<span data-ttu-id="991b0-109">Atualização de segurança de 8 de março de 2016 para o Outlook</span><span class="sxs-lookup"><span data-stu-id="991b0-109">March 8, 2016 security update for Outlook</span></span>](https://support.microsoft.com/kb/3114829)
> - [<span data-ttu-id="991b0-110">Atualização de segurança de 8 de março de 2016 para o Office (KB3114816)</span><span class="sxs-lookup"><span data-stu-id="991b0-110">March 8, 2016 security update for Office (KB3114816)</span></span>](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [<span data-ttu-id="991b0-111">Atualização de segurança de 8 de março de 2016 para o Office (KB3114828)</span><span class="sxs-lookup"><span data-stu-id="991b0-111">March 8, 2016 security update for Office (KB3114828)</span></span>](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> <span data-ttu-id="991b0-112">O suporte para comandos de suplementos no Exchange 2016 requer a [Atualização Cumulativa 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).</span><span class="sxs-lookup"><span data-stu-id="991b0-112">Support for add-in commands in Exchange 2016 requires [Cumulative Update 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).</span></span>

<span data-ttu-id="991b0-p102">Os comandos de suplementos estão disponíveis apenas para suplementos que não usam [regras ItemHasAttachment, ItemHasKnownEntity ou ItemHasRegularExpressionMatch](activation-rules.md) para limitar os tipos de itens em que são ativados. No entanto, os [suplementos contextuais](contextual-outlook-add-ins.md) podem apresentar comandos diferentes, dependendo do item selecionado no momento ser uma mensagem ou um compromisso, e podem optar por serem exibidos em cenários de leitura ou redação. É uma [prática recomendada](../concepts/add-in-development-best-practices.md) usar comandos de suplementos.</span><span class="sxs-lookup"><span data-stu-id="991b0-p102">Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).</span></span>

## <a name="creating-the-add-in-command"></a><span data-ttu-id="991b0-116">Criar o comando de suplemento</span><span class="sxs-lookup"><span data-stu-id="991b0-116">Creating the add-in command</span></span>

<span data-ttu-id="991b0-p103">Os comandos do suplemento são declarados no manifesto do suplemento no elemento [VersionOverrides](../reference/manifest/versionoverrides.md). Esse elemento é uma adição ao esquema de manifesto v1.1 que garante a compatibilidade com versões anteriores. Em um cliente que não dê suporte a `VersionOverrides`, os suplementos existentes continuarão a funcionar como faziam sem comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="991b0-p103">Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.</span></span>

<span data-ttu-id="991b0-120">As entradas de manifesto `VersionOverrides` especificam muitos itens para o suplemento, como host, tipos de controles a serem adicionados à faixa de opções, texto, ícones e quaisquer funções associadas.</span><span class="sxs-lookup"><span data-stu-id="991b0-120">The `VersionOverrides` manifest entries specify many things for the add-in, such as the host, types of controls to add to the ribbon, the text, the icons, and any associated functions.</span></span>

<span data-ttu-id="991b0-p104">Quando um suplemento precisa fornecer atualizações de status, como indicadores de progresso ou mensagens de erro, ele deve fazer isso por meio das [APIs de notificação](/javascript/api/outlook/office.NotificationMessages). O processamento para as notificações também deve ser definido em um arquivo HTML separado que é especificado no nó `FunctionFile` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="991b0-p104">When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.NotificationMessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.</span></span>

<span data-ttu-id="991b0-p105">Os desenvolvedores devem definir ícones para todos os tamanhos necessários, para que os comandos do suplemento se ajustem sem problemas junto com a faixa de opções. Os tamanhos de ícone obrigatórios são 80 x 80 pixels, 32 x 32 pixels e 16 x 16 pixels para área de trabalho e 48 x 48 pixels, 32 x 32 pixels e 25 x 25 pixels para dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="991b0-p105">Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.</span></span>

<span data-ttu-id="991b0-125">Para mais informações sobre a criação de comandos do suplemento, veja [Criar comandos de suplemento em seu manifesto](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="991b0-125">For more information about creating add-in commands, see [Create add-in commands in your manifest](../develop/create-addin-commands.md).</span></span>

## <a name="how-do-add-in-commands-appear"></a><span data-ttu-id="991b0-126">Como os comandos de suplemento são exibidos?</span><span class="sxs-lookup"><span data-stu-id="991b0-126">How do add-in commands appear?</span></span>

<span data-ttu-id="991b0-p106">Um comando de suplemento é mostrado na faixa de opções como um botão. Quando um usuário instala um suplemento, seus comandos são mostrados na interface de usuário como um grupo de botões. Pode ser na guia padrão da faixa de opções ou em uma guia personalizada. Para mensagens, o padrão é a guia **Página Inicial** ou **Mensagem**. Para o calendário, o padrão é a guia **Reunião**, **Ocorrência de Reunião**, **Série de Reuniões** ou **Compromisso**. Para extensões de módulo, o padrão é uma guia personalizada. Na guia padrão, cada suplemento pode ter um grupo da faixa de opções com até seis comandos. Em guias personalizadas, o suplemento pode ter até dez grupos, cada um com seis comandos. Os suplementos estão limitados a apenas uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="991b0-p106">An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.</span></span>

<span data-ttu-id="991b0-132">À medida que a faixa de opções fica mais cheia, os comandos de suplementos serão exibidos no menu estouro.</span><span class="sxs-lookup"><span data-stu-id="991b0-132">As the ribbon gets more crowded, add-in commands will be displayed in the overflow menu.</span></span> <span data-ttu-id="991b0-133">Geralmente, os comandos de um suplemento serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="991b0-133">The add-in commands for an add-in are usually grouped together.</span></span>

![Botões de comando do suplemento na faixa de opções](../images/commands-normal.png)

![Botões de comando do suplemento na faixa de opções e no menu estouro](../images/commands-collapsed.png)

<span data-ttu-id="991b0-p108">Quando um comando do suplemento é adicionado a um suplemento, o nome do suplemento é removido da barra do aplicativo. Permanece apenas o botão de comando de suplemento na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="991b0-p108">When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.</span></span>

### <a name="modern-outlook-on-the-web"></a><span data-ttu-id="991b0-138">Outlook na Web moderno</span><span class="sxs-lookup"><span data-stu-id="991b0-138">Modern Outlook on the web</span></span>

<span data-ttu-id="991b0-139">No Outlook na Web, o nome do suplemento é exibido em um menu estouro.</span><span class="sxs-lookup"><span data-stu-id="991b0-139">In Outlook on the web, the add-in name is displayed in an overflow menu.</span></span> <span data-ttu-id="991b0-140">Se o suplemento tiver vários comandos, você poderá expandir o menu do suplemento para ver o grupo de botões rotulados com o nome do suplemento.</span><span class="sxs-lookup"><span data-stu-id="991b0-140">If the add-in has multiple add-in commands, you can expand the add-in menu to see the group of buttons labeled with the add-in name.</span></span>

![Menu estouro em que os botões de comando do suplemento serão encontrados](../images/commands-overflow-menu-web.png)

![Menu estouro exibindo botões de comando do suplemento](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a><span data-ttu-id="991b0-143">Quais formas da experiência do usuário existem para comandos de suplemento?</span><span class="sxs-lookup"><span data-stu-id="991b0-143">What UX shapes exist for add-in commands?</span></span>

<span data-ttu-id="991b0-p110">A forma da experiência do usuário para um comando de suplemento consiste em uma guia da faixa de opções no aplicativo host que contém botões que podem executar várias funções. Atualmente, há suporte para três formas de interface do usuário:</span><span class="sxs-lookup"><span data-stu-id="991b0-p110">The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:</span></span>

- <span data-ttu-id="991b0-146">Um botão que executa uma função JavaScript</span><span class="sxs-lookup"><span data-stu-id="991b0-146">A button that executes a JavaScript function</span></span>
- <span data-ttu-id="991b0-147">Um botão que inicia um painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="991b0-147">A button that launches a task pane</span></span>
- <span data-ttu-id="991b0-148">Um botão que mostra um menu suspenso com um ou mais botões dos outros dois tipos</span><span class="sxs-lookup"><span data-stu-id="991b0-148">A button that shows a drop-down menu with one or more buttons of the other two types</span></span>

### <a name="executing-a-javascript-function"></a><span data-ttu-id="991b0-149">Executar uma função JavaScript</span><span class="sxs-lookup"><span data-stu-id="991b0-149">Executing a JavaScript function</span></span>

<span data-ttu-id="991b0-p111">Use um botão de comando de suplemento que executa uma função JavaScript para cenários em que o usuário não precisa fazer seleções adicionais para iniciar a ação. Isso pode ser para ações como acompanhar, lembrar-me ou imprimir ou cenários em que o usuário deseja obter informações mais detalhadas de um serviço.</span><span class="sxs-lookup"><span data-stu-id="991b0-p111">Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.</span></span>

<span data-ttu-id="991b0-152">Em extensões de módulo, o botão de comando de suplemento pode executar funções JavaScript que interagem com o conteúdo na interface do usuário principal.</span><span class="sxs-lookup"><span data-stu-id="991b0-152">In module extensions, the add-in command button can execute JavaScript functions that interact with the content in the main user interface.</span></span>

![Um botão que executa uma função na faixa de opções do Outlook.](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a><span data-ttu-id="991b0-154">Iniciar um painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="991b0-154">Launching a task pane</span></span>

<span data-ttu-id="991b0-p112">Use um botão de comando de suplemento para iniciar um painel de tarefas para cenários em que um usuário precisa interagir com um suplemento por um período de tempo mais longo. Por exemplo, o suplemento requer alterações em configurações ou o preenchimento de vários campos.</span><span class="sxs-lookup"><span data-stu-id="991b0-p112">Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.</span></span>

<span data-ttu-id="991b0-p113">A largura padrão do painel de tarefas vertical é de 320 px. O painel de tarefas vertical pode ser redimensionado no Outlook Explorer e no Inspetor. O painel pode ser redimensionado da mesma maneira que o painel de tarefas pendentes e a exibição de lista.</span><span class="sxs-lookup"><span data-stu-id="991b0-p113">The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.</span></span>

![Um botão que abre o painel de tarefas na faixa de opções do Outlook.](../images/commands-task-pane-button-1.png)

<br/>

<span data-ttu-id="991b0-p114">Esta captura de tela mostra um exemplo de um painel de tarefas vertical. O painel é aberto com o nome do comando de suplemento no canto superior esquerdo. Os usuários podem usar o botão **X**, no canto superior direito do painel, para fechar o suplemento ao terminar de usá-lo. Por padrão, esse painel não persistirá entre mensagens. Os suplementos podem ser [compatíveis com a fixação](pinnable-taskpane.md) do painel de tarefas e receber eventos quando uma nova mensagem for selecionada. Todos os elementos de interface do usuário renderizados no painel de tarefas, além do nome do suplemento e do botão fechar, são fornecidos pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="991b0-p114">This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.</span></span>

<span data-ttu-id="991b0-p115">Se um usuário escolher outro comando de suplemento que abre um painel de tarefas, o painel de tarefas será substituído pelo comando usado recentemente. Se um usuário escolher um botão de comando de suplemento que executa uma função ou um menu suspenso enquanto o painel de tarefas estiver aberto, a ação será concluída e o painel de tarefas permanecerá aberto.</span><span class="sxs-lookup"><span data-stu-id="991b0-p115">If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.</span></span>

### <a name="drop-down-menu"></a><span data-ttu-id="991b0-169">Menu suspenso</span><span class="sxs-lookup"><span data-stu-id="991b0-169">Drop-down menu</span></span>

<span data-ttu-id="991b0-p116">Um comando de suplemento de menu suspenso define uma lista estática de botões. Os botões no menu podem ser qualquer combinação de botões que executam uma função ou botões que abrem um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="991b0-p116">A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.</span></span>

![Um botão que exibe o menu na faixa de opções do Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a><span data-ttu-id="991b0-174">Onde os comandos de suplemento aparecem na interface de usuário?</span><span class="sxs-lookup"><span data-stu-id="991b0-174">Where do add-in commands appear in the UI?</span></span>

<span data-ttu-id="991b0-175">Os comandos de suplemento têm suporte em quatro cenários:</span><span class="sxs-lookup"><span data-stu-id="991b0-175">Add-in commands are supported for four scenarios:</span></span>

### <a name="reading-a-message"></a><span data-ttu-id="991b0-176">Ler uma mensagem</span><span class="sxs-lookup"><span data-stu-id="991b0-176">Reading a message</span></span>

<span data-ttu-id="991b0-177">Quando o usuário está lendo uma mensagem no painel de leitura ou na guia **Mensagem** por um formulário de leitura pop-out, os comandos de suplemento adicionados à guia padrão aparecem na guia **Página Inicial**.</span><span class="sxs-lookup"><span data-stu-id="991b0-177">When the user is reading a message in the reading pane or in the **Message** tab for a pop-out read form, add-in commands added to the default tab appear on the **Home** tab.</span></span>

### <a name="composing-a-message"></a><span data-ttu-id="991b0-178">Redigir uma mensagem</span><span class="sxs-lookup"><span data-stu-id="991b0-178">Composing a message</span></span>

<span data-ttu-id="991b0-179">Quando o usuário está compondo uma mensagem, os comandos de suplemento adicionados à guia padrão aparecem na guia **Mensagem**.</span><span class="sxs-lookup"><span data-stu-id="991b0-179">When the user is composing a message, add-in commands added to the default tab appear on the **Message** tab.</span></span>

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a><span data-ttu-id="991b0-180">Criar ou exibir um compromisso ou uma reunião como organizador</span><span class="sxs-lookup"><span data-stu-id="991b0-180">Creating or viewing an appointment or meeting as the organizer</span></span>

<span data-ttu-id="991b0-p117">Quando você cria ou exibe um compromisso ou uma reunião como organizador, os comandos de suplemento adicionados à guia padrão aparecem nas guias **Reunião**, **Ocorrência de Reunião**, **Série de Reuniões** ou **Compromisso** em formulários pop-out. No entanto, se o usuário selecionar um item no calendário, mas não abrir o pop-out, o grupo da faixa de opções do suplemento não ficará visível na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="991b0-p117">When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.</span></span>

### <a name="viewing-a-meeting-as-an-attendee"></a><span data-ttu-id="991b0-183">Exibir uma reunião como participante</span><span class="sxs-lookup"><span data-stu-id="991b0-183">Viewing a meeting as an attendee</span></span>

<span data-ttu-id="991b0-p118">Quando você exibe uma reunião como participante, os comandos de suplemento adicionados à guia padrão aparecem nas guias **Reunião**, **Ocorrência de Reunião** ou **Série de Reuniões** em formulários pop-out. No entanto, se um usuário selecionar um item no calendário, mas não abrir o pop-out, o grupo da faixa de opções do suplemento não ficará visível na faixa de opções</span><span class="sxs-lookup"><span data-stu-id="991b0-p118">When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon</span></span>

### <a name="using-a-module-extension"></a><span data-ttu-id="991b0-186">Usar uma extensão de módulo</span><span class="sxs-lookup"><span data-stu-id="991b0-186">Using a module extension</span></span>

<span data-ttu-id="991b0-187">Quando você usa uma extensão de módulo, os comandos de suplemento aparecem na guia personalizada da extensão.</span><span class="sxs-lookup"><span data-stu-id="991b0-187">When using a module extension, add-in commands appear on the extension's custom tab.</span></span>

## <a name="see-also"></a><span data-ttu-id="991b0-188">Confira também</span><span class="sxs-lookup"><span data-stu-id="991b0-188">See also</span></span>

- [<span data-ttu-id="991b0-189">Definir comandos de suplemento em seu manifesto</span><span class="sxs-lookup"><span data-stu-id="991b0-189">Define add-in commands in your manifest</span></span>](../develop/create-addin-commands.md)
- [<span data-ttu-id="991b0-190">Suplemento do Outlook para demonstração de comando de suplemento</span><span class="sxs-lookup"><span data-stu-id="991b0-190">Add-in command demo Outlook add-in</span></span>](https://github.com/officedev/outlook-add-in-command-demo)