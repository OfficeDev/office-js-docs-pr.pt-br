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
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Comandos de suplemento para Excel, Word e PowerPoint

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Aplicativo do Office](https://channel9.msdn.com/events/Build/2016/P551).

> [!NOTE]
> SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.

> [!IMPORTANT]
> Os comandos de suplemento também são compatíveis com o Outlook. Para saber mais, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md).

*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*

![Captura de tela de um comando de suplemento no Excel](../images/add-in-commands-1.png)

*Figura 2. Suplemento com comandos em execução no Excel na Web*

![Captura de tela de um comando de suplemento no Excel na Web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a>Recursos de comandos

Os seguintes recursos de comando são compatíveis no momento.

> [!NOTE]
> Atualmente os suplementos de conteúdo não dão suporte a comandos de suplemento.

### <a name="extension-points"></a>Pontos de extensão

- Guias da faixa de opções: estender as guias internas ou criar uma nova guia personalizada.
- Menus de contexto: estender menus de contexto selecionados.

### <a name="control-types"></a>Tipos de controle

- Botões simples: disparar ações específicas.
- Menus – menu suspenso simples com botões que disparam ações.

### <a name="actions"></a>Ações

- ShowTaskpane: exibe um ou vários painéis que carregam páginas HTML personalizadas dentro deles.
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.  

### <a name="default-enabled-or-disabled-status-preview"></a>Status padrão Habilitado ou Desabilitado (visualização)

Você pode especificar se o comando está ativado ou desativado quando o suplemento é iniciado e alterar programaticamente a configuração.

> [!NOTE]
> Esse recurso está em visualização e não tem suporte em todos os hosts ou cenários. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](disable-add-in-commands.md).

## <a name="supported-platforms"></a>Plataformas compatíveis

Os comandos de suplemento atualmente têm suporte nas seguintes plataformas.

- Office no Windows (Build 16.0.6769+, conectado à assinatura do Microsoft 365)
- Office 2019 no Windows
- Office no Windows (Build 15.33+, conectado à assinatura do Microsoft 365)
- Office 2019 no Mac
- Office na Web

> [!NOTE]
> Para saber mais sobre o suporte do Outlook, confira [comandos de suplemento do Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debugging"></a>Depuração

Para depurar um comando de Suplemento, você deve executá-lo no Office na Web. Para obter detalhes, confira [Depurar suplementos no Office na Web](../testing/debug-add-ins-in-office-online.md).

## <a name="best-practices"></a>Práticas recomendadas

Aplique as seguintes práticas recomendadas ao desenvolver comandos de suplementos:

- Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.
- Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.
- Para o posicionamento dos comandos na faixa de opções do Aplicativo do Office:
    - Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).
    - Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).  
    - Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.
    - Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.
    - Não adicione botões supérfluos para aumentar o estado real do seu suplemento.

     > [!NOTE]
     > Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/legal/marketplace/certification-policies).

- Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).
- Forneça uma versão do seu suplemento que também funcione em hosts que não tenham suporte para comandos. Um manifesto de suplemento único pode funcionar tanto em hosts cientes do comando (com os comandos) quanto em hosts não cientes do comando (como um painel de tarefas).

   *Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*

   ![Uma captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a>Próximas etapas

A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.

Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](../reference/manifest/versionoverrides.md).
