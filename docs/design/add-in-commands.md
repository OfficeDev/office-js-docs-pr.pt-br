---
title: Conceitos básicos para comandos de suplemento
description: Aprenda a adicionar botões e itens de menu personalizados da faixa de opções ao Office como parte de um suplemento do Office.
ms.date: 07/27/2021
localization_priority: Priority
ms.openlocfilehash: 285229c643f0e9ab9008905a07767c985b050ad2a46397aedd741a2b0ce9dfa5
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082703"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Comandos de suplemento para Excel, Word e PowerPoint

Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. Os comandos de suplemento ajudam os usuários a localizar e usar o suplemento, o que pode ajudá-lo a aumentar a adoção e a reutilização do suplemento, além de melhorar a retenção de clientes.

Para uma visão geral do recurso, confira o vídeo [Comandos de Suplemento na Faixa de Opções do Aplicativo do Office](https://channel9.msdn.com/events/Build/2016/P551).

> [!NOTE]
> Os catálogos do Microsoft Office SharePoint Online não são compatíveis com os comandos de suplemento. Você pode implantar comandos de suplemento por meio de [Aplicativos integrados](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) ou do [AppSource](/office/dev/store/submit-to-appsource-via-partner-center)ou usar o [sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para implantar o comando de suplemento para teste.

> [!IMPORTANT]
> Os comandos de suplemento também são compatíveis com o Outlook. Para saber mais, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md).

*Figura 1. Suplemento com comandos em execução na Área de Trabalho do Excel*

![Captura de tela mostrando comandos de suplemento realçados na faixa de opções do Excel.](../images/add-in-commands-1.png)

*Figura 2. Suplemento com comandos em execução no Excel na Web*

![Captura de tela de um comando de suplemento no Excel na Web.](../images/add-in-commands-2.png)

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
- ExecuteFunction: carrega uma página HTML invisível e executa uma função JavaScript dentro dela. Para mostrar a interface do usuário dentro de sua função (como erros, progresso ou entrada adicional), você pode usar a API [displayDialog](/javascript/api/office/office.ui).  

### <a name="default-enabled-or-disabled-status"></a>Status padrão Habilitado ou Desabilitado

Você pode especificar se o comando está ativado ou desativado quando o suplemento é iniciado e alterar programaticamente a configuração.

> [!NOTE]
> Esse recurso não tem suporte em todos os aplicativos ou cenários do Office. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](disable-add-in-commands.md).

### <a name="position-on-the-ribbon-preview"></a>Posição na faixa de opções (visualização)

Você pode especificar onde uma guia personalizada é exibida na faixa de opções do aplicativo do Office, como "à direita da guia Página inicial".

> [!NOTE]
> Esse recurso não tem suporte em todos os aplicativos ou cenários do Office. Para saber mais, confira [Posicionar uma guia personalizada na faixa de opções](custom-tab-placement.md).

### <a name="integration-of-built-in-office-buttons-preview"></a>Integração de botões internos do Office (visualização)

Você pode inserir os botões internos da faixa de opções do Office em seus grupos de comandos personalizados e nas guias personalizadas da faixa de opções.

> [!NOTE]
> Esse recurso não tem suporte em todos os aplicativos ou cenários do Office. Para saber mais, confira [Integrar os botões internos do Office em guias personalizadas](built-in-button-integration.md).

### <a name="contextual-tabs-preview"></a>Guias contextuais (pré-visualização)

Você pode especificar que uma guia só seja visível na faixa de opções em determinados contextos, como quando um gráfico é selecionado no Excel.

> [!NOTE]
> Esse recurso não tem suporte em todos os aplicativos ou cenários do Office. Para obter mais informações, confira [Criar guias contextuais personalizadas em Suplementos do Office](contextual-tabs.md).

## <a name="supported-platforms"></a>Plataformas compatíveis

Os comandos de suplemento são atualmente suportados nas plataformas a seguir, exceto para limitações especificadas nas subseções de [Recursos de comandos](#command-capabilities) anteriores.

- Office no Windows (Build 16.0.6769 ou superior, conectado a uma assinatura do Microsoft 365)
- Office 2019 no Windows
- Office no Mac (build 15.33 ou superior, conectado a uma assinatura do Microsoft 365)
- Office 2019 no Mac
- Office na Web

> [!NOTE]
> Para saber mais sobre o suporte do Outlook, confira [comandos de suplemento do Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debug"></a>Depurar

Para depurar um comando de Suplemento, você deve executá-lo no Office na Web. Para obter detalhes, confira [Depurar suplementos no Office na Web](../testing/debug-add-ins-in-office-online.md).

## <a name="best-practices"></a>Práticas recomendadas

Aplique as práticas recomendadas a seguir ao desenvolver comandos de suplemento.

- Use os comandos para representar uma ação específica com um resultado claro e específico para os usuários. Não combine várias ações em um único botão.
- Forneça ações granulares que tornam a realização de tarefas comuns no seu suplemento mais eficiente. Minimize o número de etapas necessárias para concluir uma tarefa.
- Para o posicionamento dos comandos na faixa de opções do Aplicativo do Office:
  - Insira os comandos em uma guia existente (Inserir, Revisar e assim por diante) se a funcionalidade fornecida se encaixar ali. Por exemplo, se seu suplemento permitir que os usuários insiram mídia, adicione um grupo à guia Inserir. Observe que nem todas as guias estão disponíveis em todas as versões do Office. Para saber mais, confira o [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).
  - Insira comandos na guia Página Inicial se a funcionalidade não se encaixar em outra guia e você menos de seis comandos de nível superior. Você também pode adicionar comandos à guia Página Inicial se seu suplemento precisar funcionar em diferentes versões do Office (como o Office para área de trabalho e o Office na Web) e uma guia não está disponível em todas as versões (por exemplo, a guia Design não existe no Office na Web).  
  - Coloque os comandos em uma guia personalizada se você tiver mais de seis comandos de nível superior.
  - Nomeie seu grupo de acordo com o nome do seu suplemento. Se você tiver vários grupos, nomeie cada grupo com base na funcionalidade que os comandos nesse grupo fornecem.
  - Não adicione botões supérfluos para aumentar o estado real do seu suplemento.
  - Não posicione uma guia personalizada à esquerda da guia Página inicial ou dê a ela o foco por padrão quando o documento for aberto, a menos que seu suplemento seja a principal maneira como os usuários vão interagir com o documento. Dar destaque excessivo as inconveniências do seu suplemento e incomodar os usuários e os administradores.
  - Se o seu suplemento for a principal maneira como os usuários interagem com o documento e você tiver uma guia personalizada na faixa de opções, considere integrar na guia os botões das funções do Office que os usuários frequentemente precisarão.
  - Se a funcionalidade fornecida com uma guia personalizada deve estar disponível apenas em determinados contextos, use [guias contextuais personalizadas](contextual-tabs.md). Se você usar guias contextuais personalizadas, certifique-se de implementar uma experiência de [fallback para quando o suplemento for executado em plataformas que não oferecem suporte a guias contextuais personalizadas](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

  > [!NOTE]
  > Os suplementos que ocupam muito espaço podem não passar na [Validação do AppSource](/legal/marketplace/certification-policies).

- Para todos os ícones, siga as [diretrizes de design de ícones](add-in-icons.md).
- Fornece uma versão do seu suplemento que também funciona em aplicativos do Office que não oferecem suporte a comandos. Um único manifesto de suplemento pode funcionar em aplicativos com reconhecimento de comando (com comandos) e sem reconhecimento de comando (como um painel de tarefas).

   *Figura 3. Suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016*

   ![Captura de tela que mostra um suplemento de painel de tarefas no Office 2013 e o mesmo suplemento usando comandos de suplementos no Office 2016. Na versão 2013, o painel de tarefas deve conter todos os comandos, enquanto na versão 2016, os comandos podem estar na faixa de opções.](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a>Próximas etapas

A melhor maneira de começar a usar os comandos de suplemento é conferir os [exemplos de comandos de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) no GitHub.

Saiba mais sobre como especificar comandos de suplemento no manifesto em [Criar comandos de suplemento no manifesto](../develop/create-addin-commands.md) e no conteúdo de referência [VersionOverrides](../reference/manifest/versionoverrides.md).
