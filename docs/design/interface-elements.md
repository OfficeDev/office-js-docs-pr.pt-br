---
title: Elementos da interface do usuário do Office para suplementos do Office
description: Obter uma visão geral dos diferentes tipos de elementos da interface do usuário em um Office Add-in.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: b09530aa25e7383e189520e7f1030a5f35b94ab4348fe4bad40773092cd5e08b
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081773"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Elementos da interface do usuário do Office para suplementos do Office

Você pode usar vários tipos de elementos para estender a interface do usuário do Office, incluindo comandos de suplemento e contêineres HTML. Esses elementos de interface do usuário parecem uma extensão natural do Office e funcionam entre plataformas. Você pode inserir um código personalizado baseado na Web em qualquer um desses elementos.

A imagem a seguir mostra os tipos de elementos de interface do usuário do Office que você pode criar.

![Diagrama mostrando comandos de add-in na faixa de opções, um painel de tarefas e uma caixa de diálogo/um Office de conteúdo.](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>Comandos de suplemento

Use [comandos de add-in](add-in-commands.md) para adicionar pontos de entrada ao seu add-in à faixa Aplicativo do Office faixa de opções. Comandos iniciam ações no suplemento executando código JavaScript ou iniciando um contêiner HTML. Você pode criar dois tipos de comandos de suplemento.

|Tipo de comando|Descrição|
|:---------------|:--------------|
|Botões, menus e guias da faixa de opções|Use para adicionar botões personalizados, menus (menus suspensos) ou guias à faixa de opções padrão no Office. Use botões e menus para disparar uma ação no Office. Use guias para agrupar e organizar botões e menus.|
|Menus de contexto| Use para estender o menu de contexto padrão. Menus de contexto são exibidos quando os usuários clicam com o botão direito do mouse no texto em um documento do Office ou uma tabela no Excel.|

## <a name="html-containers"></a>Contêineres HTML

Use contêineres HTML para inserir código de interface do usuário baseado em HTML em clientes Office. Essas páginas da Web podem fazer referência à API do JavaScript do Office para interagir com conteúdo no documento. Você pode criar três tipos de contêineres HTML.

|Contêiner HTML|Descrição|
|:-----------------|:--------------|
|[Painéis de tarefas](task-pane-add-ins.md)|Exibir a interface do usuário personalizada no painel à direita do documento do Office. Use os painéis de tarefas para permitir que os usuários interajam com o suplemento lado a lado com o documento do Office.|
|[Suplementos de conteúdo](content-add-ins.md)|Exibir a interface do usuário personalizada inserida em documentos do Office. Use os suplementos de conteúdo para permitir que os usuários interajam com o suplemento diretamente no documento do Office. Por exemplo, talvez você queira mostrar conteúdo externo, como vídeos ou visualizações de dados de outras fontes. |
|[Caixas de diálogo](dialog-boxes.md)|Exibir uma interface do usuário personalizada em uma caixa de diálogo que se sobrepõe ao documento do Office. Use uma caixa de diálogo para interações que requerem foco e mais espaço, e não exigem uma interação lado a lado com o documento.|

## <a name="see-also"></a>Confira também

- [Comandos de suplemento para Excel, Word e PowerPoint](add-in-commands.md)
- [Painéis de tarefas](task-pane-add-ins.md)
- [Suplementos de conteúdo](content-add-ins.md)
- [Caixas de diálogo](dialog-boxes.md)
