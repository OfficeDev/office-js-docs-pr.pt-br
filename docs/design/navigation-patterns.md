---
title: Padrões de navegação para Suplementos do Office
description: Saiba mais sobre as práticas recomendadas para usar barras de comandos, barras de guias e botões voltar para projetar a navegação de um suplemento do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132029"
---
# <a name="navigation-patterns"></a><span data-ttu-id="3beb5-103">Padrões de navegação</span><span class="sxs-lookup"><span data-stu-id="3beb5-103">Navigation patterns</span></span>

<span data-ttu-id="3beb5-104">Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada.</span><span class="sxs-lookup"><span data-stu-id="3beb5-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="3beb5-105">É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3beb5-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="3beb5-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="3beb5-106">Best practices</span></span>

| <span data-ttu-id="3beb5-107">Fazer</span><span class="sxs-lookup"><span data-stu-id="3beb5-107">Do</span></span>    | <span data-ttu-id="3beb5-108">Não fazer</span><span class="sxs-lookup"><span data-stu-id="3beb5-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="3beb5-109">Certifique-se de que o usuário tenha uma opção de navegação claramente visível.</span><span class="sxs-lookup"><span data-stu-id="3beb5-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="3beb5-110">Não complique o processo de navegação usando a interface de usuário não padrão.</span><span class="sxs-lookup"><span data-stu-id="3beb5-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="3beb5-111">Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="3beb5-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="3beb5-112">Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3beb5-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>

## <a name="command-bar"></a><span data-ttu-id="3beb5-113">Barra de comandos</span><span class="sxs-lookup"><span data-stu-id="3beb5-113">Command Bar</span></span>

<span data-ttu-id="3beb5-114">O CommandBar é uma superfície dentro do painel de tarefas que abriga comandos que operam no conteúdo da janela, painel ou região pai que residem acima.</span><span class="sxs-lookup"><span data-stu-id="3beb5-114">The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="3beb5-115">Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.</span><span class="sxs-lookup"><span data-stu-id="3beb5-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Ilustração mostrando uma barra de comandos dentro de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a><span data-ttu-id="3beb5-118">Barra de guias</span><span class="sxs-lookup"><span data-stu-id="3beb5-118">Tab Bar</span></span>

<span data-ttu-id="3beb5-119">A barra de guias mostra a navegação usando botões com texto empilhado verticalmente e ícones.</span><span class="sxs-lookup"><span data-stu-id="3beb5-119">The tab bar shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="3beb5-120">Use a barra de guias para fornecer a navegação usando guias com títulos curtos e descritivos.</span><span class="sxs-lookup"><span data-stu-id="3beb5-120">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Ilustração mostrando uma barra de guias dentro de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-tab-bar.png)

## <a name="back-button"></a><span data-ttu-id="3beb5-123">Botão Voltar</span><span class="sxs-lookup"><span data-stu-id="3beb5-123">Back Button</span></span>

<span data-ttu-id="3beb5-124">O botão voltar permite que os usuários se recuperem de uma ação de navegação de busca detalhada.</span><span class="sxs-lookup"><span data-stu-id="3beb5-124">The back button allows users to recover from a drill-down navigational action.</span></span> <span data-ttu-id="3beb5-125">Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.</span><span class="sxs-lookup"><span data-stu-id="3beb5-125">This pattern helps ensure users follow an ordered series of steps.</span></span>

![Ilustração mostrando um botão voltar dentro de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-back-button.png)
