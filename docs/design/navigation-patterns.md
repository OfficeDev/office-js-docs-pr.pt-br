---
title: Padrões de navegação para Suplementos do Office
description: Saiba mais sobre as práticas recomendadas para usar barras de comandos, barras de guias e botões voltar para projetar a navegação de um suplemento do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 6fb025a897cfc820117a0b6153acc92c2aeb837e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718752"
---
# <a name="navigation-patterns"></a><span data-ttu-id="64e49-103">Padrões de navegação</span><span class="sxs-lookup"><span data-stu-id="64e49-103">Navigation patterns</span></span>

<span data-ttu-id="64e49-104">Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada.</span><span class="sxs-lookup"><span data-stu-id="64e49-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="64e49-105">É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.</span><span class="sxs-lookup"><span data-stu-id="64e49-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="64e49-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="64e49-106">Best practices</span></span>

| <span data-ttu-id="64e49-107">Fazer</span><span class="sxs-lookup"><span data-stu-id="64e49-107">Do</span></span>    | <span data-ttu-id="64e49-108">Não fazer</span><span class="sxs-lookup"><span data-stu-id="64e49-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="64e49-109">Certifique-se de que o usuário tenha uma opção de navegação claramente visível.</span><span class="sxs-lookup"><span data-stu-id="64e49-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="64e49-110">Não complique o processo de navegação usando a interface de usuário não padrão.</span><span class="sxs-lookup"><span data-stu-id="64e49-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="64e49-111">Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="64e49-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="64e49-112">Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="64e49-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="64e49-113">Barra de comandos</span><span class="sxs-lookup"><span data-stu-id="64e49-113">Command Bar</span></span>

<span data-ttu-id="64e49-114">A Barra de comandos é uma superfície que abriga comandos que operam no conteúdo da janela, painel ou região pai sobre o qual ela reside.</span><span class="sxs-lookup"><span data-stu-id="64e49-114">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="64e49-115">Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.</span><span class="sxs-lookup"><span data-stu-id="64e49-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Comandos: especificações para o painel de tarefas da área de trabalho](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="64e49-117">Barra de guias</span><span class="sxs-lookup"><span data-stu-id="64e49-117">Tab Bar</span></span>

<span data-ttu-id="64e49-118">Mostra a navegação usando botões com texto empilhado na vertical e ícones.</span><span class="sxs-lookup"><span data-stu-id="64e49-118">Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="64e49-119">Use a barra de guias para proporcionar uma navegação em guias com títulos curtos e descritivos.</span><span class="sxs-lookup"><span data-stu-id="64e49-119">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Barra de guias: especificações para o painel de tarefas da área de trabalho](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="64e49-121">Botão Voltar</span><span class="sxs-lookup"><span data-stu-id="64e49-121">Back Button</span></span>

<span data-ttu-id="64e49-122">O botão Voltar permite que os usuários se recuperem de uma ação de navegação detalhada.</span><span class="sxs-lookup"><span data-stu-id="64e49-122">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="64e49-123">Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.</span><span class="sxs-lookup"><span data-stu-id="64e49-123">This pattern helps ensure users follow an ordered series of steps.</span></span>  

![Botão Voltar: especificações para o painel de tarefas da área de trabalho](../images/add-in-back-button.png)
