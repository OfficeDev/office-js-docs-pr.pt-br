---
title: Padrões de navegação para Suplementos do Office
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: b7fee6fad703ce7c8f4c5f8b848d6bf28b239b09
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432967"
---
# <a name="navigation-patterns"></a><span data-ttu-id="04f40-102">Padrões de navegação</span><span class="sxs-lookup"><span data-stu-id="04f40-102">Navigation patterns</span></span>

<span data-ttu-id="04f40-103">Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada.</span><span class="sxs-lookup"><span data-stu-id="04f40-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="04f40-104">É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.</span><span class="sxs-lookup"><span data-stu-id="04f40-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="04f40-105">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="04f40-105">Best practices</span></span>

| <span data-ttu-id="04f40-106">Fazer</span><span class="sxs-lookup"><span data-stu-id="04f40-106">Do</span></span>    | <span data-ttu-id="04f40-107">Não fazer</span><span class="sxs-lookup"><span data-stu-id="04f40-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="04f40-108">Certifique-se de que o usuário tenha uma opção de navegação claramente visível.</span><span class="sxs-lookup"><span data-stu-id="04f40-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="04f40-109">Não complique o processo de navegação usando a interface de usuário não padrão.</span><span class="sxs-lookup"><span data-stu-id="04f40-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="04f40-110">Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="04f40-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="04f40-111">Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="04f40-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="04f40-112">Barra de comandos</span><span class="sxs-lookup"><span data-stu-id="04f40-112">Command Bar</span></span>

<span data-ttu-id="04f40-113">A Barra de comandos é uma superfície que abriga comandos que operam no conteúdo da janela, painel ou região pai sobre o qual ela reside.</span><span class="sxs-lookup"><span data-stu-id="04f40-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="04f40-114">Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.</span><span class="sxs-lookup"><span data-stu-id="04f40-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Comandos: especificações para o painel de tarefas da área de trabalho](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="04f40-116">Barra de guias</span><span class="sxs-lookup"><span data-stu-id="04f40-116">Tab Bar</span></span>

<span data-ttu-id="04f40-117">Mostra a navegação usando botões com texto empilhado na vertical e ícones.</span><span class="sxs-lookup"><span data-stu-id="04f40-117">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="04f40-118">Use a barra de guias para proporcionar uma navegação em guias com títulos curtos e descritivos.</span><span class="sxs-lookup"><span data-stu-id="04f40-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Barra de guias: especificações para o painel de tarefas da área de trabalho](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="04f40-120">Botão Voltar</span><span class="sxs-lookup"><span data-stu-id="04f40-120">Back button</span></span>

<span data-ttu-id="04f40-121">O botão Voltar permite que os usuários se recuperem de uma ação de navegação detalhada.</span><span class="sxs-lookup"><span data-stu-id="04f40-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="04f40-122">Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.</span><span class="sxs-lookup"><span data-stu-id="04f40-122">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Botão Voltar: especificações para o painel de tarefas da área de trabalho](../images/add-in-back-button.png)
