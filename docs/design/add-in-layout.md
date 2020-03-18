---
title: Diretrizes de layout para Suplementos do Office
description: Obter diretrizes sobre como fazer o layout de um painel de tarefas ou de uma caixa de diálogo em um suplemento do Office.
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: 38c98aeed1ddd1af5fcda95aa6d44ff1f1f2e53b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718502"
---
# <a name="layout"></a><span data-ttu-id="27a3b-103">Layout</span><span class="sxs-lookup"><span data-stu-id="27a3b-103">Layout</span></span>
<span data-ttu-id="27a3b-p101">Cada contêiner HTML inserido no Office terá um layout. Esses layouts são das telas principais do suplemento. Nelas, você criará experiências que permitem que os clientes iniciem ações, modifiquem configurações, exibam, rolem ou naveguem pelo conteúdo. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experiência. Se você tiver um site existente com o qual ps clientes estão familiarizados, considere a reutilização de layouts de páginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em contêineres HTML do Office.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p101">Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you will create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.</span></span>

<span data-ttu-id="27a3b-p102">Para obter diretrizes de layout, confira [Painel de tarefas](task-pane-add-ins.md), [Conteúdo](content-add-ins.md) e [Caixa de diálogo](dialog-boxes.md). Para obter mais informações sobre como montar componentes do Office UI Fabric em layouts comuns e fluxos de experiência do usuário, confira [Modelos de padrões de design da experiência do usuário](ux-design-pattern-templates.md).</span><span class="sxs-lookup"><span data-stu-id="27a3b-p102">For guidelines on layout, see [Task pane](task-pane-add-ins.md), [Content](content-add-ins.md), and [Dialog box](dialog-boxes.md). For more information about how to assemble Office UI Fabric components into common layouts and user experience flows, see [UX design patterns templates](ux-design-pattern-templates.md).</span></span>

<span data-ttu-id="27a3b-112">Aplique as seguintes diretrizes gerais aos layouts:</span><span class="sxs-lookup"><span data-stu-id="27a3b-112">Apply the following general guidelines for layouts:</span></span>

*   <span data-ttu-id="27a3b-p103">Evite margens estreitas ou amplas em contêineres HTML. 20 pixels é um ótimo padrão.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p103">Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default.</span></span>
*   <span data-ttu-id="27a3b-p104">Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p104">Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.</span></span>
*   <span data-ttu-id="27a3b-p105">As interfaces do Office estão em uma grade de 4px. Procure manter o preenchimento entre os elementos como múltiplos de 4.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p105">Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4.</span></span>
*   <span data-ttu-id="27a3b-119">Sobrecarregar a interface pode causar confusão e prejudicar a facilidade de uso com interações de toque.</span><span class="sxs-lookup"><span data-stu-id="27a3b-119">Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions.</span></span>
*   <span data-ttu-id="27a3b-p106">Mantenha layouts consistentes entre as telas. Alterações de layout inesperadas parecem bugs visuais que contribuem para a falta de confiança na solução.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p106">Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution.</span></span>
*   <span data-ttu-id="27a3b-p107">Siga os padrões de layout comuns. As convenções ajudam os usuários a compreender como usar uma interface.</span><span class="sxs-lookup"><span data-stu-id="27a3b-p107">Follow common layout patterns. Conventions help users understand how to use an interface.</span></span>
*   <span data-ttu-id="27a3b-124">Evite elementos redundantes como identidade visual ou comandos.</span><span class="sxs-lookup"><span data-stu-id="27a3b-124">Avoid redundant elements like branding or commands.</span></span>
*   <span data-ttu-id="27a3b-125">Consolide os controles e modos de exibição para evitar exigir muitos movimentos do mouse.</span><span class="sxs-lookup"><span data-stu-id="27a3b-125">Consolidate controls and views to avoid requiring too much mouse movement.</span></span>
*   <span data-ttu-id="27a3b-126">Crie experiências ágeis que se adaptem a alturas e larguras de contêineres HTML.</span><span class="sxs-lookup"><span data-stu-id="27a3b-126">Create responsive experiences that adapt to HTML container widths and heights.</span></span>
