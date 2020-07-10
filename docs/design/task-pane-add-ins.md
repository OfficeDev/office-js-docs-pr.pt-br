---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093753"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="1bc69-103">Painéis de tarefas nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1bc69-103">Task panes in Office Add-ins</span></span>
 
<span data-ttu-id="1bc69-104">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook.</span><span class="sxs-lookup"><span data-stu-id="1bc69-104">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook.</span></span> <span data-ttu-id="1bc69-105">Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source.</span><span class="sxs-lookup"><span data-stu-id="1bc69-105">Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source.</span></span> <span data-ttu-id="1bc69-106">Use task panes when you don't need to embed functionality directly into the document.</span><span class="sxs-lookup"><span data-stu-id="1bc69-106">Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="1bc69-107">*Figura 1. Layout típico do painel de tarefa*</span><span class="sxs-lookup"><span data-stu-id="1bc69-107">*Figure 1. Typical task pane layout*</span></span>

![Imagem exibindo um layout típico do painel de tarefas](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="1bc69-109">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="1bc69-109">Best practices</span></span>

|<span data-ttu-id="1bc69-110">**Faça**</span><span class="sxs-lookup"><span data-stu-id="1bc69-110">**Do**</span></span>|<span data-ttu-id="1bc69-111">**Não faça**</span><span class="sxs-lookup"><span data-stu-id="1bc69-111">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="1bc69-112">Inclua o nome do seu suplemento no título.</span><span class="sxs-lookup"><span data-stu-id="1bc69-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1bc69-113">Não adicione o nome da sua empresa ao título.</span><span class="sxs-lookup"><span data-stu-id="1bc69-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1bc69-114">Use nomes descritivos curtos no título.</span><span class="sxs-lookup"><span data-stu-id="1bc69-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1bc69-115">Não acrescente cadeias de caracteres, como "suplemento", "para Word" ou "para Office", ao título do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1bc69-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1bc69-116">Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</span><span class="sxs-lookup"><span data-stu-id="1bc69-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="1bc69-117">Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.</span><span class="sxs-lookup"><span data-stu-id="1bc69-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||


## <a name="variants"></a><span data-ttu-id="1bc69-118">Variantes</span><span class="sxs-lookup"><span data-stu-id="1bc69-118">Variants</span></span>

<span data-ttu-id="1bc69-119">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution.</span><span class="sxs-lookup"><span data-stu-id="1bc69-119">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution.</span></span> <span data-ttu-id="1bc69-120">For Excel, additional vertical space is required to accommodate the formula bar.</span><span class="sxs-lookup"><span data-stu-id="1bc69-120">For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="1bc69-121">*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*</span><span class="sxs-lookup"><span data-stu-id="1bc69-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![Imagem exibindo os tamanhos de painel de tarefas da área de trabalho em 1366 x 768](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="1bc69-123">Excel – 320 x 455</span><span class="sxs-lookup"><span data-stu-id="1bc69-123">Excel - 320x455</span></span>
- <span data-ttu-id="1bc69-124">PowerPoint – 320 x 531</span><span class="sxs-lookup"><span data-stu-id="1bc69-124">PowerPoint - 320x531</span></span>
- <span data-ttu-id="1bc69-125">Word – 320 x 531</span><span class="sxs-lookup"><span data-stu-id="1bc69-125">Word - 320x531</span></span>
- <span data-ttu-id="1bc69-126">Outlook – 348 x 535</span><span class="sxs-lookup"><span data-stu-id="1bc69-126">Outlook - 348x535</span></span>

<br/>

<span data-ttu-id="1bc69-127">*Figura 3. Tamanhos de painel de tarefas do Office*</span><span class="sxs-lookup"><span data-stu-id="1bc69-127">*Figure 3. Office task pane sizes*</span></span>

![Imagem exibindo os tamanhos de painel de tarefas da área de trabalho em 1366 x 768](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="1bc69-129">Excel – 350 x 378</span><span class="sxs-lookup"><span data-stu-id="1bc69-129">Excel - 350x378</span></span>
- <span data-ttu-id="1bc69-130">PowerPoint – 348 x 391</span><span class="sxs-lookup"><span data-stu-id="1bc69-130">PowerPoint - 348x391</span></span>
- <span data-ttu-id="1bc69-131">Word – 329 x 445</span><span class="sxs-lookup"><span data-stu-id="1bc69-131">Word - 329x445</span></span>
- <span data-ttu-id="1bc69-132">Outlook (na Web) - 320x570</span><span class="sxs-lookup"><span data-stu-id="1bc69-132">Outlook (on the web) - 320x570</span></span>

## <a name="personality-menu"></a><span data-ttu-id="1bc69-133">Menu de personalidade</span><span class="sxs-lookup"><span data-stu-id="1bc69-133">Personality menu</span></span>

<span data-ttu-id="1bc69-134">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span><span class="sxs-lookup"><span data-stu-id="1bc69-134">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span></span> <span data-ttu-id="1bc69-135">The following are the current dimensions of the personality menu on Windows and Mac.</span><span class="sxs-lookup"><span data-stu-id="1bc69-135">The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="1bc69-136">No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.</span><span class="sxs-lookup"><span data-stu-id="1bc69-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="1bc69-137">*Figura 4. Menu de personalidade no Windows*</span><span class="sxs-lookup"><span data-stu-id="1bc69-137">*Figure 4. Personality menu on Windows*</span></span>

![Imagem mostrando o menu do personalidade na área de trabalho do Windows](../images/personality-menu-win.png)

<span data-ttu-id="1bc69-139">No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.</span><span class="sxs-lookup"><span data-stu-id="1bc69-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="1bc69-140">*Figura 5. Menu de personalidade no Mac*</span><span class="sxs-lookup"><span data-stu-id="1bc69-140">*Figure 5. Personality menu on Mac*</span></span>

![Imagem mostrando o menu de personalidade na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="1bc69-142">Implementação</span><span class="sxs-lookup"><span data-stu-id="1bc69-142">Implementation</span></span>

<span data-ttu-id="1bc69-143">Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="1bc69-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span> 


## <a name="see-also"></a><span data-ttu-id="1bc69-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="1bc69-144">See also</span></span>

- [<span data-ttu-id="1bc69-145">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1bc69-145">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md) 
- [<span data-ttu-id="1bc69-146">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1bc69-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)

