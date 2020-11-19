---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: ed3f3b8fdf7cf62b6016fe8b03393de0d56dfb33
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132015"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="da122-103">Painéis de tarefas nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="da122-103">Task panes in Office Add-ins</span></span>

<span data-ttu-id="da122-p101">Painéis de tarefas são superfícies de interface que normalmente são exibidas no lado direito da janela no Word, PowerPoint, Excel e Outlook. As painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefa quando não precisar inserir a funcionalidade diretamente no documento.</span><span class="sxs-lookup"><span data-stu-id="da122-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="da122-107">*Figura 1. Layout típico do painel de tarefa*</span><span class="sxs-lookup"><span data-stu-id="da122-107">*Figure 1. Typical task pane layout*</span></span>

![Ilustração exibindo um layout de painel de tarefas típico com guias de seção na parte superior, logotipo da empresa e nome da empresa na parte inferior esquerda e um ícone de configurações no canto inferior direito](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="da122-109">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="da122-109">Best practices</span></span>

|<span data-ttu-id="da122-110">Fazer</span><span class="sxs-lookup"><span data-stu-id="da122-110">Do</span></span>|<span data-ttu-id="da122-111">Não fazer</span><span class="sxs-lookup"><span data-stu-id="da122-111">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="da122-112">Inclua o nome do seu suplemento no título.</span><span class="sxs-lookup"><span data-stu-id="da122-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="da122-113">Não adicione o nome da sua empresa ao título.</span><span class="sxs-lookup"><span data-stu-id="da122-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="da122-114">Use nomes descritivos curtos no título.</span><span class="sxs-lookup"><span data-stu-id="da122-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="da122-115">Não acrescente cadeias de caracteres, como "suplemento", "para Word" ou "para Office", ao título do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="da122-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="da122-116">Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</span><span class="sxs-lookup"><span data-stu-id="da122-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="da122-117">Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.</span><span class="sxs-lookup"><span data-stu-id="da122-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||

## <a name="variants"></a><span data-ttu-id="da122-118">Variantes</span><span class="sxs-lookup"><span data-stu-id="da122-118">Variants</span></span>

<span data-ttu-id="da122-p102">As imagens a seguir mostram os vários tamanhos de painel de tarefas com a faixa de opções do aplicativo do Office em uma resolução 1366x768. Para o Excel, é necessário espaço vertical adicional para acomodar a barra de fórmulas.</span><span class="sxs-lookup"><span data-stu-id="da122-p102">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="da122-121">*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*</span><span class="sxs-lookup"><span data-stu-id="da122-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![Diagrama exibindo os tamanhos de painel de tarefas da área de trabalho na Resolução 1366x768](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="da122-123">Excel-320 x 455 pixels</span><span class="sxs-lookup"><span data-stu-id="da122-123">Excel - 320x455 pixels</span></span>
- <span data-ttu-id="da122-124">PowerPoint-320 x 531 pixels</span><span class="sxs-lookup"><span data-stu-id="da122-124">PowerPoint - 320x531 pixels</span></span>
- <span data-ttu-id="da122-125">Pixels do Word-320 x 531</span><span class="sxs-lookup"><span data-stu-id="da122-125">Word - 320x531 pixels</span></span>
- <span data-ttu-id="da122-126">Pixels do Outlook-348 x 535</span><span class="sxs-lookup"><span data-stu-id="da122-126">Outlook - 348x535 pixels</span></span>

<br/>

<span data-ttu-id="da122-127">*Figura 3. Tamanhos de painel de tarefas do Office*</span><span class="sxs-lookup"><span data-stu-id="da122-127">*Figure 3. Office task pane sizes*</span></span>

![Diagrama exibindo os tamanhos de painel de tarefas na Resolução 1366x768](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="da122-129">Excel-350 x 378 pixels</span><span class="sxs-lookup"><span data-stu-id="da122-129">Excel - 350x378 pixels</span></span>
- <span data-ttu-id="da122-130">PowerPoint-348 x 391 pixels</span><span class="sxs-lookup"><span data-stu-id="da122-130">PowerPoint - 348x391 pixels</span></span>
- <span data-ttu-id="da122-131">Pixels do Word-329 x 445</span><span class="sxs-lookup"><span data-stu-id="da122-131">Word - 329x445 pixels</span></span>
- <span data-ttu-id="da122-132">Outlook (na Web)-320 x 570 pixels</span><span class="sxs-lookup"><span data-stu-id="da122-132">Outlook (on the web) - 320x570 pixels</span></span>

## <a name="personality-menu"></a><span data-ttu-id="da122-133">Menu de personalidade</span><span class="sxs-lookup"><span data-stu-id="da122-133">Personality menu</span></span>

<span data-ttu-id="da122-p103">Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="da122-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="da122-136">No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.</span><span class="sxs-lookup"><span data-stu-id="da122-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="da122-137">*Figura 4. Menu de personalidade no Windows*</span><span class="sxs-lookup"><span data-stu-id="da122-137">*Figure 4. Personality menu on Windows*</span></span>

![Diagrama mostrando o menu de personalidade na área de trabalho do Windows](../images/personality-menu-win.png)

<span data-ttu-id="da122-139">No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.</span><span class="sxs-lookup"><span data-stu-id="da122-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="da122-140">*Figura 5. Menu de personalidade no Mac*</span><span class="sxs-lookup"><span data-stu-id="da122-140">*Figure 5. Personality menu on Mac*</span></span>

![Diagrama mostrando o menu de personalidade na área de trabalho Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="da122-142">Implementação</span><span class="sxs-lookup"><span data-stu-id="da122-142">Implementation</span></span>

<span data-ttu-id="da122-143">Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="da122-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="da122-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="da122-144">See also</span></span>

- [<span data-ttu-id="da122-145">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="da122-145">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="da122-146">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="da122-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
