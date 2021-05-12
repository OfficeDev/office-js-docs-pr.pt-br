---
title: Suplementos de conteúdo do Office
description: Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou do PowerPoint que concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 89c8c7eb96d6b1cf6484133b52acbe7d0ccb5298
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52329993"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="a1857-103">Suplementos de conteúdo do Office</span><span class="sxs-lookup"><span data-stu-id="a1857-103">Content Office Add-ins</span></span>

<span data-ttu-id="a1857-104">Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="a1857-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="a1857-105">Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="a1857-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="a1857-106">Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.</span><span class="sxs-lookup"><span data-stu-id="a1857-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="a1857-107">*Figura 1. Layout típico dos suplementos de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="a1857-107">*Figure 1. Typical layout for content add-ins*</span></span>

![Layout típico para os complementos de conteúdo em um Office aplicativo](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="a1857-109">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="a1857-109">Best practices</span></span>

- <span data-ttu-id="a1857-110">Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a1857-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="a1857-111">Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento (aplica-se apenas a suplementos do Excel e do PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="a1857-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="a1857-112">Variantes</span><span class="sxs-lookup"><span data-stu-id="a1857-112">Variants</span></span>

<span data-ttu-id="a1857-113">Tamanhos de complemento de conteúdo para Excel e PowerPoint no Office desktop e Microsoft 365 são especificados pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="a1857-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Microsoft 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="a1857-114">Menu de personalidade</span><span class="sxs-lookup"><span data-stu-id="a1857-114">Personality menu</span></span>

<span data-ttu-id="a1857-p102">Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="a1857-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="a1857-117">No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.</span><span class="sxs-lookup"><span data-stu-id="a1857-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="a1857-118">*Figura 2. Menu de personalidade no Windows*</span><span class="sxs-lookup"><span data-stu-id="a1857-118">*Figure 2. Personality menu on Windows*</span></span>

![Menu de personalidade de 12x32 pixels Windows área de trabalho](../images/personality-menu-win.png)

<span data-ttu-id="a1857-120">No Mac, o menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço ocupado para 34 x 32 pixels, como mostrado.</span><span class="sxs-lookup"><span data-stu-id="a1857-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="a1857-121">*Figura 3. Menu de personalidade no Mac*</span><span class="sxs-lookup"><span data-stu-id="a1857-121">*Figure 3. Personality menu on Mac*</span></span>

![Menu de personalidade de 34 x 32 pixels na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="a1857-123">Implementação</span><span class="sxs-lookup"><span data-stu-id="a1857-123">Implementation</span></span>

<span data-ttu-id="a1857-124">Para ver um exemplo que implementa um suplemento de conteúdo, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="a1857-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="a1857-125">Considerações sobre o suporte</span><span class="sxs-lookup"><span data-stu-id="a1857-125">Support considerations</span></span>

- <span data-ttu-id="a1857-126">Verifique se o seu Office de usuário funcionará em um [aplicativo Office ou plataforma específico.](../overview/office-add-in-availability.md)</span><span class="sxs-lookup"><span data-stu-id="a1857-126">Check to see if your Office Add-in will work on a [specific Office application or platform](../overview/office-add-in-availability.md).</span></span>
- <span data-ttu-id="a1857-127">Alguns suplementos de conteúdo podem obrigar o usuário a "confiar" nele para ler e gravar no Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="a1857-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="a1857-128">Você pode declarar no manifesto do suplemento quais [níveis de permissão](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) deseja que o usuário tenha.</span><span class="sxs-lookup"><span data-stu-id="a1857-128">You can declare what [level of permissions](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="a1857-p104">Os suplementos de conteúdo são compatíveis com o Excel e PowerPoint nas versões do Office 2013 e posteriores. Se você abrir um suplemento em uma versão do Office não compatível com os suplementos web do Office, eles aparecerão como imagem.</span><span class="sxs-lookup"><span data-stu-id="a1857-p104">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="a1857-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="a1857-131">See also</span></span>

- [<span data-ttu-id="a1857-132">Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a1857-132">Office client application and platform availability for Office Add-ins</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="a1857-133">Fabric Core em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="a1857-133">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="a1857-134">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a1857-134">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="a1857-135">Solicitar permissões para uso da API em suplementos </span><span class="sxs-lookup"><span data-stu-id="a1857-135">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
