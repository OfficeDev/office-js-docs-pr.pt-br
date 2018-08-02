---
title: Suplementos de conteúdo do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f2632e94e0a797836f73caf0d53fdc0f24bd6790
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703914"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="06eb2-102">Suplementos de conteúdo do Office</span><span class="sxs-lookup"><span data-stu-id="06eb2-102">Content Office Add-ins</span></span>

<span data-ttu-id="06eb2-p101">Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Word, Excel ou PowerPoint. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.</span><span class="sxs-lookup"><span data-stu-id="06eb2-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="06eb2-106">*Figura 1. Layout típico dos suplementos de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="06eb2-106">*Figure 1. Typical layout for content add-ins*</span></span>

![Imagem de exemplo exibindo um layout típico de suplementos de conteúdo.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="06eb2-108">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="06eb2-108">Best practices</span></span>

- <span data-ttu-id="06eb2-109">Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</span><span class="sxs-lookup"><span data-stu-id="06eb2-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="06eb2-110">Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento (aplica-se apenas a suplementos do Word, Excel e PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="06eb2-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="06eb2-111">Variantes</span><span class="sxs-lookup"><span data-stu-id="06eb2-111">Variants</span></span>

<span data-ttu-id="06eb2-112">Os tamanhos dos suplementos de conteúdo para Word, Excel e PowerPoint na área de trabalho do Office 2016 e do Office 365 são especificados pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="06eb2-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="06eb2-113">Menu de personalidade</span><span class="sxs-lookup"><span data-stu-id="06eb2-113">Personality menu</span></span>

<span data-ttu-id="06eb2-p102">Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="06eb2-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="06eb2-116">No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.</span><span class="sxs-lookup"><span data-stu-id="06eb2-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="06eb2-117">*Figura 2. Menu de personalidade no Windows*</span><span class="sxs-lookup"><span data-stu-id="06eb2-117">*Figure 2. Personality menu on Windows*</span></span> 

![Imagem mostrando o menu do personalidade na área de trabalho do Windows](../images/personality-menu-win.png)


<span data-ttu-id="06eb2-119">No Mac, o menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço ocupado para 34 x 32 pixels, como mostrado.</span><span class="sxs-lookup"><span data-stu-id="06eb2-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="06eb2-120">*Figura 3. Menu de personalidade no Mac*</span><span class="sxs-lookup"><span data-stu-id="06eb2-120">*Figure 3. Personality menu on Mac*</span></span>

![Imagem mostrando o menu de personalidade na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="06eb2-122">Implementação</span><span class="sxs-lookup"><span data-stu-id="06eb2-122">Implementation</span></span>

<span data-ttu-id="06eb2-123">Para ver um exemplo que implementa um suplemento de conteúdo, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="06eb2-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="06eb2-124">Considerações sobre o suporte</span><span class="sxs-lookup"><span data-stu-id="06eb2-124">Support considerations</span></span>
- <span data-ttu-id="06eb2-125">Verifique se os suplementos do Office funcionarão em uma [plataforma de host do Office específica](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span><span class="sxs-lookup"><span data-stu-id="06eb2-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="06eb2-126">Alguns suplementos de conteúdo podem obrigar o usuário a "confiar" nele para ler e gravar no Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="06eb2-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="06eb2-127">Você pode declarar no manifesto do suplemento quais [níveis de permissão](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) deseja que o usuário tenha.</span><span class="sxs-lookup"><span data-stu-id="06eb2-127">You can declare what [level of permissions](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="06eb2-128">Os suplementos de conteúdo são compatíveis com o Excel e PowerPoint nas versões do Office 2013 e posteriores.</span><span class="sxs-lookup"><span data-stu-id="06eb2-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="06eb2-129">Se você abrir um suplemento em uma versão do Office não compatível com os suplementos web do Office, eles aparecerão como imagem.</span><span class="sxs-lookup"><span data-stu-id="06eb2-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="06eb2-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="06eb2-130">See also</span></span>
- [<span data-ttu-id="06eb2-131">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="06eb2-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="06eb2-132">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="06eb2-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="06eb2-133">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="06eb2-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [<span data-ttu-id="06eb2-134">Solicitar permissões para uso da API em suplementos do painel de tarefas e conteúdo</span><span class="sxs-lookup"><span data-stu-id="06eb2-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
