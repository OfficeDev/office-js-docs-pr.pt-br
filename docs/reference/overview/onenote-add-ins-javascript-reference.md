---
title: Visão geral da API JavaScript do OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 53b120fbe2bba3967c1b89699daef6bd452b5c24
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450251"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="9eb3d-102">Visão geral da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9eb3d-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="9eb3d-103">Aplica-se a: OneNote Online</span><span class="sxs-lookup"><span data-stu-id="9eb3d-103">Applies to: OneNote Online</span></span>

<span data-ttu-id="9eb3d-104">Os links a seguir mostram os objetos de alto nível do OneNote disponíveis na API.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-104">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="9eb3d-105">Os link de página dos objetos contêm uma descrição dos respectivos eventos, propriedades e métodos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-105">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="9eb3d-106">Acesse esses links para saber mais.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-106">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="9eb3d-107">[Application](/javascript/api/onenote/onenote.application): o objeto de nível superior usado para acessar todos os objetos do OneNote globalmente endereçados, como o bloco de anotações ativo e a sessão ativa.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-107">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="9eb3d-p102">[Notebook](/javascript/api/onenote/onenote.notebook): um bloco de anotações. Blocos de anotações contêm grupos de seções e seções.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="9eb3d-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): uma coleção de blocos de anotações.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="9eb3d-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): um grupo de seções. Os grupos de seções contêm seções e grupos de seções.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="9eb3d-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): uma coleção de grupos de seção.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="9eb3d-p104">[Section](/javascript/api/onenote/onenote.section): uma seção. As seções contêm páginas.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="9eb3d-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): uma coleção de seções.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="9eb3d-p105">[Page](/javascript/api/onenote/onenote.page): uma página. As páginas contêm objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="9eb3d-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): uma coleção de páginas.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="9eb3d-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): uma região de nível superior em uma página que contém os tipos de conteúdo como estrutura de tópicos ou imagem. Um objeto PageContent pode ser atribuído a uma posição na página.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="9eb3d-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): uma coleção de objetos PageContent, que representam os conteúdos da página.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="9eb3d-p107">[Outline](/javascript/api/onenote/onenote.outline): um contêiner para objetos Paragraph. Uma estrutura de tópicos é um filho direto do objeto PageContent.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="9eb3d-p108">[Image](/javascript/api/onenote/onenote.image): um objeto Image. Um Image pode ser um filho direto de um objeto PageContent ou Paragraph.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="9eb3d-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): um contêiner para o conteúdo visível em uma página. Um parágrafo é um filho direto de uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="9eb3d-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): uma coleção de objetos Paragraph em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="9eb3d-130">[RichText](/javascript/api/onenote/onenote.richtext): um objeto RichText.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-130">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="9eb3d-131">[Table](/javascript/api/onenote/onenote.table): um contêiner para objetos TableRow.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-131">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="9eb3d-132">[TableRow](/javascript/api/onenote/onenote.tablerow): um contêiner para objetos TableCell.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-132">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="9eb3d-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): um conjunto de objetos TableRow em uma Table.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="9eb3d-134">[TableCell](/javascript/api/onenote/onenote.tablecell): um contêiner para objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-134">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="9eb3d-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): um conjunto de objetos TableCell em uma TableRow.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="9eb3d-136">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9eb3d-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="9eb3d-137">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="9eb3d-138">Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento.</span><span class="sxs-lookup"><span data-stu-id="9eb3d-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="9eb3d-139">Para saber mais sobre conjuntos de requisitos da API JavaScript do OneNote, consulte o artigo [Conjuntos de requisitos da API JavaScript do OneNote](../requirement-sets/onenote-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9eb3d-139">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="9eb3d-140">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9eb3d-140">OneNote JavaScript API reference</span></span>

<span data-ttu-id="9eb3d-141">Para saber mais sobre a API JavaScript do OneNote, consulte a [Documentação de referência da API JavaScript do OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="9eb3d-141">For detailed information about the OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="9eb3d-142">Confira também</span><span class="sxs-lookup"><span data-stu-id="9eb3d-142">See also</span></span>

- [<span data-ttu-id="9eb3d-143">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9eb3d-143">OneNote JavaScript API programming overview</span></span>](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="9eb3d-144">Crie seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="9eb3d-144">Build your first OneNote add-in</span></span>](../../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="9eb3d-145">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="9eb3d-145">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="9eb3d-146">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9eb3d-146">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
