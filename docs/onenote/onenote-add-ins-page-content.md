---
title: Trabalhar com conte?do da p?gina do OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d05f251a798a7670983187bfa4c80140b30f6147
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="ee6d2-102">Trabalhar com conte?do da p?gina do OneNote</span><span class="sxs-lookup"><span data-stu-id="ee6d2-102">Work with OneNote page content</span></span> 

<span data-ttu-id="ee6d2-103">Na API JavaScript de suplementos do OneNote, o conte?do da p?gina ? representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da p?gina do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="ee6d2-105">Um objeto Page cont?m um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="ee6d2-106">Um objeto PageContent cont?m um tipo de conte?do de Estrutura de T?picos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="ee6d2-107">Um objeto Outline cont?m um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="ee6d2-108">Um objeto Paragraph cont?m um tipo de conte?do RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="ee6d2-109">Para criar uma p?gina em branco do OneNote, use um dos seguintes m?todos:</span><span class="sxs-lookup"><span data-stu-id="ee6d2-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="ee6d2-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="ee6d2-110">Section.addPage</span></span>](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [<span data-ttu-id="ee6d2-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="ee6d2-111">Page.insertPageAsSibling</span></span>](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

<span data-ttu-id="ee6d2-112">Em seguida, use m?todos nos seguintes objetos para trabalhar com o conte?do da p?gina, como Page.addOutline e Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="ee6d2-113">P?gina</span><span class="sxs-lookup"><span data-stu-id="ee6d2-113">Page</span></span>](https://dev.office.com/reference/add-ins/onenote/page)
- [<span data-ttu-id="ee6d2-114">Estrutura de t?picos</span><span class="sxs-lookup"><span data-stu-id="ee6d2-114">Outline</span></span>](https://dev.office.com/reference/add-ins/onenote/outline)
- [<span data-ttu-id="ee6d2-115">Par?grafo</span><span class="sxs-lookup"><span data-stu-id="ee6d2-115">Paragraph</span></span>](https://dev.office.com/reference/add-ins/onenote/paragraph)

<span data-ttu-id="ee6d2-p101">O conte?do e a estrutura da p?gina do OneNote s?o representados por HTML. Apenas um subconjunto de HTML ? compat?vel com a cria??o e a atualiza??o do conte?do da p?gina, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="ee6d2-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="ee6d2-118">Supported HTML</span></span>

<span data-ttu-id="ee6d2-119">A API JavaScript do suplemento do OneNote d? suporte ao HTML a seguir para a cria??o e a atualiza??o do conte?do da p?gina:</span><span class="sxs-lookup"><span data-stu-id="ee6d2-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="ee6d2-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="ee6d2-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="ee6d2-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="ee6d2-121"></span></span> 
- <span data-ttu-id="ee6d2-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="ee6d2-122"></span></span>
- <span data-ttu-id="ee6d2-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="ee6d2-123"></span></span>
- <span data-ttu-id="ee6d2-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="ee6d2-124"></span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="ee6d2-125">Acessar o conte?do da p?gina</span><span class="sxs-lookup"><span data-stu-id="ee6d2-125">Accessing page contents</span></span>

<span data-ttu-id="ee6d2-p102">S? ? poss?vel acessar o *Conte?do da P?gina* via `Page#load` para a p?gina ativa no momento. Para alterar a p?gina ativa, invoque `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="ee6d2-128">Metadados, como t?tulos, ainda podem ser consultados para qualquer p?gina.</span><span class="sxs-lookup"><span data-stu-id="ee6d2-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee6d2-129">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="ee6d2-129">See also</span></span>

- [<span data-ttu-id="ee6d2-130">Vis?o geral da programa??o da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ee6d2-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="ee6d2-131">Refer?ncia da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ee6d2-131">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="ee6d2-132">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="ee6d2-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="ee6d2-133">Vis?o geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ee6d2-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
