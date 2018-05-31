---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d05f251a798a7670983187bfa4c80140b30f6147
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438855"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="4047c-102">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="4047c-102">Work with OneNote page content</span></span> 

<span data-ttu-id="4047c-103">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="4047c-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="4047c-105">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="4047c-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="4047c-106">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="4047c-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="4047c-107">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="4047c-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="4047c-108">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="4047c-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="4047c-109">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="4047c-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="4047c-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="4047c-110">Section.addPage</span></span>](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [<span data-ttu-id="4047c-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="4047c-111">Page.insertPageAsSibling</span></span>](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

<span data-ttu-id="4047c-112">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como Page.addOutline e Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="4047c-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="4047c-113">Página</span><span class="sxs-lookup"><span data-stu-id="4047c-113">Page</span></span>](https://dev.office.com/reference/add-ins/onenote/page)
- [<span data-ttu-id="4047c-114">Estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="4047c-114">Outline</span></span>](https://dev.office.com/reference/add-ins/onenote/outline)
- [<span data-ttu-id="4047c-115">Parágrafo</span><span class="sxs-lookup"><span data-stu-id="4047c-115">Paragraph</span></span>](https://dev.office.com/reference/add-ins/onenote/paragraph)

<span data-ttu-id="4047c-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="4047c-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="4047c-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="4047c-118">Supported HTML</span></span>

<span data-ttu-id="4047c-119">A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:</span><span class="sxs-lookup"><span data-stu-id="4047c-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="4047c-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="4047c-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="4047c-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="4047c-121"></span></span> 
- <span data-ttu-id="4047c-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="4047c-122"></span></span>
- <span data-ttu-id="4047c-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="4047c-123"></span></span>
- <span data-ttu-id="4047c-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="4047c-124"></span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="4047c-125">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="4047c-125">Accessing page contents</span></span>

<span data-ttu-id="4047c-p102">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="4047c-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="4047c-128">Metadados, como títulos, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="4047c-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="4047c-129">Veja também</span><span class="sxs-lookup"><span data-stu-id="4047c-129">See also</span></span>

- [<span data-ttu-id="4047c-130">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4047c-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="4047c-131">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4047c-131">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="4047c-132">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="4047c-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="4047c-133">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4047c-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
