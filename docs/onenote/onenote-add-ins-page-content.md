---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 246c864cfb6a63b5f78da8c1189ac5545411168c
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505661"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="9840e-102">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="9840e-102">Work with OneNote page content</span></span> 

<span data-ttu-id="9840e-103">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="9840e-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto de uma página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="9840e-105">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="9840e-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="9840e-106">Um objeto PageContent contém um tipo de conteúdo Outline, Image, ou Other.</span><span class="sxs-lookup"><span data-stu-id="9840e-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="9840e-107">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="9840e-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="9840e-108">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="9840e-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="9840e-109">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="9840e-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="9840e-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="9840e-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [<span data-ttu-id="9840e-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="9840e-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

<span data-ttu-id="9840e-112">Depois, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como Page.addOutline e Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="9840e-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="9840e-113">Page</span><span class="sxs-lookup"><span data-stu-id="9840e-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [<span data-ttu-id="9840e-114">Outline</span><span class="sxs-lookup"><span data-stu-id="9840e-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [<span data-ttu-id="9840e-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="9840e-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

<span data-ttu-id="9840e-p101">O conteúdo e a estrutura de uma página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="9840e-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="9840e-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="9840e-118">Supported HTML</span></span>

<span data-ttu-id="9840e-119">A API JavaScript de suplementos do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo de uma página:</span><span class="sxs-lookup"><span data-stu-id="9840e-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="9840e-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="9840e-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="9840e-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="9840e-121">`<ul>`, `<ol>`, `<li>`</span></span> 
- <span data-ttu-id="9840e-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="9840e-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="9840e-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="9840e-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="9840e-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="9840e-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="9840e-125">Acessar o conteúdo de uma página</span><span class="sxs-lookup"><span data-stu-id="9840e-125">Accessing page contents</span></span>

<span data-ttu-id="9840e-p102">Só é possível acessar o *Conteúdo de uma Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="9840e-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="9840e-128">Metadados, como títulos, podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="9840e-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="9840e-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="9840e-129">See also</span></span>

- [<span data-ttu-id="9840e-130">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9840e-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="9840e-131">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="9840e-131">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="9840e-132">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="9840e-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="9840e-133">Visão geral da plataforma de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="9840e-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
