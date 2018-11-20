---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f44c58ac9cb3502889e280c63538603901b63a88
ms.sourcegitcommit: 86724e980f720ed05359c9525948cb60b6f10128
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/09/2018
ms.locfileid: "26237406"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="cc5c8-102">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="cc5c8-102">Work with OneNote page content</span></span> 

<span data-ttu-id="cc5c8-103">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="cc5c8-105">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="cc5c8-106">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="cc5c8-107">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="cc5c8-108">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="cc5c8-109">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="cc5c8-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="cc5c8-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="cc5c8-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [<span data-ttu-id="cc5c8-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="cc5c8-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

<span data-ttu-id="cc5c8-112">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como Page.addOutline e Outline.appendHtml.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="cc5c8-113">Página</span><span class="sxs-lookup"><span data-stu-id="cc5c8-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [<span data-ttu-id="cc5c8-114">Estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="cc5c8-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [<span data-ttu-id="cc5c8-115">Parágrafo</span><span class="sxs-lookup"><span data-stu-id="cc5c8-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

<span data-ttu-id="cc5c8-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="cc5c8-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="cc5c8-118">Supported HTML</span></span>

<span data-ttu-id="cc5c8-119">A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:</span><span class="sxs-lookup"><span data-stu-id="cc5c8-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="cc5c8-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="cc5c8-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="cc5c8-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="cc5c8-121"></span></span> 
- <span data-ttu-id="cc5c8-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="cc5c8-122"></span></span>
- <span data-ttu-id="cc5c8-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="cc5c8-123"></span></span>
- <span data-ttu-id="cc5c8-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="cc5c8-124"></span></span>

> [!NOTE]
> <span data-ttu-id="cc5c8-125">Importar o HTML para o OneNote consolida o espaço em branco.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="cc5c8-126">O conteúdo resultante é colado em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-126">The resulting content is pasted into one outline.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="cc5c8-127">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="cc5c8-127">Accessing page contents</span></span>

<span data-ttu-id="cc5c8-p103">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-p103">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="cc5c8-130">Metadados, como títulos, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="cc5c8-130">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="cc5c8-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="cc5c8-131">See also</span></span>

- [<span data-ttu-id="cc5c8-132">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="cc5c8-132">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="cc5c8-133">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="cc5c8-133">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="cc5c8-134">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="cc5c8-134">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="cc5c8-135">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cc5c8-135">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
