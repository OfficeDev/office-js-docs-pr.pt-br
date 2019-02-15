---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 01/10/2019
localization_priority: Normal
ms.openlocfilehash: b92f83ee85943ca4e819b04e2b6fa90e6d4fe4b3
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052718"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="4098e-102">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="4098e-102">Work with OneNote page content</span></span>

<span data-ttu-id="4098e-103">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="4098e-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="4098e-105">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="4098e-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="4098e-106">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="4098e-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="4098e-107">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="4098e-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="4098e-108">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="4098e-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="4098e-109">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="4098e-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="4098e-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="4098e-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="4098e-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="4098e-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="4098e-112">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="4098e-112">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="4098e-113">Page</span><span class="sxs-lookup"><span data-stu-id="4098e-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="4098e-114">Estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="4098e-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="4098e-115">Parágrafo</span><span class="sxs-lookup"><span data-stu-id="4098e-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="4098e-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="4098e-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="4098e-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="4098e-118">Supported HTML</span></span>

<span data-ttu-id="4098e-119">A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:</span><span class="sxs-lookup"><span data-stu-id="4098e-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="4098e-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="4098e-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="4098e-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="4098e-121">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="4098e-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="4098e-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="4098e-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="4098e-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="4098e-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="4098e-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="4098e-125">Importar o HTML para o OneNote consolida o espaço em branco.</span><span class="sxs-lookup"><span data-stu-id="4098e-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="4098e-126">O conteúdo resultante é colado em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="4098e-126">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="4098e-127">O OneNote faz o máximo para converter o HTML em conteúdo de página e garantir a segurança dos usuários.</span><span class="sxs-lookup"><span data-stu-id="4098e-127">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="4098e-128">Os padrões HTML e CSS não correspondem exatamente ao modelo de conteúdo do OneNote; portanto, haverá diferenças nas aparências, especialmente no estilo CSS.</span><span class="sxs-lookup"><span data-stu-id="4098e-128">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="4098e-129">É recomendável usar os objetos do JavaScript se uma formatação específica for necessária.</span><span class="sxs-lookup"><span data-stu-id="4098e-129">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="4098e-130">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="4098e-130">Accessing page contents</span></span>

<span data-ttu-id="4098e-p104">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="4098e-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="4098e-133">Metadados, como títulos, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="4098e-133">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="4098e-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="4098e-134">See also</span></span>

- [<span data-ttu-id="4098e-135">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4098e-135">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="4098e-136">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="4098e-136">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="4098e-137">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="4098e-137">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="4098e-138">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4098e-138">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
