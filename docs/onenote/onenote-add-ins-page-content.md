---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f60cdee7eb549acc0f2c84a1aa9acea7fe77274a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872183"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="2da5a-102">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="2da5a-102">Work with OneNote page content</span></span>

<span data-ttu-id="2da5a-103">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="2da5a-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="2da5a-105">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="2da5a-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="2da5a-106">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="2da5a-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="2da5a-107">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="2da5a-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="2da5a-108">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="2da5a-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="2da5a-109">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="2da5a-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="2da5a-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="2da5a-110">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="2da5a-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="2da5a-111">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="2da5a-112">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="2da5a-112">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="2da5a-113">Page</span><span class="sxs-lookup"><span data-stu-id="2da5a-113">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="2da5a-114">Outline</span><span class="sxs-lookup"><span data-stu-id="2da5a-114">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="2da5a-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="2da5a-115">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="2da5a-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="2da5a-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="2da5a-118">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="2da5a-118">Supported HTML</span></span>

<span data-ttu-id="2da5a-119">A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:</span><span class="sxs-lookup"><span data-stu-id="2da5a-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="2da5a-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="2da5a-120"></span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="2da5a-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="2da5a-121"></span></span>
- <span data-ttu-id="2da5a-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="2da5a-122"></span></span>
- <span data-ttu-id="2da5a-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="2da5a-123"></span></span>
- <span data-ttu-id="2da5a-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="2da5a-124"></span></span>

> [!NOTE]
> <span data-ttu-id="2da5a-125">Importar o HTML para o OneNote consolida o espaço em branco.</span><span class="sxs-lookup"><span data-stu-id="2da5a-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="2da5a-126">O conteúdo resultante é colado em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="2da5a-126">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="2da5a-127">O OneNote faz o melhor para traduzir o HTML no conteúdo da página enquanto garante a segurança para os usuários.</span><span class="sxs-lookup"><span data-stu-id="2da5a-127">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="2da5a-128">Os padrões HTML e CSS não correspondem exatamente ao modelo de conteúdo do OneNote, portanto, haverá diferenças em aparências, particularmente com estilos de CSS.</span><span class="sxs-lookup"><span data-stu-id="2da5a-128">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="2da5a-129">Recomendamos usar os objetos JavaScript se for necessário formatar uma formatação específica.</span><span class="sxs-lookup"><span data-stu-id="2da5a-129">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="2da5a-130">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="2da5a-130">Accessing page contents</span></span>

<span data-ttu-id="2da5a-p104">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, chame `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="2da5a-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="2da5a-133">Metadados, como título, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="2da5a-133">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="2da5a-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="2da5a-134">See also</span></span>

- [<span data-ttu-id="2da5a-135">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="2da5a-135">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="2da5a-136">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="2da5a-136">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="2da5a-137">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="2da5a-137">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="2da5a-138">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2da5a-138">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
