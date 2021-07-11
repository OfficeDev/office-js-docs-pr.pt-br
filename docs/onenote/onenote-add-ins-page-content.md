---
title: Trabalhar com conteúdo da página do OneNote
description: Saiba como trabalhar com o OneNote de página usando a API JavaScript.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f506617bfdbc97e94f8fb16930dfc2a935385d5f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349046"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="71958-103">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="71958-103">Work with OneNote page content</span></span>

<span data-ttu-id="71958-104">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="71958-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote diagrama de modelo de objeto de página.](../images/one-note-om-page.png)

- <span data-ttu-id="71958-106">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="71958-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="71958-107">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="71958-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="71958-108">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="71958-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="71958-109">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="71958-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="71958-110">Para criar uma página de OneNote vazia, use um dos seguintes métodos.</span><span class="sxs-lookup"><span data-stu-id="71958-110">To create an empty OneNote page, use one of the following methods.</span></span>

- [<span data-ttu-id="71958-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="71958-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="71958-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="71958-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="71958-113">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="71958-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="71958-114">Page</span><span class="sxs-lookup"><span data-stu-id="71958-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="71958-115">Outline</span><span class="sxs-lookup"><span data-stu-id="71958-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="71958-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="71958-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="71958-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="71958-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="71958-119">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="71958-119">Supported HTML</span></span>

<span data-ttu-id="71958-120">A OneNote api JavaScript de complemento oferece suporte ao SEGUINTE HTML para criar e atualizar conteúdo de página.</span><span class="sxs-lookup"><span data-stu-id="71958-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content.</span></span>

- <span data-ttu-id="71958-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="71958-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="71958-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="71958-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="71958-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="71958-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="71958-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="71958-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="71958-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="71958-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="71958-126">Importar o HTML para o OneNote consolida o espaço em branco.</span><span class="sxs-lookup"><span data-stu-id="71958-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="71958-127">O conteúdo resultante é colado em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="71958-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="71958-128">OneNote o melhor para converter HTML em conteúdo de página ao mesmo tempo em que garante a segurança para os usuários.</span><span class="sxs-lookup"><span data-stu-id="71958-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="71958-129">Os padrões HTML e CSS não são exatamente OneNote o modelo de conteúdo do OneNote, portanto, haverá diferenças nas aparências, especialmente com estilo CSS.</span><span class="sxs-lookup"><span data-stu-id="71958-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="71958-130">Recomendamos usar os objetos JavaScript se for necessário formatação específica.</span><span class="sxs-lookup"><span data-stu-id="71958-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="71958-131">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="71958-131">Accessing page contents</span></span>

<span data-ttu-id="71958-p104">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, chame `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="71958-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="71958-134">Metadados, como título, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="71958-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="71958-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="71958-135">See also</span></span>

- [<span data-ttu-id="71958-136">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="71958-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="71958-137">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="71958-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="71958-138">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="71958-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="71958-139">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="71958-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
