---
title: Trabalhar com conteúdo da página do OneNote
description: Saiba como trabalhar com o conteúdo da página do OneNote usando a API JavaScript.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ec8a6a92bf6bf58fac9c3c2d22987bc027414
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720936"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="03834-103">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="03834-103">Work with OneNote page content</span></span>

<span data-ttu-id="03834-104">Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="03834-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- <span data-ttu-id="03834-106">Um objeto Page contém um conjunto de objetos PageContent.</span><span class="sxs-lookup"><span data-stu-id="03834-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="03834-107">Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.</span><span class="sxs-lookup"><span data-stu-id="03834-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="03834-108">Um objeto Outline contém um conjunto de objetos Paragraph.</span><span class="sxs-lookup"><span data-stu-id="03834-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="03834-109">Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.</span><span class="sxs-lookup"><span data-stu-id="03834-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="03834-110">Para criar uma página em branco do OneNote, use um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="03834-110">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="03834-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="03834-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="03834-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="03834-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="03834-113">Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.</span><span class="sxs-lookup"><span data-stu-id="03834-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="03834-114">Page</span><span class="sxs-lookup"><span data-stu-id="03834-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="03834-115">Outline</span><span class="sxs-lookup"><span data-stu-id="03834-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="03834-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="03834-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="03834-p101">O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="03834-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="03834-119">HTML com suporte</span><span class="sxs-lookup"><span data-stu-id="03834-119">Supported HTML</span></span>

<span data-ttu-id="03834-120">A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:</span><span class="sxs-lookup"><span data-stu-id="03834-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="03834-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="03834-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="03834-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="03834-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="03834-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="03834-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="03834-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="03834-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="03834-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="03834-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="03834-126">Importar o HTML para o OneNote consolida o espaço em branco.</span><span class="sxs-lookup"><span data-stu-id="03834-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="03834-127">O conteúdo resultante é colado em uma estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="03834-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="03834-128">O OneNote faz o melhor para traduzir o HTML no conteúdo da página enquanto garante a segurança para os usuários.</span><span class="sxs-lookup"><span data-stu-id="03834-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="03834-129">Os padrões HTML e CSS não correspondem exatamente ao modelo de conteúdo do OneNote, portanto, haverá diferenças em aparências, particularmente com estilos de CSS.</span><span class="sxs-lookup"><span data-stu-id="03834-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="03834-130">Recomendamos usar os objetos JavaScript se for necessário formatar uma formatação específica.</span><span class="sxs-lookup"><span data-stu-id="03834-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="03834-131">Acessar o conteúdo da página</span><span class="sxs-lookup"><span data-stu-id="03834-131">Accessing page contents</span></span>

<span data-ttu-id="03834-p104">Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, chame `navigateToPage($page)`.</span><span class="sxs-lookup"><span data-stu-id="03834-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="03834-134">Metadados, como título, ainda podem ser consultados para qualquer página.</span><span class="sxs-lookup"><span data-stu-id="03834-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="03834-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="03834-135">See also</span></span>

- [<span data-ttu-id="03834-136">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="03834-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="03834-137">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="03834-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="03834-138">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="03834-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="03834-139">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="03834-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
