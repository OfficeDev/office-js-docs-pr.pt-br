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
# <a name="work-with-onenote-page-content"></a>Trabalhar com conteúdo da página do OneNote 

Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.

  ![Diagrama do modelo de objeto de uma página do OneNote](../images/one-note-om-page.png)

- Um objeto Page contém um conjunto de objetos PageContent.
- Um objeto PageContent contém um tipo de conteúdo Outline, Image, ou Other.
- Um objeto Outline contém um conjunto de objetos Paragraph.
- Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.

Para criar uma página em branco do OneNote, use um dos seguintes métodos:

- [Section.addPage](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [Page.insertPageAsSibling](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

Depois, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como Page.addOutline e Outline.appendHtml. 

- [Page](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [Outline](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [Paragraph](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

O conteúdo e a estrutura de uma página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.

## <a name="supported-html"></a>HTML com suporte

A API JavaScript de suplementos do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo de uma página:

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>Acessar o conteúdo de uma página

Só é possível acessar o *Conteúdo de uma Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.

Metadados, como títulos, podem ser consultados para qualquer página.

## <a name="see-also"></a>Confira também

- [Visão geral da programação da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma de suplementos do Office](../overview/office-add-ins.md)
