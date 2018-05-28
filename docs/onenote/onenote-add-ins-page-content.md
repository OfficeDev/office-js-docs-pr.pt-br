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
# <a name="work-with-onenote-page-content"></a>Trabalhar com conte?do da p?gina do OneNote 

Na API JavaScript de suplementos do OneNote, o conte?do da p?gina ? representado pelo seguinte modelo de objeto.

  ![Diagrama do modelo de objeto da p?gina do OneNote](../images/one-note-om-page.png)

- Um objeto Page cont?m um conjunto de objetos PageContent.
- Um objeto PageContent cont?m um tipo de conte?do de Estrutura de T?picos, Imagem ou Outro.
- Um objeto Outline cont?m um conjunto de objetos Paragraph.
- Um objeto Paragraph cont?m um tipo de conte?do RichText, Image, Table ou Other.

Para criar uma p?gina em branco do OneNote, use um dos seguintes m?todos:

- [Section.addPage](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [Page.insertPageAsSibling](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

Em seguida, use m?todos nos seguintes objetos para trabalhar com o conte?do da p?gina, como Page.addOutline e Outline.appendHtml. 

- [P?gina](https://dev.office.com/reference/add-ins/onenote/page)
- [Estrutura de t?picos](https://dev.office.com/reference/add-ins/onenote/outline)
- [Par?grafo](https://dev.office.com/reference/add-ins/onenote/paragraph)

O conte?do e a estrutura da p?gina do OneNote s?o representados por HTML. Apenas um subconjunto de HTML ? compat?vel com a cria??o e a atualiza??o do conte?do da p?gina, conforme descrito abaixo.

## <a name="supported-html"></a>HTML com suporte

A API JavaScript do suplemento do OneNote d? suporte ao HTML a seguir para a cria??o e a atualiza??o do conte?do da p?gina:

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>Acessar o conte?do da p?gina

S? ? poss?vel acessar o *Conte?do da P?gina* via `Page#load` para a p?gina ativa no momento. Para alterar a p?gina ativa, invoque `navigateToPage($page)`.

Metadados, como t?tulos, ainda podem ser consultados para qualquer p?gina.

## <a name="see-also"></a>Veja tamb?m

- [Vis?o geral da programa??o da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Refer?ncia da API JavaScript do OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
