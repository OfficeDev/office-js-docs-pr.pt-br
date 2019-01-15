---
title: Trabalhar com conteúdo da página do OneNote
description: ''
ms.date: 1/10/2019
ms.openlocfilehash: 617c30f2a9a0c72b1c309ce299f388b5a16b983f
ms.sourcegitcommit: 384e217fd51d73d13ccfa013bfc6e049b66bd98c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/11/2019
ms.locfileid: "27896333"
---
# <a name="work-with-onenote-page-content"></a>Trabalhar com conteúdo da página do OneNote

Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.

  ![Diagrama do modelo de objeto da página do OneNote](../images/one-note-om-page.png)

- Um objeto Page contém um conjunto de objetos PageContent.
- Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.
- Um objeto Outline contém um conjunto de objetos Paragraph.
- Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.

Para criar uma página em branco do OneNote, use um dos seguintes métodos:

- [Section.addPage](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [Page.insertPageAsSibling](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.

- [Page](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [Estrutura de tópicos](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [Parágrafo](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.

## <a name="supported-html"></a>HTML com suporte

A API JavaScript do suplemento do OneNote dá suporte ao HTML a seguir para a criação e a atualização do conteúdo da página:

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>`
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>`
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

> [!NOTE]
> Importar o HTML para o OneNote consolida o espaço em branco. O conteúdo resultante é colado em uma estrutura de tópicos.

O OneNote faz o máximo para converter o HTML em conteúdo de página e garantir a segurança dos usuários. Os padrões HTML e CSS não correspondem exatamente ao modelo de conteúdo do OneNote; portanto, haverá diferenças nas aparências, especialmente no estilo CSS. É recomendável usar os objetos do JavaScript se uma formatação específica for necessária.

## <a name="accessing-page-contents"></a>Acessar o conteúdo da página

Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, invoque `navigateToPage($page)`.

Metadados, como títulos, ainda podem ser consultados para qualquer página.

## <a name="see-also"></a>Confira também

- [Visão geral da programação da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
