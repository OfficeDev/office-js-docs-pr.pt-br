---
title: Trabalhar com conteúdo da página do OneNote
description: Saiba como trabalhar com o OneNote de página usando a API JavaScript.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 01aa4a65f6f1d7ae8fccf490986c10035d30b0c3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938585"
---
# <a name="work-with-onenote-page-content"></a>Trabalhar com conteúdo da página do OneNote

Na API JavaScript de suplementos do OneNote, o conteúdo da página é representado pelo seguinte modelo de objeto.

  ![OneNote diagrama de modelo de objeto de página.](../images/one-note-om-page.png)

- Um objeto Page contém um conjunto de objetos PageContent.
- Um objeto PageContent contém um tipo de conteúdo de Estrutura de Tópicos, Imagem ou Outro.
- Um objeto Outline contém um conjunto de objetos Paragraph.
- Um objeto Paragraph contém um tipo de conteúdo RichText, Image, Table ou Other.

Para criar uma página de OneNote vazia, use um dos seguintes métodos.

- [Section.addPage](/javascript/api/onenote/onenote.section#addPage_title_)
- [Page.insertPageAsSibling](/javascript/api/onenote/onenote.section#insertSectionAsSibling_location__title_)

Em seguida, use métodos nos seguintes objetos para trabalhar com o conteúdo da página, como `Page.addOutline` e `Outline.appendHtml`.

- [Page](/javascript/api/onenote/onenote.page)
- [Outline](/javascript/api/onenote/onenote.outline)
- [Paragraph](/javascript/api/onenote/onenote.paragraph)

O conteúdo e a estrutura da página do OneNote são representados por HTML. Apenas um subconjunto de HTML é compatível com a criação e a atualização do conteúdo da página, conforme descrito abaixo.

## <a name="supported-html"></a>HTML com suporte

A OneNote api JavaScript de complemento oferece suporte ao SEGUINTE HTML para criar e atualizar conteúdo de página.

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

OneNote o melhor para converter HTML em conteúdo de página ao mesmo tempo em que garante a segurança para os usuários. Os padrões HTML e CSS não são exatamente OneNote o modelo de conteúdo do OneNote, portanto, haverá diferenças nas aparências, especialmente com estilo CSS. Recomendamos usar os objetos JavaScript se for necessário formatação específica.

## <a name="accessing-page-contents"></a>Acessar o conteúdo da página

Só é possível acessar o *Conteúdo da Página* via `Page#load` para a página ativa no momento. Para alterar a página ativa, chame `navigateToPage($page)`.

Metadados, como título, ainda podem ser consultados para qualquer página.

## <a name="see-also"></a>Confira também

- [Visão geral da programação da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
