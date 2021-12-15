---
title: PowerPoint APIs de visualização do JavaScript
description: Detalhes sobre as próximas POWERPOINT APIs JavaScript.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 406808b4b4ff16df72d9c37468696525c8be642f
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513988"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint APIs de visualização do JavaScript

Novas PowerPoint APIs JavaScript são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são adquiridos.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Gerenciamento de slides | Adiciona suporte para adicionar slides, bem como gerenciar layouts de slides e mestres de slides. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Formas | Adiciona suporte para obter referências às formas em um slide. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as PowerPoint APIs JavaScript atualmente em visualização. Para uma lista completa de todas as POWERPOINT JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)Excel JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Especifica se os marcadores no parágrafo estão visíveis.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Representa o formato de marcador do parágrafo.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Representa o alinhamento horizontal do parágrafo.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Retorna a formatação de preenchimento dessa forma.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Especifica a altura, em pontos, da forma.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|A distância, em pontos, do lado esquerdo da forma até o lado esquerdo do slide.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Retorna a formatação de linha do objeto de forma.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Especifica o nome dessa forma.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Retorna o objeto text frame de uma forma.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior do slide.|
||[tipo](/javascript/api/powerpoint/powerpoint.shape#type)|Retorna o tipo dessa forma.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Especifica a largura, em pontos, da forma.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Especifica a altura, em pontos, da forma.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Especifica a distância, em pontos, do lado esquerdo da forma até o lado esquerdo do slide.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Especifica a distância, em pontos, da borda superior da forma até a borda superior do slide.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Especifica a largura, em pontos, da forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint. GeometricShapeType, opções?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Adiciona uma forma geométrica ao slide.|
||[addLine(connectorType?: PowerPoint. ConnectorType, opções?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Adiciona uma linha ao slide.|
||[addTextBox(text: string, options?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Adiciona uma caixa de texto ao slide com o texto fornecido como o conteúdo.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Representa a cor de primeiro plano de preenchimento da forma no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Define a formatação de preenchimento de um formato com uma cor uniforme.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Especifica a porcentagem de transparência do preenchimento como um valor de 0,0 (opaco) a 1,0 (claro).|
||[tipo](/javascript/api/powerpoint/powerpoint.shapefill#type)|Retorna o tipo de preenchimento da forma.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|Representação de código de cor HTML da cor do texto (por exemplo, "#FF0000" representa vermelho).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Representa o estilo de traço da linha.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Representa o estilo de linha da forma.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Especifica a porcentagem de transparência da linha como um valor de 0,0 (opaco) a 1,0 (claro).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Especifica se a formatação de linha de um elemento de forma está visível.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Representa a espessura da linha, em pontos.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|As configurações de redação automáticas do quadro de texto.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Exclui todo o texto no quadro de texto.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Especifica se o quadro de texto contém texto.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Representa margem direita, em pontos, do quadro de texto.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Representa o alinhamento vertical do quadro de texto.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Determina se as linhas quebram automaticamente para caber o texto dentro da forma.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|Retorna um `ShapeFont` objeto que representa os atributos de fonte para o intervalo de texto.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Retorna um `TextRange` objeto para a subdistragem no intervalo determinado.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Representa o formato de parágrafo do intervalo de texto.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|

## <a name="see-also"></a>Confira também

- [PowerPoint de referência da API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)