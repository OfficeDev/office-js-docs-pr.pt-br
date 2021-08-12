---
title: Conjunto de requisitos da API JavaScript do Word 1.2
description: Detalhes sobre o conjunto de requisitos do WordApi 1.2
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 9576069fba08948a76d3e83b3b1af588aa436ddd7f81c4c71681dc7b3dd5bb15
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087880"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Quais são as novidades na API JavaScript do Word 1.2

O WordApi 1.2 adicionou suporte para imagens em linha.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.2. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.2 ou anterior, consulte APIs do Word no conjunto de requisitos [1.2](/javascript/api/word?view=word-js-1.2&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insere uma imagem no corpo, no local especificado.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insere uma imagem embutida no local especificado dentro do controle de conteúdo.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|Exclui a imagem embutida do documento.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|Insere uma quebra no local especificado no documento principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|Insere um documento no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|Insere HTML no local especificado.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insere uma imagem embutida no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|Insere um texto no local especificado.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Obtém o parágrafo pai que inclui a imagem embutida.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#select_selectionMode_)|Seleciona a imagem embutida.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Insere uma imagem no local especificado.|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|Obtém a coleção de objetos de imagem embutida presentes no intervalo.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
