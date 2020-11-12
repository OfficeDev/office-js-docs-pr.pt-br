---
title: Conjunto de requisitos de API JavaScript do Word 1,2
description: Detalhes sobre o conjunto de requisitos WordApi 1,2
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ee9bf60a3a944a3a01a2ca5aa10d01958e3d3475
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996421"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Quais são as novidades na API JavaScript do Word 1.2

WordApi 1,2 adicionado suporte para imagens embutidas.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos de API JavaScript do Word, 1,2. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Word 1,2 ou anterior, confira [APIs do Word no conjunto de requisitos 1,2 ou anterior](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem no corpo, no local especificado.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem embutida no local especificado dentro do controle de conteúdo.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Exclui a imagem embutida do documento.|
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal.|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no local especificado.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Insere HTML no local especificado.|
||[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem embutida no local especificado.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Insere um formato OOXML no local especificado.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Insere um texto no local especificado.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Obtém o parágrafo pai que inclui a imagem embutida.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Seleciona a imagem embutida.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem no local especificado.|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Obtém a coleção de objetos de imagem embutida presentes no intervalo.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
