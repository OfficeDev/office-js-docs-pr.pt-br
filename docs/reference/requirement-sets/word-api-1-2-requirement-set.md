---
title: Conjunto de requisitos de API JavaScript do Word 1,2
description: Detalhes sobre o conjunto de requisitos WordApi 1,2
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: a71dc9b5954faaab7317d398d5e4453ecb979721
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430524"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Quais são as novidades na API JavaScript do Word 1.2

WordApi 1,2 adicionado suporte para imagens embutidas.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos de API JavaScript do Word, 1,2. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Word 1,2 ou anterior, confira [APIs do Word no conjunto de requisitos 1,2 ou anterior](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem no corpo, no local especificado. O valor de insertLocation pode ser 'Start' ou 'End'.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem embutida no local especificado dentro do controle de conteúdo. O valor de insertLocation pode ser 'Replace', 'Start' ou 'End'.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|Exclui a imagem embutida do documento.|
||[insertBreak (breaktype: Word. Breaktype, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|Insere uma quebra no local especificado no documento principal. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertFileFromBase64 (base64file: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|Insere um documento no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[Métodoinserthtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|Insere HTML no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem embutida no local especificado. O valor insertLocation pode ser ' replace ', ' before ' ou ' after '.|
||[Métodoinsertooxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|Insere um formato OOXML no local especificado.  O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|Insere um parágrafo no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|Insere um texto no local especificado. O valor de insertLocation pode ser 'Before' ou 'After'.|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|Obtém o parágrafo pai que inclui a imagem embutida. Somente leitura.|
||[selecionar (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|Seleciona a imagem embutida. Isso faz com que o Word role até a seleção.|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Insere uma imagem no local especificado. O valor insertLocation pode ser ' replace ', ' Start ', ' End ', ' before ' ou ' after '.|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|Obtém a coleção de objetos de imagem embutida presentes no intervalo. Somente leitura.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
