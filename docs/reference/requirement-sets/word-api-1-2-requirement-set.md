---
title: Conjunto de requisitos da API JavaScript do Word 1.2
description: Detalhes sobre o conjunto de requisitos do WordApi 1.2
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
---

# <a name="whats-new-in-word-javascript-api-12"></a>Quais são as novidades na API JavaScript do Word 1.2

O WordApi 1.2 adicionou suporte para imagens em linha.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Word 1.2. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos da API JavaScript do Word 1.2 ou anterior, consulte APIs do Word no conjunto de requisitos [1.2 ou anterior](/javascript/api/word?view=word-js-1.2&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|Insere uma imagem no corpo, no local especificado.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|Insere uma imagem embutida no local especificado dentro do controle de conteúdo.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|Exclui a imagem embutida do documento.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|Insere uma quebra no local especificado no documento principal.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|Insere um documento no local especificado.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|Insere HTML no local especificado.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|Insere uma imagem embutida no local especificado.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|Insere um formato OOXML no local especificado.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|Insere um parágrafo no local especificado.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|Insere um texto no local especificado.|
||[paragraph](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-paragraph-member)|Obtém o parágrafo pai que inclui a imagem embutida.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|Seleciona a imagem embutida.|
|[Range](/javascript/api/word/word.range)|[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|Obtém a coleção de objetos de imagem embutida presentes no intervalo.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|Insere uma imagem no local especificado.|

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Conjuntos de requisitos da API JavaScript do Word](word-api-requirement-sets.md)
