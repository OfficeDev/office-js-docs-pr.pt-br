---
title: PowerPoint conjunto de requisitos da API JavaScript 1.2
description: Detalhes sobre o conjunto de requisitos do PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Novidades na API JavaScript 1.2 PowerPoint JavaScript

O PowerPointApi 1.2 adicionou suporte para inserir slides de outra apresentação na apresentação atual e para excluir slides.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Inserir e excluir slides](../../powerpoint/insert-slides-into-presentation.md) | Permite a inserção de slides existentes na apresentação atual de outra apresentação, bem como a capacidade de excluir slides. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|

## <a name="api-list"></a>Lista de API

A tabela a seguir lista o PowerPoint de requisitos da API JavaScript 1.2. Para uma lista completa de todas as POWERPOINT JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs do JavaScript PowerPoint JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatação](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-formatting-member)|Especifica qual formatação usar durante a inserção de slides.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-sourceslideids-member)|Especifica os slides da apresentação de origem que serão inseridos na apresentação atual.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-targetslideid-member)|Especifica onde na apresentação os novos slides serão inseridos.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|Insere os slides especificados de uma apresentação na apresentação atual.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slides-member)|Retorna uma coleção ordenada de slides na apresentação.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-delete-member(1))|Exclui o slide da apresentação.|
||[id](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-id-member)|Obtém a ID exclusiva do slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getcount-member(1))|Obtém o número de slides na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitem-member(1))|Obtém um slide usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1))|Obtém um slide usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemornullobject-member(1))|Obtém um slide usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [PowerPoint de referência da API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
