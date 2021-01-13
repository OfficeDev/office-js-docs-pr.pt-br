---
title: Conjunto de requisitos 1.2 da API JavaScript do PowerPoint
description: Detalhes sobre o conjunto de requisitos do PowerPointApi 1.2.
ms.date: 01/08/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 0f6d1e766de81fef5d071152f6116ab56613ec9d
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49841523"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Quais são as novidades na API JavaScript do PowerPoint 1.2

O PowerPointApi 1.2 adicionou suporte para inserir slides de outra apresentação na apresentação atual e para excluir slides.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Inserir e excluir slides | Permite a inserção de slides existentes na apresentação atual de outra apresentação, bem como a capacidade de excluir slides. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista o conjunto de requisitos 1.2 da API JavaScript do PowerPoint. Para ver uma lista completa de todas as APIs JavaScript do PowerPoint (incluindo APIs de visualização e APIs lançadas anteriormente), confira todas as [APIs JavaScript do PowerPoint.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Classe | Campos | Descrição |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatação](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Especifica a formatação a ser usada durante a inserção do slide.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Especifica os slides da apresentação de origem que serão inseridos na apresentação atual.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Especifica onde os novos slides serão inseridos na apresentação.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Insere os slides especificados de uma apresentação na apresentação atual.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Retorna uma coleção ordenada de slides da apresentação.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Exclui o slide da apresentação.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtém a ID exclusiva do slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Obtém o número de slides na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Obtém um slide usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Obtém um slide usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Obtém um slide usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
