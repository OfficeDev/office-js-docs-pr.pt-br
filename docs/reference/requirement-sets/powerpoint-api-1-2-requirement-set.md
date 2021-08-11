---
title: PowerPoint Conjunto de requisitos da API JavaScript 1.2
description: Detalhes sobre o conjunto de requisitos do PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 971617bc2bd70525fc3d5adf34fc0ad092ae66f9892ed52f0d83053b142caa10
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098683"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Novidades na API JavaScript 1.2 PowerPoint JavaScript

O PowerPointApi 1.2 adicionou suporte para inserir slides de outra apresentação na apresentação atual e para excluir slides.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Inserir e excluir slides](../../powerpoint/insert-slides-into-presentation.md) | Permite a inserção de slides existentes na apresentação atual de outra apresentação, bem como a capacidade de excluir slides. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Lista de API

A tabela a seguir lista o PowerPoint de requisitos da API JavaScript 1.2. Para uma lista completa de todas as POWERPOINT JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs javascript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)PowerPoint JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatação](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Especifica qual formatação usar durante a inserção de slides.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Especifica os slides da apresentação de origem que serão inseridos na apresentação atual.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Especifica onde na apresentação os novos slides serão inseridos.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Insere os slides especificados de uma apresentação na apresentação atual.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Retorna uma coleção ordenada de slides na apresentação.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Exclui o slide da apresentação.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtém a ID exclusiva do slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Obtém o número de slides na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Obtém um slide usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Obtém um slide usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Obtém um slide usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [PowerPoint Documentação de referência da API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
