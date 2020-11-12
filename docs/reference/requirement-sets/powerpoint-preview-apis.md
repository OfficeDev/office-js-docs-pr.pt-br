---
title: APIs de visualização JavaScript do PowerPoint
description: Detalhes sobre as APIs JavaScript do PowerPoint em breve.
ms.date: 11/09/2020
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: b53b6638b16b2028342003b9a77aa59e7406d5f3
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996519"
---
# <a name="powerpoint-javascript-preview-apis"></a>APIs de visualização JavaScript do PowerPoint

Novas APIs JavaScript do PowerPoint são primeiro introduzidas em "Preview" e mais tarde se tornam parte de um conjunto de requisitos específico e numerado após o teste suficiente e o feedback do usuário é adquirido.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Inserir e excluir slides | Permite a inserção de slides existentes na apresentação atual de outra apresentação, bem como a capacidade de excluir o sildes. | [Slide. Delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do PowerPoint atualmente em versão prévia. Para obter uma lista completa de todas as APIs JavaScript do PowerPoint (incluindo APIs de prévia e APIs previamente lançadas), confira [todas as APIs JavaScript do PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatação](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Especifica a formatação a ser usada durante a inserção do slide.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Especifica os slides da apresentação de origem que serão inseridos na apresentação atual.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Especifica onde os novos slides serão inseridos na apresentação.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64file: cadeia de caracteres, opções?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Insere os slides especificados de uma apresentação na apresentação atual.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Retorna uma coleção ordenada de slides da apresentação.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Exclui o slide da apresentação.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtém a ID exclusiva do slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Obtém o número de slides na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Obtém um slide usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Obtém um slide usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Obtém um slide usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
