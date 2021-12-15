---
title: PowerPoint conjunto de requisitos da API JavaScript 1.3
description: Detalhes sobre o conjunto de requisitos do PowerPointApi 1.3.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 74f17923f7bc8b26416c39bdbbeea9cc0a94029a
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514241"
---
# <a name="whats-new-in-powerpoint-javascript-api-13"></a>Novidades na API JavaScript 1.3 PowerPoint JavaScript

O PowerPointApi 1.3 adicionou suporte adicional para gerenciamento de slides e marcação personalizada.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Gerenciamento de slides](../../powerpoint/add-slides.md) | Adiciona suporte para adicionar slides, bem como gerenciar layouts de slides e mestres de slides. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Marcas](../../powerpoint/tagging-presentations-slides-shapes.md) | Permite que os complementos anexem metadados personalizados, na forma de pares de valores-chave. | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista o PowerPoint de requisitos da API JavaScript 1.3. Para uma lista completa de todas as POWERPOINT JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs javascript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)PowerPoint JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Especifica a ID de um Layout de Slide a ser usado para o novo slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Especifica a ID de um Slide Master a ser usado para o novo slide.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Retorna a coleção `SlideMaster` de objetos que estão na apresentação.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Retorna uma coleção de marcas anexadas à apresentação.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Exclui a forma da coleção de formas.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtém a ID exclusiva da forma.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Retorna uma coleção de marcas na forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Obtém o número de formas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Obtém uma forma usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Obtém uma forma usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Obtém uma forma usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtém o layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Retorna uma coleção de formas no slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Obtém `SlideMaster` o objeto que representa o conteúdo padrão do slide.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Retorna uma coleção de marcas no slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Adiciona um novo slide no final da coleção.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtém a ID exclusiva do layout do slide.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtém o nome do layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|Retorna uma coleção de formas no layout do slide.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Obtém o número de layouts na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Obtém um layout usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Obtém um layout usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Obtém um layout usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtém a ID exclusiva do Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtém a coleção de layouts fornecidos pelo Slide Master para slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtém o nome exclusivo do Slide Master.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|Retorna uma coleção de formas no Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Obtém o número de Slide Masters na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Obtém um Slide Master usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Obtém um Slide Master usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Obtém um Slide Master usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtém a ID exclusiva da marca.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtém o valor da marca.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Adiciona uma nova marca no final da coleção.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Exclui a marca com a `key` determinada nesta coleção.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Obtém o número de marcas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Obtém uma marca usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Obtém uma marca usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Obtém uma marca usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [PowerPoint de referência da API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
