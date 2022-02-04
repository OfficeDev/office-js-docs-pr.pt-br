---
title: PowerPoint conjunto de requisitos da API JavaScript 1.3
description: Detalhes sobre o conjunto de requisitos do PowerPointApi 1.3.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-13"></a>Novidades na API JavaScript 1.3 PowerPoint JavaScript

O PowerPointApi 1.3 adicionou suporte adicional para gerenciamento de slides e marcação personalizada.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Gerenciamento de slides](../../powerpoint/add-slides.md) | Adiciona suporte para adicionar slides, bem como gerenciar layouts de slides e mestres de slides. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Marcas](../../powerpoint/tagging-presentations-slides-shapes.md) | Permite que os complementos anexem metadados personalizados, na forma de pares de valores-chave. | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista o PowerPoint de requisitos da API JavaScript 1.3. Para uma lista completa de todas as POWERPOINT JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs do JavaScript PowerPoint JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|Especifica a ID de um Layout de Slide a ser usado para o novo slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|Especifica a ID de um Slide Master a ser usado para o novo slide.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|Retorna a coleção de `SlideMaster` objetos que estão na apresentação.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|Retorna uma coleção de marcas anexadas à apresentação.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|Exclui a forma da coleção de formas.|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|Obtém a ID exclusiva da forma.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|Retorna uma coleção de marcas na forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|Obtém o número de formas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|Obtém uma forma usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|Obtém uma forma usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|Obtém uma forma usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|Obtém o layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|Retorna uma coleção de formas no slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|Obtém `SlideMaster` o objeto que representa o conteúdo padrão do slide.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|Retorna uma coleção de marcas no slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|Adiciona um novo slide no final da coleção.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|Obtém a ID exclusiva do layout do slide.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|Obtém o nome do layout do slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|Retorna uma coleção de formas no layout do slide.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|Obtém o número de layouts na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|Obtém um layout usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|Obtém um layout usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|Obtém um layout usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|Obtém a ID exclusiva do Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|Obtém a coleção de layouts fornecidos pelo Slide Master para slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|Obtém o nome exclusivo do Slide Master.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|Retorna uma coleção de formas no Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|Obtém o número de Slide Masters na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|Obtém um Slide Master usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|Obtém um Slide Master usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|Obtém um Slide Master usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|Obtém a ID exclusiva da marca.|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|Obtém o valor da marca.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|Adiciona uma nova marca no final da coleção.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|Exclui a marca com a determinada `key` nesta coleção.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|Obtém o número de marcas na coleção.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|Obtém uma marca usando sua ID exclusiva.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|Obtém uma marca usando seu índice baseado em zero na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|Obtém uma marca usando sua ID exclusiva.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [PowerPoint de referência da API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [Conjuntos de requisitos de API JavaScript do PowerPoint](powerpoint-api-requirement-sets.md)
