---
title: Conjunto de requisitos somente da API JavaScript do Excel online
description: Detalhes sobre o conjunto de requisitos do ExcelApiOnline.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 282e11e415d51a6724715091d894df64ebaabfae
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604677"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Conjunto de requisitos somente da API JavaScript do Excel online

O conjunto de requisitos é um conjunto de requisitos especial que inclui recursos que `ExcelApiOnline` estão disponíveis apenas para o Excel na Web. ApIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações comportamentais ou estruturais não documentados) para o Excel no aplicativo Web. `ExcelApiOnline` As APIs são consideradas APIs de "visualização" para outras plataformas (Windows, Mac, iOS) e podem não ser suportadas por nenhuma dessas plataformas.

Quando as APIs no conjunto de requisitos são suportadas em todas as plataformas, elas serão adicionadas ao próximo conjunto de requisitos lançado `ExcelApiOnline` ( `ExcelApi 1.[NEXT]` ). Depois que esse novo requisito for público, essas APIs serão removidas de `ExcelApiOnline` . Pense nisso como um processo de promoção semelhante a uma API que está mudando da visualização para a versão.

> [!IMPORTANT]
> `ExcelApiOnline` é um superconjunto do conjunto de requisitos numerado mais recente.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` é a única versão das APIs somente online. Isso porque o Excel na Web sempre terá uma única versão disponível para os usuários que são a versão mais recente.

A tabela a seguir fornece um resumo conciso das APIs, enquanto a tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada das `ExcelApiOnline` APIs atuais.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Exibições de planilha nomeadas | Fornece controle programático de exibições de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>Uso recomendado

Como as APIs só têm suporte do Excel na Web, o seu complemento deve verificar se o conjunto de requisitos é suportado antes de `ExcelApiOnline` chamar essas APIs. Isso evita chamar uma API somente online em uma plataforma diferente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Depois que a API está em um conjunto de requisitos entre plataformas, você deve remover ou editar a `isSetSupported` verificação. Isso habilita o recurso do seu complemento em outras plataformas. Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.

> [!IMPORTANT]
> Seu manifesto não pode especificar `ExcelApiOnline 1.1` como um requisito de ativação. Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs JavaScript do Excel incluídas no conjunto `ExcelApiOnline` de requisitos. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs e APIs lançadas anteriormente), consulte todas as `ExcelApiOnline` [APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Ativa esse modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Remove o exibição de planilha da planilha.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Cria uma cópia desse exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Cria um novo exibição de planilha com o nome determinado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sai do exibição de planilha ativa no momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtém a exibição de planilha ativa da planilha no momento.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtém o número de exibições de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtém uma exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtém uma exibição de planilha pelo índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Retorna um objeto range que inclui o intervalo atual e até a borda do intervalo, com base na direção fornecida.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Retorna um `RangeAreas` objeto que representa as áreas mescladas nesse intervalo.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Retorna um objeto range que é a célula de borda da região de dados que corresponde à direção fornecida.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize a tabela para o novo intervalo.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Retorna uma coleção de exibições de planilha presentes na planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [APIs de visualização do JavaScript para Excel](excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
