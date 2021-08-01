---
title: Excel Conjunto de requisitos somente para API JavaScript online
description: Detalhes sobre o conjunto de requisitos do ExcelApiOnline.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8e49d2a9d655e7bc1bf21233bdaf23321d1e48d1
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671874"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel Conjunto de requisitos somente para API JavaScript online

O conjunto de requisitos é um conjunto de requisitos especial que inclui recursos que estão disponíveis apenas `ExcelApiOnline` para Excel na Web. As APIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações comportamentais ou estruturais não documentados) para o aplicativo Excel na Web. `ExcelApiOnline`As APIs são consideradas APIs de "visualização" para outras plataformas (Windows, Mac, iOS) e podem não ser suportadas por nenhuma dessas plataformas.

Quando as APIs no conjunto de requisitos são suportadas em todas as plataformas, elas serão adicionadas ao próximo conjunto de requisitos lançado `ExcelApiOnline` ( `ExcelApi 1.[NEXT]` ). Depois que esse novo requisito for público, essas APIs serão removidas de `ExcelApiOnline` . Pense nisso como um processo de promoção semelhante a uma API que está mudando da visualização para a versão.

> [!IMPORTANT]
> `ExcelApiOnline` é um superconjunto do conjunto de requisitos numerado mais recente.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` é a única versão das APIs somente online. Isso porque Excel na Web sempre terá uma única versão disponível para os usuários que são a versão mais recente.

A tabela a seguir fornece um resumo conciso das APIs, enquanto a tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada das `ExcelApiOnline` APIs atuais.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Exibições de planilha nomeadas | Fornece controle programático de exibições de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>Uso recomendado

Como as APIs só têm suporte Excel na Web, o seu complemento deve verificar se o conjunto de requisitos é suportado antes de `ExcelApiOnline` chamar essas APIs. Isso evita chamar uma API somente online em uma plataforma diferente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Depois que a API está em um conjunto de requisitos entre plataformas, você deve remover ou editar a `isSetSupported` verificação. Isso habilita o recurso do seu complemento em outras plataformas. Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.

> [!IMPORTANT]
> Seu manifesto não pode especificar `ExcelApiOnline 1.1` como um requisito de ativação. Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente incluídas no conjunto `ExcelApiOnline` de requisitos. Para uma lista completa de todas as EXCEL JavaScript (incluindo APIs e APIs lançadas anteriormente), consulte todas as `ExcelApiOnline` [APIs JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true)Excel JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Limpa os critérios de filtro de coluna do AutoFilter.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Ativa esse modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Remove o exibição de planilha da planilha.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Cria uma cópia desse exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Cria um novo exibição de planilha com o nome determinado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Sai do exibição de planilha ativa no momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Obtém a exibição de planilha ativa da planilha no momento.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Obtém o número de exibições de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Obtém uma exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Obtém uma exibição de planilha pelo índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Retorna uma coleção de exibições de planilha presentes na planilha.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [APIs de visualização do JavaScript para Excel](excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
