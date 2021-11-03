---
title: Excel Conjunto de requisitos somente para API JavaScript online
description: Detalhes sobre o conjunto de requisitos do ExcelApiOnline.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ae338b6bd361113ee04ae3dd9076df6c66125345
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681490"
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
| Guias de trabalho vinculadas | Gerencie links entre as guias de trabalho, incluindo o suporte para atualizar e quebrar links de livros de trabalho. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Exibições de planilha nomeadas | Fornece controle programático de exibições de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| Eventos de movimentação de planilha | Detectar quando as planilhas são movidas dentro de uma coleção, a posição da planilha e a origem da alteração. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

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
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Faz uma solicitação para quebrar os links apontando para a lista de trabalho vinculada.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|A URL original apontando para a lista de trabalho vinculada.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Faz uma solicitação para atualizar os dados recuperados da lista de trabalho vinculada.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Quebra todos os links para as guias de trabalho vinculadas.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Faz uma solicitação para atualizar todos os links da workbook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Representa o modo de atualização dos links da agenda de trabalho.|
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
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Exclua várias linhas de uma tabela.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Exclua um número especificado de linhas de uma tabela, começando em um determinado índice.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Retorna uma coleção de guias de trabalho vinculadas.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Retorna uma coleção de exibições de planilha presentes na planilha.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Ocorre quando o nome da planilha é alterado.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Ocorre quando uma planilha é movida por um usuário dentro de uma pasta de trabalho.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Ocorre quando o nome da planilha é alterado na coleção de planilhas.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada na coleção de planilhas.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Obtém a nova posição da planilha, após a movimentação.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Obtém a posição anterior da planilha, antes da movimentação.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Obtém a ID da planilha que foi movida.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Obtém o novo nome da planilha, após a alteração do nome.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Obtém o nome anterior da planilha, antes de o nome ser alterado.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Obtém a ID da planilha com o novo nome.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Obtém o tipo do evento.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Obtém a nova configuração de visibilidade da planilha, após a alteração de visibilidade.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Obtém a configuração de visibilidade anterior da planilha, antes da alteração de visibilidade.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Obtém a ID da planilha cuja visibilidade foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [APIs de visualização do JavaScript para Excel](excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
