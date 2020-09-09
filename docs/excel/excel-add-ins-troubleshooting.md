---
title: Solucionando problemas de suplementos do Excel
description: Saiba como solucionar erros de desenvolvimento em suplementos do Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409372"
---
# <a name="troubleshooting-excel-add-ins"></a>Solucionando problemas de suplementos do Excel

Este artigo discute a solução de problemas exclusivos para o Excel. Use a ferramenta de comentários na parte inferior da página para sugerir outros problemas que podem ser adicionados ao artigo.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Limitações de API quando a pasta de trabalho ativa alterna

Os suplementos para Excel se destinam a operar em uma única pasta de trabalho por vez. Os erros podem ocorrer quando uma pasta de trabalho separada da que está executando o suplemento Obtém o foco. Isso ocorre apenas quando determinados métodos estão no processo de chamada quando o foco é alterado.

As seguintes APIs são afetadas por essa opção de pasta de trabalho:

|API JavaScript do Excel | Erro gerado |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> Isso aplica-se apenas a várias pastas de trabalho do Excel abertas no Windows ou Mac.

## <a name="coauthoring"></a>Coautoria

Veja [coautoria em suplementos do Excel](co-authoring-in-excel-add-ins.md) para padrões a serem usados com eventos em um ambiente de coautoria. O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="see-also"></a>Confira também

- [Solucionar erros de desenvolvimento com suplementos do Office](../testing/troubleshoot-development-errors.md)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
