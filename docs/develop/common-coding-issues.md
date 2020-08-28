---
title: Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292908"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a>Diretrizes de codificação para problemas comuns e comportamentos de plataforma inesperados

Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado. Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>APIs comuns e APIs do Outlook não são baseados em promessa

As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um aplicativo específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada. A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída. Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).

Esses métodos comuns de API e API do Outlook não retornam [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Portanto, você não pode usar [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) para pausar a execução até que a operação assíncrona seja concluída. Se você precisar `await` de comportamento, você pode encapsule a chamada do método em uma promessa criada explicitamente.

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> A documentação de referência contém a implementação com a promessa do [arquivo. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).

## <a name="some-properties-cannot-be-set-directly"></a>Algumas propriedades não podem ser definidas diretamente

> [!NOTE]
> Esta seção só se aplica às APIs específicas do aplicativo para Excel e Word.

Algumas propriedades não podem ser definidas, apesar de serem graváveis. Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto. Isso ocorre porque a propriedade Parent depende das subpropriedades que têm relações lógicas específicas. Essas propriedades pai devem ser definidas usando a notação literal de objeto para definir o objeto inteiro, em vez de definir as subpropriedades individuais desse objeto. Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout). A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;` . Essa instrução gera um erro porque `zoom` não está carregada. Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito. Todas as operações de contexto acontecem em `zoom` , atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.

Esse comportamento difere das [Propriedades de navegação](application-specific-api-model.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format). As propriedades de `format` podem ser definidas usando a navegação de objeto, conforme mostrado aqui:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente verificando seu modificador somente leitura. Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente. Propriedades graváveis como `PageLayout.zoom` devem ser definidas com um objeto nesse nível. Em Resumo:

- Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.
- Propriedade writable: as subpropriedades não podem ser definidas por meio de navegação (devem ser definidas como parte da atribuição de objeto pai inicial).

## <a name="setting-read-only-properties"></a>Configuração de propriedades somente leitura

As [definições do TypeScript](referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro. O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a>Remover manipuladores de eventos

Manipuladores de eventos devem ser removidos usando o mesmo `RequestContext` em que foram adicionados. Se você precisar que seu suplemento remova um manipulador de eventos durante a execução, será necessário armazenar o objeto Context usado para adicionar o manipulador.

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a>Suporte do Internet Explorer

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a>Problemas específicos do Excel

### <a name="api-limitations-when-the-active-workbook-switches"></a>Limitações de API quando a pasta de trabalho ativa alterna

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

### <a name="coauthoring"></a>Coautoria

Veja [coautoria em suplementos do Excel](../excel/co-authoring-in-excel-add-ins.md) para padrões a serem usados com eventos em um ambiente de coautoria. O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="see-also"></a>Confira também

- [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
- [OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.
- [Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office. Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.
- [UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.
