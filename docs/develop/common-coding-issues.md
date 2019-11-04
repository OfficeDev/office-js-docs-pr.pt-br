---
title: Problemas de codificação comuns e comportamentos de plataforma inesperados
description: Uma lista de problemas da plataforma de API JavaScript do Office frequentemente encontrada pelos desenvolvedores.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924791"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Problemas de codificação comuns e comportamentos de plataforma inesperados

Este artigo realça aspectos da API JavaScript do Office que podem resultar em comportamento inesperado ou exigir padrões de codificação específicos para obter o resultado desejado. Se você encontrar um problema que pertença à lista, informe-nos usando o formulário de comentários na parte inferior do artigo.

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>APIs comuns e APIs do Outlook não são baseados em promessa

As [APIs comuns](/javascript/api/office) (aquelas que não estão vinculadas a um host específico do Office) e [APIs do Outlook](/javascript/api/outlook) usam um modelo de programação baseado em retorno de chamada. A interação com o documento subjacente do Office requer uma chamada de leitura ou gravação assíncrona que especifica um retorno de chamada a ser executado quando a operação for concluída. Para obter um exemplo desse padrão, consulte [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).

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

## <a name="some-properties-must-be-set-with-json-structs"></a>Algumas propriedades devem ser definidas com as estruturas JSON

> [!NOTE]
> Esta seção só se aplica às APIs específicas do host para Excel e Word.

Algumas propriedades devem ser definidas como estruturas JSON, em vez de definir suas subpropriedades individuais. Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout). A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;`. Essa instrução gera um erro porque `zoom` não está carregada. Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito. Todas as operações de contexto `zoom`acontecem em, atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.

Esse comportamento difere das [Propriedades de navegação](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format). As propriedades `format` de podem ser definidas usando a navegação de objeto, conforme mostrado aqui:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Você pode identificar uma propriedade que deve ter suas subpropriedades definidas com uma estrutura JSON verificando seu modificador somente leitura. Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente. Propriedades graváveis como `PageLayout.zoom` devem ser definidas com uma estrutura JSON. Em Resumo:

- Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.
- Propriedade writable: as subpropriedades devem ser definidas com uma estrutura JSON (e não podem ser definidas por meio de navegação).

## <a name="excel-range-limits"></a>Limites de intervalo do Excel

Se você estiver criando um suplemento do Excel que usa intervalos, esteja ciente das seguintes limitações de tamanho:

- O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB. `RichAPI.Error` será lançado se esse limite for excedido.
- Um intervalo é limitado a 5 milhões células para operações de conjunto.

Se você espera que a entrada do usuário exceda esses limites, verifique os dados e divida os intervalos em vários objetos. Você também precisará enviar várias `context.sync()` chamadas para evitar que as operações de intervalo menores fiquem novamente em lotes.

O suplemento pode ser capaz de usar o [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar as células estrategicamente em um intervalo maior. Confira [trabalhar com vários intervalos simultaneamente em suplementos do Excel](../excel/excel-add-ins-multiple-ranges.md) para obter mais informações.

## <a name="setting-read-only-properties"></a>Configuração de propriedades somente leitura

As [definições do TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro. O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>Confira também

- [OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): o local para relatar e exibir problemas com a plataforma de suplementos do Office e APIs JavaScript.
- [Estouro de pilha](https://stackoverflow.com/questions/tagged/office-js): o local para solicitar e exibir perguntas de programação sobre as APIs JavaScript do Office. Certifique-se de aplicar a marca "Office-js" à sua pergunta ao postar no estouro de pilha.
- [UserVoice](https://officespdev.uservoice.com/): o local para sugerir novos recursos para a plataforma de suplementos do Office e APIs JavaScript do Office.
