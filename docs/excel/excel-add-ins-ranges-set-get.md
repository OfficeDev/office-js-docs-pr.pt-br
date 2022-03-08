---
title: Definir e obter o intervalo selecionado usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter o intervalo selecionado usando a API JavaScript Excel JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ff8690d1d79063114441320232bdef2000af71d5
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340726"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a>Definir e obter o intervalo selecionado usando a EXCEL JavaScript

Este artigo fornece exemplos de código que configuram e selecionam o intervalo com a API JavaScript Excel javascript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>Definir o intervalo selecionado

O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:E6");

    range.select();

    await context.sync();
});
```

### <a name="selected-range-b2e6"></a>Intervalo selecionado B2:E6

![Intervalo selecionado Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Obter o intervalo selecionado

O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("address");

    await context.sync();
    
    console.log(`The address of the selected range is "${range.address}"`);
});
```

## <a name="select-the-edge-of-a-used-range"></a>Selecione a borda de um intervalo usado

Os [métodos Range.getRangeEdge](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1)) e [Range.getExtendedRange](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1)) permitem que o seu complemento replique o comportamento dos atalhos de seleção do teclado, selecionando a borda do intervalo usado com base no intervalo selecionado no momento. Para saber mais sobre intervalos usados, consulte [Obter intervalo usado](excel-add-ins-ranges-get.md#get-used-range).

Na captura de tela a seguir, o intervalo usado é a tabela com valores em cada célula, **C5:F12**. As células vazias fora desta tabela estão fora do intervalo usado.

![Uma tabela com dados de C5:F12 Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a>Selecione a célula na borda do intervalo usado atual

O exemplo de código a seguir mostra como usar `Range.getRangeEdge` o método para selecionar a célula na borda mais distante do intervalo usado atual, na direção para cima. Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Up enquanto um intervalo é selecionado.

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    let rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    await context.sync();
});
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a>Antes de selecionar a célula na borda do intervalo usado

A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado. O intervalo usado é uma tabela com dados **em C5:F12**. Dentro desta tabela, o intervalo **D8:E9** está selecionado. Essa seleção é o *estado anterior* , antes de executar o `Range.getRangeEdge` método.

![Uma tabela com dados de C5:F12 Excel. O intervalo D8:E9 está selecionado.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a>Depois de selecionar a célula na borda do intervalo usado

A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**. Dentro desta tabela, o intervalo **D5** é selecionado. Essa seleção é *após o* estado, depois de executar `Range.getRangeEdge` o método para selecionar a célula na borda do intervalo usado na direção para cima.

![Uma tabela com dados de C5:F12 Excel. O intervalo D5 está selecionado.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a>Selecione todas as células do intervalo atual até a borda mais distante do intervalo usado

O exemplo de código a `Range.getExtendedRange` seguir mostra como usar o método para selecionar todas as células do intervalo selecionado no momento até a borda mais distante do intervalo usado, na direção para baixo. Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Shift+Down enquanto um intervalo é selecionado.

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    let extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    await context.sync();
});
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Antes de selecionar todas as células do intervalo atual até a borda do intervalo usado

A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado. O intervalo usado é uma tabela com dados **em C5:F12**. Dentro desta tabela, o intervalo **D8:E9** está selecionado. Essa seleção é o *estado anterior* , antes de executar o `Range.getExtendedRange` método.

![Uma tabela com dados de C5:F12 Excel. O intervalo D8:E9 está selecionado.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Depois de selecionar todas as células do intervalo atual até a borda do intervalo usado

A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**. Dentro desta tabela, o intervalo **D8:E12** está selecionado. Essa seleção é *após o* estado, `Range.getExtendedRange` depois de executar o método para selecionar todas as células do intervalo atual até a borda do intervalo usado na direção para baixo.

![Uma tabela com dados de C5:F12 Excel. O intervalo D8:E12 está selecionado.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Definir e obter valores de intervalo, texto ou fórmulas usando a EXCEL JavaScript](excel-add-ins-ranges-set-get-values.md)
- [Definir o formato de intervalo usando a EXCEL JavaScript](excel-add-ins-ranges-set-format.md)
