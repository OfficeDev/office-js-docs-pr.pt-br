---
title: Trabalhar com intervalos usando a API JavaScript do Excel (avançado)
description: Funções e cenários de objetos de intervalo avançados, como células especiais, remoção de duplicatas e trabalho com datas.
ms.date: 10/13/2020
localization_priority: Normal
ms.openlocfilehash: 144012177e0e070149f6cef825c63392a468773d
ms.sourcegitcommit: 6fa29989dfaec4dfa0f8df3fe5fb038d7afbae30
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/16/2020
ms.locfileid: "48487884"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Trabalhar com intervalos usando a API JavaScript do Excel (avançado)

Este artigo baseia-se em informações em [Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md) fornecendo exemplos de código que mostram como executar tarefas mais avançadas com intervalos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais o `Range` objeto oferece suporte, consulte [objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Trabalhar com datas usando o plug-in Moment-MSDate

A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora. O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel. Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.

O código a seguir mostra como definir o intervalo em ** B4 ** para o carimbo de data/hora de um momento:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível. O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57". Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="work-with-multiple-ranges-simultaneously"></a>Trabalhar com vários intervalos simultaneamente

O objeto [RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o suplemento realize operações em vários intervalos de uma só vez. Esses intervalos poderão ser contíguos, mas não precisam ser. `RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).

## <a name="find-special-cells-within-a-range"></a>Localizar células especiais em um intervalo

Os métodos [Range. getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) e [Range. getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) localizam intervalos com base nas características de suas células e nos tipos de valores de suas células. Os dois métodos retornam `RangeAreas` objetos. Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

O exemplo a seguir usa o `getSpecialCells` método para localizar células com fórmulas. Sobre este código, observe:

- Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.
- O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Se nenhuma célula com característica destino existe no intervalo, `getSpecialCells` exibe um erro **ItemNotFound**. Isso desvia o fluxo de controle para um `catch` bloco, se houver um. Se não houver um `catch` bloco, o erro interromperá o método.

Se você espera que células com característica direcionada sempre deveriam existir, provavelmente desejará o código para gerar um erro se as células não estiverem lá. Se for um cenário válido que não há uma ou mais células correspondentes, o código deve verificar se há essa possibilidade e tratar normalmente sem enviar um erro. Você pode obter esse comportamento com o `getSpecialCellsOrNullObject` método e sua propriedade retornada `isNullObject`. O exemplo a seguir usa esse padrão. Sobre este código, observe:

- O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, `null` nunca está no sentido comum do JavaScript. Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.
- Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`. Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la. No entanto, não é necessário carregar *explicitamente* a propriedade`isNullObject`. Será carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto. Para obter mais informações, consulte [ \* métodos e propriedades do OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).
- Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o. Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

Para manter a simplicidade, todos os outros exemplos deste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.

### <a name="narrow-the-target-cells-with-cell-value-types"></a>Restrinja as células de destino com tipos de valor de célula

As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos aceitam um segundo parâmetro opcional usado para restringir ainda mais as células de destino. Este segundo parâmetro é uma `Excel.SpecialCellValueType` você usar para especificar que você quer apenas células que contêm determinados tipos de valores.

> [!NOTE]
> O `Excel.SpecialCellValueType` parâmetro só pode ser usado se a `Excel.SpecialCellType` está `Excel.SpecialCellType.formulas` ou `Excel.SpecialCellType.constants`.

#### <a name="test-for-a-single-cell-value-type"></a>Teste para um tipo de valor da célula única

O `Excel.SpecialCellValueType` enumeração com esses quatro tipos básicos (além dos outros valores combinados descritos nesta seção posterior):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (ou seja, booliano)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

O exemplo a seguir localiza as células especiais que são constantes numéricos e colore essas células em rosa. Sobre este código, observe:

- Ele apenas irá realçar células que contêm um valor numérico literal. Ele não destacará as células que têm uma fórmula (mesmo se o resultado for um número) ou células de estado booliano, de texto ou de erro.
- Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a>Teste para vários tipos de valores de célula

Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). (`Excel.SpecialCellValueType.logical`). O `Excel.SpecialCellValueType` enumeração tem valores com tipos combinado. Por exemplo, `Excel.SpecialCellValueType.logicalText` segmentará todas as células boolianas e todos os valores de texto. `Excel.SpecialCellValueType.all` é o valor padrão, que não limita os tipos de valor da célula retornados. O exemplo a seguir destaca todas as células com fórmulas que produzem valores ou números boolianos.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="cut-copy-and-paste"></a>Recortar, copiar e colar

### <a name="copy-and-paste"></a>Copy and paste

O método [Range. copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) Replica as ações de **copiar** e **colar** da interface do usuário do Excel. O objeto de intervalo para o qual a função`copyFrom` é chamada é o destino. A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.

O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` tem três parâmetros opcionais.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` especifica quais dados são copiados da origem para o destino.

- `Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos de fórmulas. As entradas que não sejam uma fórmula são copiadas no seu estado original.
- `Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.
- `Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.
- `Excel.RangeCopyType.all` (a opção padrão) copia dados e formatação, preservando as fórmulas das células, se encontradas.

`skipBlanks` define se as células em branco são copiadas para o destino. Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.
As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino. O padrão é false.

`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.
Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.

O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Antes da função precedente ter sido executada.*

![Dados no Excel antes do método Copy do intervalo ter sido executado](../images/excel-range-copyfrom-skipblanks-before.png)

*Após a função precedente ter sido executada.*

![Dados no Excel após a execução do método Copy do intervalo](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a>Células recortar e colar (mover)

O método [Range. MoveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) move as células para um novo local na pasta de trabalho. Esse comportamento de movimentação de célula funciona da mesma forma que quando as células são movidas [arrastando-se a borda do intervalo](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) ou ao pegar as ações **recortar** e **colar** . Tanto a formatação quanto os valores do intervalo são movidos para o local especificado como o `destinationRange` parâmetro.

O exemplo de código a seguir mostra um intervalo que está sendo movido com o `Range.moveTo` método. Observe que, se o intervalo de destino for menor do que a fonte, ele será expandido para abranger o conteúdo de origem.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a>Remover duplicatas

O método [Range. removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) remove linhas com entradas duplicadas nas colunas especificadas. O método passa por todas as linhas no intervalo do índice de valor mais baixo para o índice de valor mais alto no intervalo (de cima para baixo). Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo. Linhas no intervalo abaixo da linha excluída são deslocadas para cima. `removeDuplicates` não afeta a posição de células fora do intervalo.

`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas. Essa matriz é baseada em zero e relativa ao intervalo, não à planilha. O método também utiliza um parâmetro Boolean que especifica se a primeira linha é um cabeçalho. Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas. O `removeDuplicates` método retorna um `RemoveDuplicatesResult` objeto que especifica o número de linhas removidas e o número de linhas exclusivas restantes.

Ao usar o método de um intervalo `removeDuplicates` , lembre-se do seguinte:

- `removeDuplicates` considera valores de célula, não resultados de função. Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.
- Células vazias não serão ignoradas por `removeDuplicates`. O valor de uma célula vazia é tratado como qualquer outro valor. Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.

O exemplo a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Antes da função precedente ter sido executada.*

![Dados no Excel antes da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-before.png)

*Após a função precedente ter sido executada.*

![Dados no Excel após a execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a>Agrupar dados para uma estrutura de tópicos

As linhas ou colunas de um intervalo podem ser agrupadas para criar uma [estrutura de tópicos](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF). Esses grupos podem ser recolhidos e expandidos para ocultar e mostrar as células correspondentes. Isso facilita a análise rápida dos dados de linha principal. Use [Range. Group](/javascript/api/excel/excel.range#group-groupoption-) para tornar esses grupos de estrutura de tópicos.

Uma estrutura de tópicos pode ter uma hierarquia, onde grupos menores estão aninhados em grupos maiores. Isso permite que a estrutura de tópicos seja exibida em diferentes níveis. Alterar o nível de estrutura de tópicos visível pode ser feito programaticamente por meio do método [Worksheet. showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) . Observe que o Excel só oferece suporte a oito níveis de grupos de estrutura de tópicos.

O exemplo de código a seguir mostra como criar uma estrutura de tópicos com dois níveis de grupos para ambas as linhas e colunas. A imagem subsequente mostra os agrupamentos dessa estrutura de tópicos. Observe que, no exemplo de código, os intervalos que estão sendo agrupados não incluem a linha ou coluna do controle de estrutura de tópicos (o "total" para este exemplo). Um grupo define o que será recolhido, não a linha ou coluna com o controle.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![Um intervalo com um contorno de duas dimensões de dois níveis](../images/excel-outline.png)

Para desagrupar um grupo de linhas ou colunas, use o método [Range. Ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) . Isso remove o nível mais externo da estrutura de tópicos. Se vários grupos do mesmo tipo de linha ou coluna estiverem no mesmo nível no intervalo especificado, todos esses grupos serão desagrupados.

## <a name="handle-dynamic-arrays-and-spilling"></a>Manipular matrizes dinâmicas e derramamento

Algumas fórmulas do Excel retornam [matrizes dinâmicas](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531). Eles preenchem os valores de várias células fora da célula original da fórmula. Esse estouro de valor é chamado de "derramar". O suplemento pode localizar o intervalo usado para um despejo com o método [Range. getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) . Há também uma [versão do * OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

O exemplo a seguir mostra uma fórmula básica que copia o conteúdo de um intervalo em uma célula, que é despejada nas células vizinhas. O suplemento então registra o intervalo que contém o despejo.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

Você também pode encontrar a célula responsável por despejar em uma determinada célula usando o método [Range. getSpillParent](/javascript/api/excel/excel.range#getspillparent--) . Observe que `getSpillParent` só funciona quando o objeto Range é uma única célula. A chamada `getSpillParent` em um intervalo com várias células resultará em um erro que será lançado (ou um intervalo nulo que está sendo retornado `Range.getSpillParentOrNullObject` ).

## <a name="get-formula-precedents"></a>Obter precedentes de fórmulas

Uma fórmula do Excel freqüentemente se refere a outras células. Quando uma célula fornece dados para uma fórmula, ela é conhecida como uma fórmula "precedente". Para saber mais sobre os recursos do Excel relacionados às relações entre as células, confira o artigo [exibir as relações entre fórmulas e células](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) . 

Com [Range. getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), seu suplemento pode localizar as células precedentes de uma fórmula. `Range.getDirectPrecedents` Retorna um `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os precedentes na pasta de trabalho. Ele tem um `RangeAreas` objeto separado para cada planilha que contém pelo menos uma fórmula precedente. Confira [trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md) para obter mais informações sobre como trabalhar com o `RangeAreas` objeto.

Na interface do usuário do Excel, o botão **rastrear precedentes** desenha uma seta de células precedentes à fórmula selecionada. Ao contrário do botão de interface do usuário do Excel, o `getDirectPrecedents` método não Desenha setas. 

> [!IMPORTANT]
> O `getDirectPrecedents` método não pode recuperar células precedentes entre pastas de trabalho. 

O exemplo a seguir obtém os precedentes diretos do intervalo ativo e altera a cor do plano de fundo dessas células precedentes para amarelo. 

> [!NOTE]
> O intervalo ativo deve conter uma fórmula que faz referência a outras células na mesma pasta de trabalho para que o realce funcione corretamente. 

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md)
- [Modelo de objeto do JavaScript do Excel em suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
