---
title: Trabalhar simultaneamente com vários intervalos em suplementos do Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: 37f9c8a9f3127d78e1cc794aea9e6d1502cdeaf9
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270975"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Trabalhar simultaneamente com vários intervalos em suplementos do Excel (Visualização)

A biblioteca de JavaScript do Excel permite que o suplemento realize operações e defina propriedades, em vários intervalos simultaneamente. Os intervalos não precisam ser contíguos. Além de tornar seu código mais simples, essa maneira de definir uma propriedade é executada muito mais rapidamente do que definir a mesma propriedade individualmente para cada um dos intervalos.

> [!NOTE]
> As APIs descritas neste artigo requerem a ** versão 1809 Build 10820.20000 clique para executar do Office 2016** ou posterior. (Talvez seja necessário ingressar o [programa Office Insider](https://products.office.com/office-insider) para obter uma compilação apropriada.) Além disso, você deve carregar a versão beta da biblioteca JavaScript do Office [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Por fim, ainda não temos páginas de referência para essas APIs. Mas o seguinte arquivo de tipo de definição tem descrições para eles: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto `Excel.RangeAreas`. Possui propriedades e métodos semelhantes ao tipo `Range` (muitos com os mesmos nomes ou semelhantes), mas foram feitos ajustes para:

- Os tipos de dados para propriedades e o comportamento dos setters e getters.
- Os tipos de dados dos parâmetros do método e os comportamentos do método.
- Os tipos de dados de forma retornam valores.

Alguns exemplos:

- `RangeAreas` tem uma propriedade `address` que retorna uma cadeia de caracteres delimitada por vírgula de intervalo de endereços, em vez de apenas um endereço como na propriedade`Range.address`.
- `RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em`RangeAreas`, se for consistente. A propriedade é `null` se objetos idênticos `DataValidation` não forem aplicados a todos os intervalos em `RangeAreas`. Esse é um princípio geral, mas não universal com o objeto `RangeAreas`: *se uma propriedade não têm valores consistentes em todos os todos os intervalos em `RangeAreas`, então será `null`.* Ver [ler as propriedades de RangeAreas](#read-properties-of-rangeareas) para mais informações e algumas exceções.
- `RangeAreas.cellCount` é o número total de células em todos os intervalos no `RangeAreas`.
- `RangeAreas.calculate` recalcula as células de todos os intervalos no `RangeAreas`.
- `RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retornar outra `RangeAreas` objeto que representa todas as colunas (ou linhas) em todos os intervalos no `RangeAreas`. Por exemplo, se `RangeAreas` representa "A1: C4" e "F14:L15" em seguida, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".
- `RangeAreas.copyFrom` pode ter o parâmetro `Range` ou `RangeAreas` que representam os intervalos de origem da operação de cópia.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Lista completa de membros do intervalo que também estão disponíveis em RangeAreas

##### <a name="properties"></a>Propriedades

Familiarize-se com as [Propriedades de leitura do RangeAreas](#read-properties-of-rangeareas) antes de escrever o código que lê as propriedades listadas. Existem sutilezas para o que é retornado.

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- formato
- isEntireColumn
- isEntireRow
- style
- planilha

##### <a name="methods"></a>Métodos

Os métodos de intervalo na visualização são marcados.

- calculate()
- clear()
- convertDataTypeToText() (visualização)
- convertToLinkedDataType() (visualização)
- copyFrom() (visualização)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (chamada getOffsetRangeAreas no objeto RangeAreas)
- getSpecialCells() (visualização)
- getSpecialCellsOrNullObject() (visualização)
- getTables() (visualização)
- getUsedRange() (chamada getUsedRangeAreas no objeto RangeAreas)
- getUsedRangeOrNullObject() (chamada getUsedRangeAreasOrNullObject no objeto RangeAreas)
- load()
- set()
- setDirty() (visualização)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Métodos e propriedades específicos do RangeArea

O tipo `RangeAreas` tem alguns métodos e propriedades que não estão no objeto `Range`. Esta é a seleção deles:

- `areas`: O objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`. O objeto `RangeCollection` também é novidade e é semelhante a outros objetos do conjunto do Excel. É uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.
- `areaCount`: O número total de intervalos em `RangeAreas`.
- `getOffsetRangeAreas`: Funciona como [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto pelo fato de que o `RangeAreas` é retornado e contém os intervalos que são todos os deslocamentos de um dos intervalos do `RangeAreas` original.

## <a name="create-rangeareas-and-set-properties"></a>Criar RangeAreas e definir propriedades

Você pode criar o objeto`RangeAreas` de duas maneiras básicas:

- Ligue `Worksheet.getRanges()` e encaminhe-o em uma cadeia de caracteres com endereços de intervalo separado por vírgula. Se algum intervalo que você deseja incluir tiver sido feito em um [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você poderá incluir o nome, em vez do endereço, cadeia de caracteres.
- Chamar `Workbook.getSelectedRanges()`. Esse método retornará um `RangeAreas` representando todos os intervalos selecionados na planilha ativa no momento.

Quando você tiver um objeto `RangeAreas`, você pode criar outros usando os métodos de objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.

> [!NOTE]
> É possível adicionar diretamente intervalos adicionais para um objeto `RangeAreas`. Por exemplo, o conjunto `RangeAreas.areas` não tem um método`add`.


> [!WARNING] 
> Tente adicionar ou excluir membros diretamente à matriz`RangeAreas.areas.items`. Isso levará a um comportamento indesejável no seu código. Por exemplo, é possível enviar um objeto adicional `Range` para a matriz, mas isso causará erros porque as propriedades e métodos `RangeAreas` se comportam como se o novo item não estivesse ali. Por exemplo, a propriedade `areaCount` não inclui intervalos transferidos dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior que `areasCount-1`. Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causa bugs: embora o `Range`objeto* seja *excluído, as propriedades e métodos do objeto pai `RangeAreas` se comportam ou tentam se comportar, como se ele ainda existisse. Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas haverá erro porque o objeto de intervalo desapareceu.

Definir uma propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos no conjunto `RangeAreas.areas`.

A seguir, um exemplo de configuração de uma propriedade em vários intervalos. A função realça os intervalos **F3:F5** e **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo para os quais você passa para `getRanges` ou facilmente calculá-los no tempo de execução. Alguns dos cenários em que isso pode ser verdadeiro incluem: 

- O código é executado no contexto de um modelo conhecido.
- O código é executado no contexto de dados importados, em que o esquema dos dados é conhecido.

Quando você não pode saber no tempo de codificação quais intervalos você precisa operar, você deve descobri-los em tempo de execução. A seção a seguir descreve esses cenários.

### <a name="discover-range-areas-programmatically"></a>Descubra as áreas de intervalo por programação

Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem localizar no tempo de execução os intervalos nos quais você deseja operar, com base nas características das células e no tipo de valores das células. Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Este é um exemplo de como usar a primeira. Sobre este código, observe:

- Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.
- Ele passa como um parâmetro para a versão `getSpecialCells` de seqüência de caracteres de um valor do enum `Excel.SpecialCellType`. Alguns dos outros valores que podem ser passados ​​são "Blanks" para células vazias, "Constantes" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células que possuem a mesma formatação condicional que a primeira célula em `usedRange`. A primeira célula é a célula superior esquerda. Para uma lista completa dos valores na enumeração, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Às vezes, o intervalo não possui *nenhuma* célula com a característica desejada. Se `getSpecialCells` não encontrar nenhuma, ele lançará um erro **ItemNotFound**. Isso iria desviar o fluxo de controle para um bloco / método `catch`, se houver um. Se não houver, o erro interrompe a função. Pode haver cenários em que lançar o erro é exatamente o que você quer que aconteça quando não houver células com a característica de destino. 

Mas em cenários em que é normal, mas talvez incomum, não haver células correspondentes; seu código deve verificar essa possibilidade e lidar com isso sem causar erro. Para essas situações, use o método `getSpecialCellsOrNullObject` e teste a propriedade`RangeAreas.isNullObject`. Apresentamos um exemplo a seguir. Observação sobre o código:

- O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, `null` nunca está no sentido comum do JavaScript. Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.
- Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`. Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la. No entanto, não é necessário carregar *explicitamente* a propriedade`isNullObject`. Será carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto. Para saber mais, confira [ \*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o. Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Restringir as células de destino com tipos de valor de célula

Há um segundo parâmetro opcional, do tipo de enumeração  `Excel.SpecialCellValueType`, que restringe ainda mais as células de destino. Você pode usá-lo somente quando você passar por  "Fórmulas" ou "Constantes" para `getSpecialCells` ou `getSpecialCellsOrNullObject`. O parâmetro especifica que você deseja apenas células com determinados tipos de valores. Há quatro tipos básicos: "Erro", "Lógica" (ou seja, booliano), "Números" e "Texto". (O enum tem outros valores além desses quatro que são discutidos abaixo.) O seguinte é um exemplo. Sobre este código, observe:

- Ele apenas irá realçar células que contêm um valor numérico literal. Ele não destacará as células que têm uma fórmula (mesmo se o resultado for um número) ou células de estado booliano, de texto ou de erro.
- Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). A enumeração `Excel.SpecialCellValueType` tem valores que permitem que você combine tipos. Por exemplo, "LogicalText" segmentará todas as células booleanas e todas com valor de texto. Você pode combinar dois ou três dos quatro tipos básicos. Os nomes desses valores de enumeração que combinam tipos básicos estão sempre em ordem alfabética. Portanto, para combinar células com valor de erro, com valor de texto e valores boolianos, use "ErrorLogicalText", não "LogicalErrorText" ou "TextErrorLogical". O parâmetro padrão de "Todos" combina todos os quatro tipos. O exemplo a seguir destaca todas as células com fórmulas que produzem valores ou números boolianos:

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> O parâmetro `Excel.SpecialCellValueType` só poderá ser usado se o parâmetro `Excel.SpecialCellType` for “Fórmulas” ou “Constantes”

### <a name="get-rangeareas-within-rangeareas"></a>Obter RangeAreas dentro de RangeAreas

O tipo `RangeAreas` em si também possui métodos `getSpecialCells` e `getSpecialCellsOrNullObject` que usam os mesmos dois parâmetros. Esses métodos retornam todas as células de destino de todos os intervalos no conjunto`RangeAreas.areas`. Há uma pequena diferença no comportamento dos métodos quando chamado em um objeto `RangeAreas` em vez de um objeto`Range`: quando você passa "SameConditionalFormat" como o primeiro parâmetro, o método retorna todas as células que têm a mesma formatação condicional que a célula superior esquerda * no primeiro intervalo na `RangeAreas.areas`coleção*. O mesmo ponto se aplica a "SameDataValidation": quando passado para `Range.getSpecialCells`, ele retorna todas as células que possuem a mesma regra de validação de dados que a célula superior esquerda * do intervalo*. Mas quando é passado para `RangeAreas.getSpecialCells`, retorna todas as células que têm a mesma regra de validação de dados que a célula superior esquerda *no primeiro intervalo do conjunto`RangeAreas.areas` *.

## <a name="read-properties-of-rangeareas"></a>Ler propriedades de RangeAreas

A leitura de valores de propriedade `RangeAreas` requer cuidados, porque uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de`RangeAreas`. A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado. Por exemplo, no código a seguir, O código RGB para pink (`#FFC0CB`) e `true` será registrado no console porque ambos os intervalos no objeto `RangeAreas` têm um preenchimento rosa e ambos são colunas inteiras.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

As coisas ficam mais complicadas quando a consistência não é possível. O comportamento das propriedades `RangeAreas` seguem estes três princípios de três:

- Uma propriedade booliana de um `RangeAreas`retorno de objeto `false`, a menos que a propriedade seja verdadeira para todos os intervalos de membro.
- Propriedades não boolianas, com exceção da propriedade `address`, retornam `null`, a menos que a propriedade correspondente em todos os intervalos de membros tenha o mesmo valor.
- A propriedade `address` retorna uma cadeia de caracteres delimitada por vírgulas dos endereços e intervalos dos membros.

Por exemplo, o código a seguir cria um `RangeAreas` no qual apenas um intervalo é uma coluna inteira e apenas um é preenchido com rosa. O console mostrará `null` para a cor de preenchimento `false` para a propriedade `isEntireRow` e "Planilha1! F3:F5, Planilha1! H:H"(supondo que o nome da planilha  seja "Planilha1") para a propriedade`address`. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Objeto RangeAreas (API JavaScript do Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (esse link pode não funcionar enquanto a API está na visualização. Como alternativa, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)