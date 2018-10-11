---
title: Trabalhar com vários intervalos simultaneamente em suplementos do Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: 2387be8dc17d85028b1d086cb192ac1accf167d5
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459193"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Trabalhar com vários intervalos simultaneamente em suplementos do Excel (Versão Prévia)

A biblioteca JavaScript do Excel permite que o seu suplemento execute operações e defina propriedades em vários intervalos simultaneamente. Os intervalos não precisam ser contíguos. Além de simplificar o seu código, essa maneira de definir uma propriedade é mais rápida do que configurar a mesma propriedade individualmente para cada um dos intervalos.

> [!NOTE]
> As APIs descritas neste artigo exigem a **versão 1809 do Office 2016 Clique para Executar, Build 10820.20000** ou posterior. (Talvez você precise ingressar no [programa Office Insider](https://products.office.com/office-insider) para obter o build apropriado.) Além disso, você deve carregar a versão beta da biblioteca do JavaScript do Office encontrada em [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Por último, infelizmente ainda não temos páginas de referência para essas APIs. Mas o tipo de arquivo de definição a seguir traz descrições para eles: [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto `Excel.RangeAreas`. Ele tem propriedades e métodos semelhantes ao tipo `Range` (vários com nomes semelhantes ou iguais), mas alguns ajustes foram feitos em:

- Os tipos de dados para as propriedades e o comportamento dos setters e getters.
- Tipos de dados dos parâmetros do método e nos comportamentos do método.
- Valores retornados dos tipos de dados do método.

Alguns exemplos:

- `RangeAreas` tem uma propriedade  `address` que retorna uma sequência de caracteres delimitada por vírgula do intervalo de endereços, em vez de apenas um endereço como na propriedade `Range.address` .
- `RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em `RangeAreas`, caso seja consistente. A propriedade é `null` se objetos `DataValidation` idênticos não forem aplicados a todos os intervalos em `RangeAreas`. Esse é um princípio geral, mas não universal, do objeto `RangeAreas`: *Se uma propriedade não tiver valores consistentes em todos os intervalos em `RangeAreas`, então, é `null`.* Consulte [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) para obter mais informações e algumas exceções.
- `RangeAreas.cellCount` obtém o número total de células em todos os intervalos de `RangeAreas`.
- `RangeAreas.calculate` recalcula as células de todos os intervalos de `RangeAreas`.
- `RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retorna outro objeto `RangeAreas` que representa todas as colunas (ou linhas) em todos os intervalos de `RangeAreas`. Por exemplo, se `RangeAreas` representa "A1:C4" e "F14:L15", então, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".
- `RangeAreas.copyFrom` pode receber um parâmetro `Range` ou `RangeAreas` que representa o(s) intervalo(s) de origem da operação de cópia.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Lista completa dos membros de Range que também estão disponíveis em RangeAreas

##### <a name="properties"></a>Propriedades

Conheça as [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) antes de escrever códigos que leiam as propriedades listadas. Há sutilezas quanto a o que é retornado.

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- format
- isEntireColumn
- isEntireRow
- style
- worksheet

##### <a name="methods"></a>Métodos

Métodos de Range em versão prévia estão marcados.

- calculate()
- clear()
- convertDataTypeToText() (versão prévia)
- convertToLinkedDataType() (versão prévia)
- copyFrom() (versão prévia)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (chamado getOffsetRangeAreas no objeto RangeAreas)
- getSpecialCells() (versão prévia)
- getSpecialCellsOrNullObject() (versão prévia)
- getTables() (versão prévia)
- getUsedRange() (chamado getUsedRangeAreas no objeto RangeAreas)
- getUsedRangeOrNullObject() (chamado getUsedRangeAreasOrNullObject no objeto RangeAreas)
- load()
- set()
- setDirty() (versão prévia)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Propriedades e métodos específicos de RangeArea

O tipo `RangeAreas` tem algumas propriedades e métodos que não estão no objeto `Range`. Veja a seguir uma seleção deles:

- `areas`: Um objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`. O objeto `RangeCollection` também é novo e é semelhante a outros objetos da coleção do Excel. Ele tem uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.
- `areaCount`: O número total de intervalos em `RangeAreas`.
- `getOffsetRangeAreas`: Funciona exatamente como [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto que um `RangeAreas` é retornado e contém intervalos que são um deslocamento de um dos intervalos no `RangeAreas` original.

## <a name="create-rangeareas-and-set-properties"></a>Criar RangeAreas e definir propriedades

Você pode criar o objeto  `RangeAreas` de duas formas básicas:

- Chame `Worksheet.getRanges()` e passe para ele uma sequência de caracteres com endereços de intervalo delimitados por vírgula. Se algum dos intervalos que você deseja incluir tiver sido transformado em um [getNamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você pode incluir o nome, em vez do endereço, na sequência de caracteres.
- Chame `Workbook.getSelectedRanges()`. Esse método retorna `RangeAreas` que representa todos os intervalos selecionados na planilha ativa no momento.

Depois que você tiver um objeto `RangeAreas` , você pode criar outros usando os métodos no objeto que retornam `RangeAreas`, como `getOffsetRangeAreas` e `getIntersection`.

> [!NOTE]
> Você não pode adicionar intervalos adicionais diretamente em um objeto `RangeAreas`. Por exemplo, a coleção em `RangeAreas.areas` não tem um método `add`.


> [!WARNING] 
> Não tente adicionar ou excluir membros da matriz `RangeAreas.areas.items` diretamente. Isso causará um comportamento indesejável no seu código. Por exemplo, é possível inserir um objeto `Range` adicional na matriz, mas isso irá causar erros porque métodos e propriedades `RangeAreas` se comportam como se o novo item não estivesse lá. Por exemplo, a propriedade `areaCount` não inclui intervalos inseridos dessa maneira e `RangeAreas.getItemAt(index)` gera um erro se `index` for maior do que `areasCount-1`. Da mesma forma, excluir um objeto `Range` da matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando o método `Range.delete` causará erros: embora o objeto `Range` *seja* excluído, as propriedades e métodos do objeto `RangeAreas` pai se comportam, ou tentam se comportar, como se ele ainda existisse. Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas apresentará um erro, pois o objeto range não existe mais.

Configurar uma propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos na coleção `RangeAreas.areas` .

A seguir, veja um exemplo de definição de uma propriedade em vários intervalos. A função realça os intervalos **F3:F5** e **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Este exemplo se aplica a cenários nos quais você pode codificar os endereços do intervalo que você passa para `getRanges` ou facilmente calculá-los no tempo de execução. Alguns dos cenários em que isso seria possível incluem: 

- O código é executado no contexto de um modelo conhecido.
- O código é executado no contexto de dados importados onde o esquema dos dados é conhecido.

Quando você não sabe durante a codificação quais os intervalos em que você precisa para operar, você deve descobri-los no tempo de execução. A próxima seção discute esses cenários.

### <a name="discover-range-areas-programmatically"></a>Descobrir áreas de intervalo programaticamente

Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem que você descubra, durante o tempo de execução, os intervalos em que você deseja operar com base nas características das células e no tipo dos valores das células. Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

A seguir, veja um exemplo de uso do primeiro. Sobre este código, observe:

- Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` somente para aquele intervalo.
- Ele passa como um parâmetro para `getSpecialCells` a versão de sequência de caracteres de um valor a partir da enumeração `Excel.SpecialCellType`. Alguns dos outros valores que podem ser passados, em vez disso, são "Blanks" para células vazias, "Constants" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células com a mesma formatação condicional que a primeira célula em `usedRange`. A primeira célula é a célula superior mais à esquerda. Para obter uma lista completa dos valores na enumeração, consulte [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- O método `getSpecialCells` retorna um objeto `RangeAreas`, portanto todas as células com fórmulas serão cor de rosa, mesmo que não sejam contíguas. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Em alguns casos, o intervalo não tem *nenhuma* célula com a característica alvo. Se `getSpecialCells` não encontrar nenhuma, ele gera um erro de **ItemNotFound** . Isso desviaria o fluxo de controle para um bloco/método `catch`, casa haja um. Se não houver, o erro interrompe a função. Pode haver cenários nos quais emitir o erro é exatamente o que você deseja que aconteça, quando não há nenhuma célula com a característica alvo. 

Mas há cenários nos quais é normal, mas talvez incomum, que não haja nenhuma célula correspondente; seu código deve verificar essa possibilidade e lidar com ela sem dificuldades e sem gerar um erro. Para esses cenários, use o método `getSpecialCellsOrNullObject` e teste a propriedade `RangeAreas.isNullObject`. Veja um exemplo a seguir. Nota sobre este código:

- O método `getSpecialCellsOrNullObject` sempre retorna um objeto proxy, isso significa que nunca é `null` no sentido comum do JavaScript. Mas se nenhuma célula correspondente for encontrada, a propriedade `isNullObject` do objeto é definida como `true`.
- Ele chama `context.sync` *antes* de testar a propriedade `isNullObject`. Esse é um requisito de todos os métodos e propriedades `*OrNullObject`, pois você sempre precisa carregar e sincronizar uma propriedade para poder lê-la. No entanto, não é necessário carregar *explicitamente* a propriedade `isNullObject`. Ela é carregado automaticamente por `context.sync` , mesmo que `load` não seja chamado no objeto. Para obter mais informações, consulte [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Para testar esse código, selecione um intervalo que não tenha células com fórmulas e execute-o. Depois, selecione um intervalo que tenha pelo menos uma célula com fórmula e execute-o novamente.

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

Para manter a simplicidade, todos os outros exemplos neste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Restringir as células de destino com tipos de valores de célula

Este é um segundo parâmetro opcional, de `Excel.SpecialCellValueType` tipo enumerado, que restringe ainda mais as células alvo. Você pode usá-lo somente quando passa "Formulas" ou "Constants" para `getSpecialCells` ou `getSpecialCellsOrNullObject`. O parâmetro especifica que você deseja somente células com certos tipos de valores. Existem quatro tipos básicos: "Error", "Logical" (que significa booleano), "Numbers", e "Text". (A enumeração tem outros valores além desses quatro discutidos adiante.) Veja um exemplo a seguir. Sobre este código, note:

- Ele realçará somente células que tenha um valor de número literal. Ele não realçará células que tenham uma fórmula (mesmo que o resultado seja um número), um valor booleano, texto ou células de estado de erro.
- Para testar o código, certifique-se de que a planilha possui algumas células com valores numéricos literais, outras com outros tipos de valores literais e algumas células com fórmulas.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Às vezes, é necessário operar em mais de um tipo de valor de célula, como células com valores todos de texto ou todos booleanos ("Logical"). A enumeração `Excel.SpecialCellValueType` possui valores que permitem que você combine tipos. Por exemplo, "LogicalText" tem como alvo todas as células com valores completamente de texto ou completamente booleanos. Você pode combinar quaisquer dois ou três dos quatro tipos básicos. Os nomes desses valores enumerados que combinam tipos básicos sempre seguem a ordem alfabética. Então, para combinar células com valores de erros, texto e booleanos, use "ErrorLogicalText", não "LogicalErrorText" nem "TextErrorLogical". O parâmetro padrão "all" combina todos os quatro tipos. O exemplo a seguir realça todas as células com fórmulas que produzem valores numéricos ou booleanos.

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
> O parâmetro  `Excel.SpecialCellValueType` só pode ser usado se o parâmetro  `Excel.SpecialCellType` for "Formulas" ou "Constants".

### <a name="get-rangeareas-within-rangeareas"></a>Obter RangeAreas dentro de RangeAreas

O próprio tipo `RangeAreas` também tem métodos `getSpecialCells` e `getSpecialCellsOrNullObject` que usam os mesmos dois parâmetros. Esses métodos retornam todas as células alvo de todos os intervalos do conjunto `RangeAreas.areas`. Há uma pequena diferença no comportamento dos métodos quando chamados em um objeto `RangeAreas`, em vez de um objeto `Range`: quando você passa "SameConditionalFormat" como o primeiro parâmetro, o método retorna todas as células com formatação condicional igual a da célula superior mais à esquerda *do primeiro intervalo da `RangeAreas.areas` coleção*. O mesmo aplica-se a "SameDataValidation": quando passado para `Range.getSpecialCells`, retorna todas as células com a mesma regra de validação de dados que a célula superior mais à esquerda *no intervalo*. Mas quando ele é passado para `RangeAreas.getSpecialCells`, retorna todas as células com a mesma regra de validação de dados que a célula superior mais à esquerda *do primeiro intervalo da `RangeAreas.areas` coleção*.

## <a name="read-properties-of-rangeareas"></a>Propriedades de leitura das RangeAreas

A leitura de valores de propriedade de `RangeAreas` requer cuidado, pois uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de `RangeAreas`. A regra geral é que, se um valor consistente *pode* ser retornado, ele será retornado. Por exemplo, no código a seguir, o código RGB para rosa (`#FFC0CB`) e `true` serão registrados no console pois ambos os intervalos no objeto `RangeAreas` possuem preenchimento rosa e ambos são colunas inteiras.

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

As coisas se complicam quando a consistência não é possível. O comportamento das propriedades `RangeAreas` segue estes três princípios:

- Uma propriedade booleana de um objeto  `RangeAreas` retorna `false` , a menos que a propriedade seja verdadeira (true) para todos os intervalos membros.
- Propriedades não-booleanas, com exceção da propriedade `address` , retornam `null` , a menos que a propriedade correspondente em todos os intervalos membro tenha o mesmo valor.
- A propriedade  `address` retornará uma sequência de caracteres delimitada por vírgulas dos endereços dos intervalos membros.

Por exemplo, o código a seguir cria um `RangeAreas` em que somente um intervalo é uma coluna inteira e apenas um é preenchido com rosa. O console mostrará `null` para a cor de preenchimento, `false` para a propriedade `isEntireRow` e "Sheet1!F3:F5, Sheet1!H:H"(supondo que o nome da planilha seja "Sheet1") para a propriedade `address`. 

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

- [Conceitos de programação fundamentais com a API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Objeto RangeAreas (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Este link pode não funcionar enquanto a API estiver na versão prévia. Como alternativa, consulte [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)).