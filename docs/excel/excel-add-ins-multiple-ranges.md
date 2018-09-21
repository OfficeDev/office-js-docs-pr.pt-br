---
title: Trabalhar com vários intervalos simultaneamente em suplementos do Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016455"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Trabalhar com vários intervalos simultaneamente em Excel suplementos (Versão Prévia)

A biblioteca JavaScript do Excel permite ao suplemento executar operações e definir propriedades em vários intervalos simultaneamente. Os intervalos não precisam ser contíguos. Além de tornar o seu código mais simples, esta maneira de configurar uma propriedade é executada de forma muito mais rápida do que configurar a mesma propriedade individualmente para cada um dos intervalos.

> [!NOTE]
> As APIs descritas neste artigo exigem a **versão Office 2016 Click-to-Run 1809 Build 10820.20000** ou posterior. (Talvez você precise ingressar no [programa Office Insider](https://products.office.com/office-insider) para obter uma compilação apropriada.) Além disso, você deve carregar a versão beta da biblioteca do Office JavaScript do [Office. js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Por fim, ainda não temos páginas de referência para essas APIs. Mas o arquivo de definição a seguir tem descrições para elas: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Um conjunto de intervalos (possivelmente não adjacentes) é representado por um objeto `Excel.RangeAreas` . Ele tem propriedades e métodos semelhantes ao tipo `Range` (vários com nomes semelhantes ou iguais), contudo, ajustes foram feitos nos:

- Tipos de dados para as propriedades e no comportamento dos setters e getters.
- Tipos de dados dos parâmetros do método e nos comportamentos do método.
- Valores retornados dos tipos de dados do método.

Alguns exemplos:

- `RangeAreas` tem uma propriedade  `address` que retorna uma sequência de caracteres delimitada por vírgula do intervalo de endereços, em vez de apenas um endereço como na propriedade `Range.address` .
- `RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto  `DataValidation` que representa a validação de dados de todos os intervalos no `RangeAreas`, caso seja consistente. A propriedade é `null` se objetos `DataValidation` idênticos não forem aplicados a todos os intervalos em `RangeAreas`. Esse é um princípio geral, mas não universal, com o objeto `RangeAreas` : *se uma propriedade não tiver valores consistentes em todos os intervalos no `RangeAreas`, então ela é `null`.* Confira [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) para obter mais informações e conhecer algumas exceções.
- `RangeAreas.cellCount` obtém o número total de células em todos os intervalos de `RangeAreas`.
- `RangeAreas.calculate` recalcula as células de todos os intervalos de `RangeAreas`.
- `RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retorna outro objeto `RangeAreas` que representa todas as colunas (ou linhas) em todos os intervalos de `RangeAreas`. Por exemplo, se o `RangeAreas` representa "A1: C4" e "F14:L15", então `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".
- `RangeAreas.copyFrom` pode receber um parâmetro `Range` ou `RangeAreas` que representa o(s) intervalo(s) de origem da operação de cópia.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Lista completa dos membros de Range que também estão disponíveis em RangeAreas

##### <a name="properties"></a>Propriedades

Esteja familiarizado com [as propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) antes de escrever código para ler as propriedades listadas. Há sutilezas para o que é retornado.

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

O tipo `RangeAreas` tem algumas propriedades e métodos que não estão no objeto`Range`. A seguir está uma seleção deles:

- `areas`: Um objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`. O objeto  `RangeCollection` também é novo e é similar a outros objetos da coleção do Excel. Ele tem uma propriedade `items` que é uma matriz de objetos `Range` que representa os intervalos.
- `areaCount`: O número total de intervalos no `RangeAreas`.
- `getOffsetRangeAreas`: Funciona exatamente como [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto que um `RangeAreas` é retornado e contém intervalos que são um deslocamento de um dos intervalos no `RangeAreas` original.

## <a name="create-rangeareas-and-set-properties"></a>Criar RangeAreas e definir propriedades

Você pode criar o objeto  `RangeAreas` de duas formas básicas:

- Chamar `Worksheet.getRanges()` e passar a ele uma sequência de caracteres com um intervalo de endereços delimitados por vírgula. Se algum intervalo que você deseja incluir tiver sido transformado em [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você pode incluir o nome, em vez do endereço, na sequência de caracteres.
- Chame `Workbook.getSelectedRanges()`. Esse método retorna um `RangeAreas` que representa todos os intervalos selecionados na planilha ativa no momento.

Depois que você tiver um objeto `RangeAreas` , você pode criar outros usando os métodos no objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.

> [!NOTE]
> Você não pode adicionar diretamente intervalos adicionais para um objeto `RangeAreas` . Por exemplo, a coleção em `RangeAreas.areas` não tem um método `add` .


> [!WARNING] 
> Não tente adicionar ou excluir membros diretamente na matriz `RangeAreas.areas.items` . Isso levará a um comportamento indesejável em seu código. Por exemplo, é possível adicionar um objeto  `Range` na matriz, mas isso causará erros, porque os métodos e propriedades `RangeAreas` se comportarão como se o novo item não estivesse lá. Por exemplo, a propriedade  `areaCount` não inclui intervalos adicionados dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior do que `areasCount-1`. Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causará bugs: embora o objeto `Range` *esteja* excluído, as propriedades e os métodos do objeto pai `RangeAreas` se comportam, ou tentam, como se ele ainda existisse. Por exemplo, se o seu código chama `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas apresentará um erro, porque o objeto do intervalo não existirá.

Configurar a propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos na coleção `RangeAreas.areas` .

A seguir está um exemplo sobre configuração de propriedade em vários intervalos. A função realça os intervalos **F3:F5** e **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo que você passa para `getRanges` ou facilmente os calcula em tempo de execução. Alguns dos cenários em que isso seria verdadeiro são: 

- O código é executado no contexto de um modelo conhecido.
- O código é executado no contexto de dados importados onde o esquema dos dados é conhecido.

Quando você não sabe em tempo de codificação quais intervalos que você precisa operar, você deve descobri-los em tempo de execução. A próxima seção discute esses cenários.

### <a name="discover-range-areas-programmatically"></a>Descobrir áreas de intervalo programaticamente

Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem que você encontre em tempo de execução os intervalos que você deseja operar com base nas características das células e no tipo dos valores das células. Aqui estão as assinaturas dos métodos dos arquivos de dados TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

A seguir está um exemplo do primeiro caso. Sobre este código, observe:

- Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` somente para aquele intervalo.
- Ele passa em forma de parâmetro para `getSpecialCells`, a versão da sequência de caracteres de um valor da enumeração `Excel.SpecialCellType` . Alguns outros valores que também podem ser passados são "Blanks" para células vazias, "Constants" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células que possuem a mesma formatação condicional como a primeira célula no `usedRange`. A primeira célula é a primeira célula no canto esquerdo superior. Para obter uma lista completa dos valores na enumeração, consulte [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
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

Às vezes, o intervalo não tem *nenhuma* célula com a característica procurada. Se `getSpecialCells` não encontrar nada, ele gera um erro de **ItemNotFound**. Isso desviaria o fluxo de controle para um bloco/método `catch` , se houver um. Se não houver, o erro interrompe a função. Pode haver cenários nos quais gerar um erro é exatamente o que você deseja quando não há nenhuma célula com a característica alvo. 

Mas em cenários onde é normal, mas talvez incomum, não existir nenhuma célula correspondente, seu código deve verificar essa possibilidade e gerencia-la normalmente sem exibir um erro. Para esses cenários, use o método `getSpecialCellsOrNullObject` e teste a propriedade `RangeAreas.isNullObject`. Apresentamos um exemplo a seguir. Observação sobre o código:

- O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, ele nunca é `null` no sentido comum do JavaScript. Mas se nenhuma célula correspondente for encontrada, a propriedade  `isNullObject` do objeto é definida como `true`.
- Ele chama `context.sync` *antes* de testar a propriedade `isNullObject` . Esse é um requisito com todas as propriedades e métodos `*OrNullObject`, porque você sempre precisa carregar e sincronizar uma propriedade a fim de lê-la. No entanto, não é necessário *explicitamente* carregar a propriedade `isNullObject` . Ela é carregada automaticamente pelo `context.sync`, mesmo se `load`  não é chamado no objeto. Para obter mais informações, consulte [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Você pode testar esse código selecionando um intervalo que tenha células sem fórmulas e o executando. Em seguida, selecione um intervalo que tenha pelo menos uma célula com fórmula e o execute novamente.

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

Há um segundo parâmetro opcional, do tipo enumerado `Excel.SpecialCellValueType`, que restringe ainda mais as células de destino. Você pode usá-lo apenas quando você passá "Formulas" ou "Constants" para `getSpecialCells` ou `getSpecialCellsOrNullObject`. O parâmetro especifica que você apenas deseja as células com determinados tipos de valores. Há quatro tipos básicos: "Erro", "Lógico" (que é booleano), "Números" e "Texto". (A enumeração tem outros valores além desses quatro abordados abaixo). A seguir está um exemplo. Sobre este código, observe:

- Ele realçará somente células que têm um valor numérico literal. Ele não irá realçar células que têm uma fórmula (mesmo se o resultado é um número) ou uma célula que contenha um valor booleano, texto ou erro.
- Para testar o código, certifique-se de que a planilha possua algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas células com fórmulas.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Em alguns casos, você precisa operar em mais de um tipo de valor de célula, como valores de texto e valores booleanos ("Lógica"). A enumeração `Excel.SpecialCellValueType` tem valores que permitem que você combine tipos. Por exemplo, "LogicalText" irá marcar todas as células com valores booleanos e todas as células com valores de texto. Você pode combinar dois ou três dos quatro tipos básicos. Os nomes desses valores de enumeração que combinam os tipos básicos sempre estão em ordem alfabética. Portanto para combinar células com valor de erro, valor de texto e valor booleano, use "ErrorLogicalText", e não "LogicalErrorText" ou "TextErrorLogical". O parâmetro padrão "All", combina todos os quatro tipos. O exemplo a seguir realça todas as células com fórmulas que geraram valores booleanos ou números:

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

O tipo `RangeAreas` possui métodos  `getSpecialCells` e `getSpecialCellsOrNullObject` que obtêm os mesmos dois parâmetros. Esses métodos retornam todas as células de destino de todos os intervalos da coleção `RangeAreas.areas` . Há uma pequena diferença no comportamento dos métodos quando chamado em um objeto `RangeAreas` , em vez de um objeto  `Range` : quando você passá "SameConditionalFormat" como o primeiro parâmetro, o método retornará todas as células que tenham a mesma formatação condicional como o célula mais à esquerda no canto superior *do primeiro intervalo na coleção `RangeAreas.areas` *. O mesmo ponto aplica-se a "SameDataValidation": quando passados para `Range.getSpecialCells`, ele retorna todas as células que tenham a mesma regra de validação de dados como a célula mais à esquerda no canto superior *no intervalo*. Mas, quando ele é passado para `RangeAreas.getSpecialCells`, ele retorna todas as células que possuem a mesma regra de validação de dados como a célula mais à esquerda no canto superior *no primeiro intervalo na coleção `RangeAreas.areas`  *.

## <a name="read-properties-of-rangeareas"></a>Propriedades de leitura das RangeAreas

Ler os valores da propriedade de `RangeAreas` requer cuidado, pois uma determinada propriedade pode ter valores diferentes para diferentes intervalos dentro de `RangeAreas`. A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado. Por exemplo, no código a seguir, o código RGB para rosa (`#FFC0CB`) e `true` serão registrados no console porque ambos os intervalos no objeto  `RangeAreas` possuem um preenchimento rosa e ambos são colunas inteiras.

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

As coisas ficam mais complicadas quando a consistência não é possível. O comportamento das propriedades `RangeAreas` segue estes três princípios:

- Uma propriedade booleana de um objeto  `RangeAreas` retorna `false` , a menos que a propriedade seja verdadeira (true) para todos os intervalos membro.
- Propriedades não-booleanas, com exceção da propriedade `address` , retornam `null` , a menos que a propriedade correspondente em todos os intervalos membro tenha o mesmo valor.
- A propriedade  `address` retornará uma sequência de caracteres delimitada por vírgulas dos endereços dos intervalos membro.

Por exemplo, o código a seguir cria um `RangeAreas` em que somente um intervalo é uma coluna inteira e apenas um é preenchido com rosa. O console mostrará `null` para a cor de preenchimento, `false` para a propriedade `isEntireRow` e "Sheet1! F3:F5, Sheet1! H:H"(supondo que o nome da planilha seja "Sheet1") para a propriedade `address`. 

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

- [Principais conceitos da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Objeto RangeAreas (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Este link pode não funcionar enquanto a API estiver em versão prévia. Como alternativa, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)