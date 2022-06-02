---
title: Conceitos básicos dos tipos de dados da API JavaScript do Excel
description: Conheça os principais conceitos para usar os tipos de dados do Excel no Suplemento do Office.
ms.date: 05/26/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 2259d28bc87e6452e526786c0b32135e4bb27d45
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833903"
---
# <a name="excel-data-types-core-concepts-preview"></a>Principais conceitos dos tipos de dados do Excel (versão prévia)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para trabalhar com tipos de dados. Ele apresenta conceitos fundamentais para o desenvolvimento de tipos de dados.

## <a name="core-concepts"></a>Principais conceitos

Use a propriedade [`Range.valuesAsJson`](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member) para trabalhar com valores de tipo de dados. Essa propriedade é semelhante ao [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member), mas `Range.values` retorna apenas os quatro tipos básicos: cadeia de caracteres, número, booliano ou valores de erro. `Range.valuesAsJson` retorna informações expandidas sobre os quatro tipos básicos e essa propriedade pode retornar tipos de dados como valores numéricos formatados, entidades e imagens da web.

A propriedade `valuesAsJson` retorna um alias de tipo [CellValue](/javascript/api/excel/excel.cellvalue), que é uma [união](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) dos seguintes tipos de dados.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

O objeto [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) é uma [interseção](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) com o restante dos tipos `*CellValue`. Não é um tipo de dados em si. As propriedades do objeto `CellValueExtraProperties` são usadas com todos os tipos de dados para especificar detalhes relacionados à substituição de valores de células.

### <a name="json-schema"></a>Esquema JSON

Cada tipo de dados usa um esquema de metadados JSON projetado para este tipo. Isso define o [CellValueType](/javascript/api/excel/excel.cellvaluetype) dos dados e informações adicionais sobre a célula, tais como `basicValue`, `numberFormat` ou `address`. Cada `CellValueType` tem propriedades disponíveis de acordo com esse tipo. Por exemplo, o tipo `webImage` inclui as propriedades [altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member) e [atribuição](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member). As seções a seguir mostram exemplos de código JSON do valor de número formatado, valor de entidade e tipos de dados de imagem da Web.

O esquema de metadados JSON para cada tipo de dados também inclui uma ou mais propriedades somente leitura que são usadas quando os cálculos encontram cenários incompatíveis, tais como uma versão do Excel que não atende ao requisito mínimo de número de build para o recurso de tipos de dados. A propriedade `basicType` faz parte dos metadados JSON de todos os tipos de dados e é sempre uma propriedade somente leitura. A propriedade `basicType` é usada como um fallback quando o tipo de dados não é suportado ou está formatado incorretamente.

## <a name="formatted-number-values"></a>Valores de número formatados

O objeto [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permite que os suplementos do Excel definam uma propriedade `numberFormat` de um valor. Depois de atribuído, esse formato de número percorre cálculos com o valor e pode ser retornado por funções.

O exemplo de código JSON a seguir mostra o esquema completo de um valor numérico formatado. O valor do número formatado `myDate` no exemplo de código é exibido como **16/1/1990** na interface do usuário do Excel. Se os requisitos mínimos de compatibilidade para o recurso de tipos de dados não forem atendidos, os cálculos usarão o `basicValue` no lugar do número formatado.

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Valores de entidade

Um valor de entidade é um contêiner dos tipos de dados, semelhante a um objeto em programação orientada a objetos. As entidades também suportam matrizes como propriedades de um valor de entidade. O objeto [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) permite que os suplementos definam propriedades como `type`, `text` e `properties`. A propriedade `properties` permite que o valor de entidade defina e contenha tipos de dados adicionais.

As propriedades `basicType` e `basicValue` definem como os cálculos leem esse tipo de dados de entidade se os requisitos mínimos de compatibilidade para usar tipos de dados não forem atendidos. Neste cenário, esse tipo de dados de entidade é exibido como **#VALUE!** erro na interface do usuário do Excel.

O exemplo de código JSON a seguir mostra o esquema completo de um valor de entidade que contém texto, uma imagem, uma data e um valor de texto adicional.

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

Os valores de entidade também oferecem uma propriedade `layouts` que cria um cartão para a entidade. O cartão é exibido como uma janela modal na interface do usuário do Excel e pode exibir informações adicionais contidas no valor da entidade, além do que é visível na célula. Para saber mais, confira [Usar cartões com tipos de dados de valor de entidade](excel-data-types-entity-card.md).

### <a name="linked-entities"></a>Entidades Vinculadas

Os valores de entidade vinculados ou objetos [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) são um tipo de valor de entidade. Esses objetos integram os dados fornecidos por um serviço externo e podem exibir esses dados como um [cartão de entidade](excel-data-types-entity-card.md), como valores de entidade regulares. Os [Tipos de dados de Ações e Geografia](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) disponíveis através da interface do usuário do Excel são valores de entidade vinculados.

## <a name="web-image-values"></a>Valores de imagem da Web

O objeto [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) cria a capacidade de armazenar uma imagem como parte de uma [entidade](#entity-values) ou como um valor independente em um intervalo. Este objeto oferece muitas propriedades, incluindo `address`, `altText` e `relatedImagesAddress`.

As propriedades `basicType` e `basicValue` definem como os cálculos leem o tipo de dados de imagem da Web se os requisitos mínimos de compatibilidade para usar o recurso de tipos de dados não forem atendidos. Neste cenário, esse tipo de dados de imagem da Web é exibido como um **#VALUE!** erro na interface do usuário do Excel.

O exemplo de código JSON a seguir mostra o esquema completo de uma imagem da Web.

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## <a name="improved-error-support"></a>Suporte a erros aprimorado

As APIs de tipos de dados expõem erros existentes da IU do Excel como objetos. Agora que esses erros são acessíveis como objetos, os suplementos podem definir ou recuperar propriedades como `type`, `errorType` e `errorSubType`.

Veja a seguir uma lista de todos os objetos de erro com suporte expandido por meio de tipos de dados.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Cada um dos objetos de erro pode acessar uma enumeração por meio da propriedade `errorSubType`, e essa enumeração contém dados adicionais sobre o erro. Por exemplo, o objeto de erro `BlockedErrorCellValue` pode acessar a enumeração [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype). O enumeração `BlockedErrorCellValueSubType` oferece dados adicionais sobre o que causou o erro.

## <a name="see-also"></a>Confira também

- [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
- [Usar cartões com tipos de dados de valor de entidade](excel-data-types-entity-card.md)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Funções e tipos de dados personalizados](custom-functions-data-types-concepts.md)
