---
title: Conceitos básicos dos tipos de dados da API JavaScript do Excel
description: Conheça os principais conceitos para usar os tipos de dados do Excel no Suplemento do Office.
ms.date: 10/14/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 65a69838500733f8be08a15a99baa167a946b82a
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607447"
---
# <a name="excel-data-types-core-concepts"></a>Conceitos básicos dos tipos de dados do Excel

Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para trabalhar com tipos de dados. Ele apresenta conceitos fundamentais para o desenvolvimento de tipos de dados.

## <a name="the-valuesasjson-property"></a>A propriedade `valuesAsJson`

A `valuesAsJson` propriedade (ou o singular `valueAsJson` para [NamedItem](/javascript/api/excel/excel.nameditem)) é integral para a criação de tipos de dados no Excel. Essa propriedade é uma expansão de `values` propriedades, como [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member). As propriedades `values` e `valuesAsJson` são usadas para acessar o valor em uma célula, mas a propriedade `values` retorna apenas um dos quatro tipos básicos: cadeia de caracteres, número, booliano ou error (como cadeia de caracteres). Por outro lado, `valuesAsJson` retorna informações expandidas sobre os quatro tipos básicos e essa propriedade pode retornar tipos de dados, como valores numéricos formatados, entidades e imagens da web.

Os objetos a seguir oferecem a propriedade `valuesAsJson`.

- [NamedItem](/javascript/api/excel/excel.nameditem) (as `valueAsJson`)
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> Alguns valores de célula mudam com base na localidade de um usuário. A propriedade `valuesAsJsonLocal` oferece suporte à localização e está disponível em todos os mesmos objetos que `valuesAsJson`.

## <a name="cell-values"></a>Valores da célula

A propriedade `valuesAsJson` retorna um alias de tipo [CellValue](/javascript/api/excel/excel.cellvalue), que é uma [união](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) dos seguintes tipos de dados.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

O alias de tipo `CellValue` também retorna o objeto [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties), que é uma [interseção](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) com o restante dos tipos `*CellValue`. Não é um tipo de dados em si. As propriedades do objeto `CellValueExtraProperties` são usadas com todos os tipos de dados para especificar detalhes relacionados à substituição de valores de células.

### <a name="json-schema"></a>Esquema JSON

Cada tipo de valor de célula retornado por `valuesAsJson` usa um esquema de metadados JSON projetado para esse tipo. Junto com propriedades adicionais exclusivas para cada tipo de dados, todos esses esquemas de metadados JSON têm as propriedades `type`, `basicType` e `basicValue` em comum.

O `type` define o [CellValueType](/javascript/api/excel/excel.cellvaluetype) dos dados. Ele `basicType` é sempre somente leitura e é usado como um fallback quando o tipo de dados não tem suporte ou é formatado incorretamente. O `basicValue` corresponde ao valor que seria retornado pela propriedade `values`. O `basicValue` é usado como substituto quando os cálculos encontram cenários incompatíveis, como uma versão mais antiga do Excel que não oferece suporte ao recurso de tipos de dados. É `basicValue` somente leitura para `ArrayCellValue`, `EntityCellValue`e `LinkedEntityCellValue`tipos `WebImageCellValue` de dados.

Além dos três campos que todos os tipos de dados compartilham, o esquema de metadados JSON para cada `*CellValue` tem propriedades disponíveis de acordo com esse tipo. Por exemplo, o tipo [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) inclui as propriedades `altText` e `attribution`, enquanto o tipo [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) oferece os campos `properties` e `text`.

As seções a seguir mostram exemplos de código JSON do valor de número formatado, valor de entidade e tipos de dados de imagem da Web.

## <a name="formatted-number-values"></a>Valores de número formatados

O objeto [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permite que os suplementos do Excel definam uma propriedade `numberFormat` de um valor. Depois de atribuído, esse formato de número percorre cálculos com o valor e pode ser retornado por funções.

O exemplo de código JSON a seguir mostra o esquema completo de um valor numérico formatado. O valor do número formatado `myDate` no exemplo de código é exibido como **16/1/1990** na interface do usuário do Excel. Se os requisitos mínimos de compatibilidade para o recurso de tipos de dados não forem atendidos, os cálculos usarão o `basicValue` no lugar do número formatado.

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

Comece a experimentar valores numéricos formatados abrindo Script Lab [](../overview/explore-with-script-lab.md) e verificando os tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-formatted-number.yaml) snippet de números formatados em nossa **biblioteca de Exemplos**.

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
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Os valores de entidade também oferecem uma propriedade `layouts` que cria um cartão para a entidade. O cartão é exibido como uma janela modal na interface do usuário do Excel e pode exibir informações adicionais contidas no valor da entidade, além do que é visível na célula. Para saber mais, confira [Usar cartões com tipos de dados de valor de entidade](excel-data-types-entity-card.md).

Para explorar os tipos de dados de entidade, comece [](../overview/explore-with-script-lab.md) indo para Script Lab excel e abrindo os tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) criar cartões de entidade de dados em um snippet de tabela em nossa biblioteca **de exemplos**. Os [tipos de dados: valores de entidade com](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-references.yaml) referências e tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-attribution.yaml) os snippets de propriedades de atribuição de valor de entidade oferecem uma visão mais profunda dos recursos da entidade.

### <a name="linked-entities"></a>Entidades Vinculadas

Os valores de entidade vinculados ou objetos [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) são um tipo de valor de entidade. Esses objetos integram os dados fornecidos por um serviço externo e podem exibir esses dados como um [cartão de entidade](excel-data-types-entity-card.md), como valores de entidade regulares. Os [Tipos de dados de Ações e Geografia](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) disponíveis através da interface do usuário do Excel são valores de entidade vinculados.

## <a name="web-image-values"></a>Valores de imagem da Web

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

As propriedades `basicType` e `basicValue` definem como os cálculos leem o tipo de dados de imagem da Web se os requisitos mínimos de compatibilidade para usar o recurso de tipos de dados não forem atendidos. Neste cenário, esse tipo de dados de imagem da Web é exibido como um **#VALUE!** erro na interface do usuário do Excel.

O exemplo de código JSON a seguir mostra o esquema completo de uma imagem da Web.

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Experimente tipos de dados de imagem da [Web abrindo Script Lab](../overview/explore-with-script-lab.md) e selecionando os tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-web-image.yaml) snippet de imagens da Web em nossa **biblioteca de Exemplos**.

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

Saiba mais sobre os objetos de erro de tipos de dados verificando os tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-error-values.yaml) defina o snippet de valores de erro em nossa [biblioteca Script Lab](../overview/explore-with-script-lab.md) **Exemplos**.

## <a name="next-steps"></a>Próximas etapas

Saiba como os tipos de dados de entidade estendem o potencial dos suplementos do Excel além de uma grade bidimensional com o artigo Usar cartões com tipos de dados de valor [de](excel-data-types-entity-card.md) entidade.

Use o exemplo Criar e explorar tipos de dados no [Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer) em nosso repositório [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) para experimentar mais profundamente os tipos de dados criando e fazendo sideload de um suplemento que cria e edita tipos de dados em uma pasta de trabalho.

## <a name="see-also"></a>Confira também

- [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
- [Usar cartões com tipos de dados de valor de entidade](excel-data-types-entity-card.md)
- [Criar e explorar tipos de dados no Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Funções e tipos de dados personalizados](custom-functions-data-types-concepts.md)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)