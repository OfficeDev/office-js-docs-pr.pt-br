---
title: Conceitos básicos dos tipos de dados da API JavaScript do Excel
description: Conheça os principais conceitos para usar os tipos de dados do Excel no Suplemento do Office.
ms.date: 11/01/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: cb5a1e13ced03116d10c7d7a09f822485b41ff6a
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681741"
---
# <a name="excel-data-types-core-concepts-preview"></a>Principais conceitos dos tipos de dados do Excel (versão prévia)

> [!NOTE]
> No momento, as APIs de tipos de dados só estão disponíveis na visualização pública. As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.

> [!IMPORTANT]
> Alguns dos conceitos de tipos de dados descritos neste artigo, como `Range.valuesAsJSON` estão em desenvolvimento ativo e ainda não estão disponíveis na visualização pública. Este artigo destina-se a uma introdução conceitual. Os conceitos descritos neste artigo que ainda não estão em visualização pública serão lançados para versão prévia em breve.

Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para trabalhar com tipos de dados. Ele apresenta conceitos fundamentais para o desenvolvimento de tipos de dados.

## <a name="core-concepts"></a>Principais conceitos

Use a propriedade `Range.valuesAsJSON` para trabalhar com valores de tipo de dados. Essa propriedade é semelhante ao [Range.values](/javascript/api/excel/excel.range#values), mas `Range.values` retorna apenas os quatro tipos básicos: cadeia de caracteres, número, booliano ou valores de erro. `Range.valuesAsJSON` pode retornar informações expandidas sobre os quatro tipos básicos, e essa propriedade pode retornar tipos de dados, como valores de número formatados, entidades e imagens da Web.

### <a name="json-schema"></a>Esquema JSON

Os tipos de dados usam um esquema JSON consistente que define o [CellValueType](/javascript/api/excel/excel.cellvaluetype) dos dados e informações adicionais, como `basicValue`, `numberFormat` ou `address`. Cada `CellValueType` tem propriedades disponíveis de acordo com esse tipo. Por exemplo, o tipo `webImage` inclui as propriedades [altText](/javascript/api/excel/excel.webimagecellvalue#altText) e [atribuição](/javascript/api/excel/excel.webimagecellvalue#attribution). As seções a seguir mostram exemplos de código JSON do valor de número formatado, valor de entidade e tipos de dados de imagem da Web.

## <a name="formatted-number-values"></a>Valores de número formatados

O objeto [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permite que os suplementos do Excel definam uma propriedade `numberFormat` de um valor. Depois de atribuído, esse formato de número percorre cálculos com o valor e pode ser retornado por funções.

O exemplo de código JSON a seguir mostra um valor de número formatado. O valor do número formatado `myDate` no exemplo de código é exibido como **16/1/1990** na interface do usuário do Excel.

```json
// This is an example of the JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Valores de entidade

Um valor de entidade é um contêiner dos tipos de dados, semelhante a um objeto em programação orientada a objetos. As entidades também suportam matrizes como propriedades de um valor de entidade. O objeto [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) permite que os suplementos definam propriedades como `type`, `text` e `properties`. A propriedade `properties` permite que o valor de entidade defina e contenha tipos de dados adicionais.

O exemplo de código JSON a seguir mostra um valor de entidade que contém texto, uma imagem, uma data e um valor de texto adicional.

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## <a name="web-image-values"></a>Valores de imagem da Web

O objeto [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) cria a capacidade de armazenar uma imagem como parte de uma [entidade](#entity-values) ou como um valor independente em um intervalo. Esse objeto oferece muitas propriedades, incluindo `address`, `altText` e `relatedImagesAddress`.

O exemplo de código JSON a seguir mostra como representar uma imagem da Web.

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
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
- [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Cada um dos objetos de erro pode acessar uma enumeração por meio da propriedade `errorSubType`, e essa enumeração contém dados adicionais sobre o erro. Por exemplo, o objeto de erro `BlockedErrorCellValue` pode acessar a enumeração [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype). O enumeração `BlockedErrorCellValueSubType` oferece dados adicionais sobre o que causou o erro.

## <a name="see-also"></a>Confira também

- [Visão geral dos tipos de dados em suplementos do Excel](/excel-data-types-overview.md)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Visão geral de tipos de dados e funções personalizadas](/custom-functions-data-types-overview.md)