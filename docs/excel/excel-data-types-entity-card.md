---
title: Cartão de valor da entidade de tipos de dados da API JavaScript do Excel
description: Saiba como usar cartões de valor de entidade com tipos de dados em seu suplemento do Excel.
ms.date: 10/17/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1cb6c49e0e8cb07afb4b7c78a360be6c2391437a
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607566"
---
# <a name="use-cards-with-entity-value-data-types"></a>Usar cartões com tipos de dados de valor de entidade

Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para criar janelas modais de cartão na interface do usuário do Excel com tipos de dados de valor de entidade. Esses cartões podem exibir informações adicionais contidas em um valor de entidade, além do que já está visível em uma célula, como imagens relacionadas, informações de categoria de produto e atribuições de dados.

Um valor de entidade, ou [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), é um contêiner para tipos de dados e semelhante a um objeto em programação orientada a objeto. Este artigo mostra como usar propriedades de cartão de valor de entidade, opções de layout e funcionalidade de atribuição de dados para criar valores de entidade que são exibidos como cartões.

A captura de tela a seguir mostra um exemplo de um cartão de valor de entidade aberta, nesse caso, para o **produto Tofu** de uma lista de produtos de supermercado.

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="Uma captura de tela mostrando um tipo de dados de valor de entidade com a janela do cartão exibida.":::

## <a name="card-properties"></a>Propriedades do cartão

A propriedade de valor [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) da entidade permite que você defina informações personalizadas sobre seus tipos de dados. A `properties` chave aceita tipos de dados aninhados. Cada propriedade aninhada, ou tipo de dados, deve ter um e `type` uma configuração `basicValue` .

> [!IMPORTANT]
> Os tipos de dados aninhados `properties` são usados em combinação com os valores [de layout](#card-layout) de cartão descritos na seção do artigo subsequente. Depois de definir um tipo de dados aninhado `properties`, ele deve ser atribuído na propriedade `layouts` a ser exibida no cartão.

O snippet de código a seguir mostra o JSON de um valor de entidade com vários tipos de dados aninhados em `properties`.

> [!NOTE]
> Para experimentar esse snippet de código JSON em um exemplo completo, abra o [Script Lab](../overview/explore-with-script-lab.md) no Excel e selecione Tipos de dados: criar cartões de entidade [de](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) dados em uma tabela em nossa biblioteca **de Exemplos**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Image": {
            type: Excel.CellValueType.webImage,
            address: product.productImage || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

A captura de tela a seguir mostra um cartão de valor de entidade que usa o snippet de código anterior. A captura de tela mostra as informações **de ID** do **produto, nome** **do produto,** **imagem, quantidade** por unidade e preço **unitário** do snippet de código anterior.

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="Uma captura de tela mostrando um tipo de dados de valor de entidade com a janela de layout do cartão exibida. O cartão mostra o nome do produto, a ID do produto, a quantidade por unidade e as informações de preço unitário.":::

### <a name="property-metadata"></a>Metadados de propriedade

As propriedades da entidade têm um `propertyMetadata` campo opcional que usa o [`CellValuePropertyMetadata`](/javascript/api/excel/excel.cellvaluepropertymetadata) objeto e oferece as `excludeFrom``attribution`propriedades , e `sublabel`. O snippet de código a seguir mostra como adicionar um `sublabel` `"Unit Price"` à propriedade do snippet de código anterior. Nesse caso, o sub-rótulo identifica o tipo de moeda.

> [!NOTE]
> O `propertyMetadata` campo só está disponível em tipos de dados aninhados nas propriedades da entidade.

```TypeScript
// This code snippet is an excerpt from the `properties` field of the 
// preceding `EntityCellValue` snippet. "Unit Price" is a property of 
// an entity value.
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00",
            propertyMetadata: {
              sublabel: "USD"
            }
        },
```

A captura de tela a seguir mostra um cartão de valor de entidade que usa o snippet de código anterior, `sublabel` exibindo os metadados de propriedade de **USD** ao lado da propriedade **Preço** Unitário.

:::image type="content" source="../images/excel-data-types-entity-card-property-metadata.png" alt-text="Uma captura de tela mostrando o sub-rótulo USD ao lado do Preço Unitário.":::

## <a name="card-layout"></a>Layout do cartão

A propriedade de [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) valor da entidade cria uma para a entidade e especifica a [`card`](/javascript/api/excel/excel.entityviewlayouts) aparência desse cartão, como o título do cartão, uma imagem para o cartão e o número de seções a serem exibidas.

> [!IMPORTANT]
> Os valores aninhados `layouts` são usados em combinação com os tipos de dados de [propriedades](#card-properties) do cartão descritos na seção anterior do artigo. Um tipo de dados aninhado deve ser definido antes `properties` que possa ser atribuído `layouts` para exibição no cartão.

Dentro da `card` propriedade, use o [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) objeto para definir os componentes do cartão `title`, como , `subTitle`e `sections`.

O snippet de código JSON do valor da entidade a `card` seguir mostra um layout `mainImage` `title` com objetos e aninhados, bem como três `sections` dentro do cartão. Observe que a propriedade `title` tem `"Product Name"` um tipo de dados correspondente na seção anterior do artigo [de propriedades do](#card-properties) cartão. A `mainImage` propriedade também tem um tipo de `"Image"` dados correspondente na seção anterior. A `sections` propriedade usa uma matriz aninhada e usa o [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) objeto para definir a aparência de cada seção.

Em cada seção do cartão, você pode especificar elementos `layout`como , `title`e `properties`. A `layout` chave usa o [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) objeto e aceita o valor `"List"`. A `properties` chave aceita uma matriz de cadeias de caracteres. Observe que os valores `properties` , como `"Product ID"`, têm tipos de dados correspondentes na seção anterior do artigo de propriedades [do](#card-properties) cartão. As seções também podem ser recolhíveis e podem ser definidas com valores boolianos como recolhidos ou não recolhidos quando o cartão de entidade é aberto na interface do usuário do Excel.

> [!NOTE]
> Para experimentar esse snippet de código JSON em um exemplo completo, abra o [Script Lab](../overview/explore-with-script-lab.md) no Excel e selecione Tipos de dados: criar cartões de entidade [de](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) dados em uma tabela em nossa biblioteca **de Exemplos**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
            mainImage: { 
                property: "Image" 
            },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

A captura de tela a seguir mostra um cartão de valor de entidade que usa os snippets de código anteriores. A captura de tela mostra `mainImage` o objeto na parte superior, `title` seguido pelo objeto que usa o **Nome** do Produto e está definido como **Tofu**. A captura de tela também mostra `sections`. A **seção Quantidade e** preço é recolhível e contém **Quantidade por Unidade e** **Preço Unitário**. O **campo Informações Adicionais** é recolhível e recolhido quando o cartão é aberto.

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="Uma captura de tela mostrando um tipo de dados de valor de entidade com a janela de layout do cartão exibida. O cartão mostra o título e as seções do cartão.":::

## <a name="card-data-attribution"></a>Atribuição de dados de cartão

Os cartões de valor de entidade podem exibir uma atribuição de dados para dar crédito ao provedor das informações no cartão de entidade. A propriedade de valor [`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member) da entidade usa [`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes) o objeto, que define o `description`, `logoSourceAddress`e os `logoTargetAddress` valores.

A propriedade do provedor de dados exibe uma imagem no canto inferior esquerdo do cartão de entidade. Ele usa para `logoSourceAddress` especificar uma URL de origem para a imagem. O `logoTargetAddress` valor define o destino da URL se a imagem do logotipo estiver selecionada. O `description` valor é exibido como uma dica de ferramenta ao passar o mouse sobre o logotipo. O `description` valor também será exibido como um fallback `logoSourceAddress` de texto sem formatação se o endereço de origem da imagem não estiver definido ou se o endereço de origem da imagem estiver quebrado.

O snippet de código JSON a `provider` seguir mostra um valor de entidade que usa a propriedade para especificar uma atribuição de provedor de dados para a entidade.

> [!NOTE]
> Para experimentar esse snippet de código JSON em um exemplo completo, abra o [Script Lab](../overview/explore-with-script-lab.md) no Excel e selecione Tipos de dados [:](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-attribution.yaml) propriedades de atribuição de valor de entidade em nossa biblioteca **de Exemplos**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when selected.
    }
};
```

A captura de tela a seguir mostra um cartão de valor de entidade que usa o snippet de código anterior. A captura de tela mostra a atribuição do provedor de dados no canto inferior esquerdo. Nesse caso, o provedor de dados é a Microsoft e o logotipo da Microsoft é exibido.

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="Uma captura de tela mostrando um tipo de dados de valor de entidade com a janela de layout do cartão exibida. O cartão mostra a atribuição do provedor de dados no canto inferior esquerdo.":::

## <a name="next-steps"></a>Próximas etapas

Experimente o [exemplo Criar e explorar tipos de dados no Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer) em nosso repositório [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) . Este exemplo orienta você durante a criação e o sideload de um suplemento que cria e edita tipos de dados em uma pasta de trabalho.

## <a name="see-also"></a>Confira também

- [Visão geral dos tipos de dados em suplementos do Excel](excel-data-types-overview.md)
- [Conceitos básicos dos tipos de dados do Excel](excel-data-types-concepts.md)
- [Criar e explorar tipos de dados no Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)