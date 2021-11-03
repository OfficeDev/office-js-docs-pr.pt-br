---
title: Visão geral dos tipos de dados em suplementos do Excel
description: Os tipos de dados na API JavaScript do Excel permitem que os desenvolvedores de Suplementos do Office trabalhem com valores numéricos formatados, imagens da Web, valores de entidade, matrizes dentro de valores de entidade e erros aprimorados como tipos de dados.
ms.date: 11/01/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: f5866b3ec27fc2e5869150feb45564701824afcd
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681740"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Visão geral dos tipos de dados em suplementos do Excel (versão prévia)

> [!NOTE]
> No momento, as APIs de tipos de dados só estão disponíveis na visualização pública. As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.

> [!IMPORTANT]
> Algumas das APIs de tipos de dados, como `Range.valuesAsJSON` estão em desenvolvimento ativo e ainda não estão disponíveis na visualização pública. Este artigo destina-se a uma introdução conceitual. Os conceitos descritos neste artigo que ainda não estão em visualização pública serão lançados para visualização em breve.

Os tipos de dados na API JavaScript do Excel permitem que os desenvolvedores de suplementos organizem estruturas de dados complexas como objetos, como valores numéricos formatados, imagens da Web e valores de entidade.

Antes da adição de tipos de dados, a API JavaScript do Excel dá suporte a tipos de dados de cadeia de caracteres, número, booliano e erro. A camada de formatação da interface do usuário do Excel é capaz de adicionar moeda, data e outros tipos de formatação às células que contêm os quatro tipos de dados originais, mas essa camada de formatação controla apenas a exibição dos tipos de dados originais na interface do usuário do Excel. O valor do número subjacente não é alterado, mesmo quando uma célula na interface do usuário do Excel é formatada como moeda ou data. Essa lacuna entre um valor subjacente e a exibição formatada na interface do usuário do Excel pode resultar em confusão e erros durante cálculos de suplemento. Tipos de dados personalizados são uma solução para essa lacuna.

Os tipos de dados expandem o suporte à API JavaScript do Excel além dos quatro tipos de dados originais (cadeia de caracteres, número, booliano e erro) para incluir imagens da Web, valores numéricos formatados, valores de entidade, matrizes dentro de valores de entidade e tipos de dados de erro aprimorados como estruturas de dados flexíveis. Esses tipos, que potencializa muitas experiências de dados [tipos de dados vinculados](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) permitem precisão e simplicidade durante cálculos de suplementos e estendem o potencial dos suplementos do Excel além de uma grade bidimensional.

## <a name="data-types-and-custom-functions"></a>Tipos de dados e funções personalizadas

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Os tipos de dados aprimoram o poder das funções personalizadas. As funções personalizadas aceitam tipos de dados como entradas para funções personalizadas e saídas de funções personalizadas, e funções personalizadas usam o mesmo esquema JSON para tipos de dados que a API JavaScript do Excel. Esse esquema JSON de tipos de dados é mantido conforme as funções personalizadas calculam e avaliam. Para saber mais sobre a integração de tipos de dados com suas funções personalizadas, consulte [ Principais conceitos de funções e tipos de dados](/custom-functions-data-types-concepts.md).

## <a name="see-also"></a>Confira também

* [Conceitos básicos dos tipos de dados do Excel](/excel-data-types-concepts.md)
* [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
* [Visão geral de tipos de dados e funções personalizadas](/custom-functions-data-types-overview.md)