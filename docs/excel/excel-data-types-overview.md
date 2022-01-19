---
title: Visão geral dos tipos de dados em suplementos do Excel
description: Os tipos de dados na API JavaScript do Excel permitem que os desenvolvedores de Suplementos do Office trabalhem com valores numéricos formatados, imagens da Web, valores de entidade, matrizes dentro de valores de entidade e erros aprimorados como tipos de dados.
ms.date: 12/27/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: b498d445f53441cd5db97aa71f4dee36cac45a06
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074298"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Visão geral dos tipos de dados em suplementos do Excel (versão prévia)

> [!NOTE]
> No momento, as APIs de tipos de dados só estão disponíveis na visualização pública. As APIs de visualização estão sujeitas a alterações e não se destinam ao uso em um ambiente de produção. Recomendamos que você experimente apenas em ambiente de teste e desenvolvimento. Não use APIs de visualização em um ambiente de produção ou em documentos essenciais aos negócios.
>
> Para usar APIs de visualização:
>
> - Você deve fazer referência à biblioteca **beta** na rede de distribuição de conteúdo (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). O [arquivo de definição de tipo](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) da compilação TypeScript e IntelliSense pode ser encontrado na CDN e [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Você pode instalar esses tipos com `npm install --save-dev @types/office-js-preview`. Para obter informações adicionais, confira o arquivo Leiame do pacote NPM [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js).
> - Pode ser necessário ingressar no [programa Office Insider](https://insider.office.com) para acessar builds mais recentes do Office.
>
> Para testar os tipos de dados do Office no Windows, você deve ter um número de build do Excel maior ou igual a 16.0.14626.10000. Para testar os tipos de dados no Office no Mac, você deve ter um número de build do Excel maior ou igual a 16.55.21102600.

Os tipos de dados na API JavaScript do Excel permitem que os desenvolvedores de suplementos organizem estruturas de dados complexas como objetos, como valores numéricos formatados, imagens da Web e valores de entidade.

Antes da adição de tipos de dados, a API JavaScript do Excel dá suporte a tipos de dados de cadeia de caracteres, número, booliano e erro. A camada de formatação da interface do usuário do Excel é capaz de adicionar moeda, data e outros tipos de formatação às células que contêm os quatro tipos de dados originais, mas essa camada de formatação controla apenas a exibição dos tipos de dados originais na interface do usuário do Excel. O valor do número subjacente não é alterado, mesmo quando uma célula na interface do usuário do Excel é formatada como moeda ou data. Essa lacuna entre um valor subjacente e a exibição formatada na interface do usuário do Excel pode resultar em confusão e erros durante cálculos de suplemento. Tipos de dados personalizados são uma solução para essa lacuna.

Os tipos de dados expandem o suporte à API JavaScript do Excel além dos quatro tipos de dados originais (cadeia de caracteres, número, booliano e erro) para incluir imagens da Web, valores numéricos formatados, valores de entidade, matrizes dentro de valores de entidade e tipos de dados de erro aprimorados como estruturas de dados flexíveis. Esses tipos, que potencializa muitas experiências de dados [tipos de dados vinculados](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) permitem precisão e simplicidade durante cálculos de suplementos e estendem o potencial dos suplementos do Excel além de uma grade bidimensional.

## <a name="data-types-and-custom-functions"></a>Tipos de dados e funções personalizadas

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Os tipos de dados aprimoram o poder das funções personalizadas. As funções personalizadas aceitam tipos de dados como entradas para funções personalizadas e saídas de funções personalizadas, e funções personalizadas usam o mesmo esquema JSON para tipos de dados que a API JavaScript do Excel. Esse esquema JSON de tipos de dados é mantido conforme as funções personalizadas calculam e avaliam. Para saber mais sobre a integração de tipos de dados com suas funções personalizadas, consulte [Funções personalizadas e tipos de dados](custom-functions-data-types-concepts.md).

## <a name="see-also"></a>Confira também

- [Conceitos básicos dos tipos de dados do Excel](excel-data-types-concepts.md)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Funções e tipos de dados personalizados](custom-functions-data-types-concepts.md)