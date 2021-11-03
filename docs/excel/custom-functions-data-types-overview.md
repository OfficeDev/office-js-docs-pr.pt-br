---
title: Visão geral de tipos de dados e funções personalizadas
description: Use os tipos de dados do Excel com suas funções personalizadas e Suplementos do Office.
ms.date: 11/01/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: ddf881cc2f92f430c8d68d346cc5f494be51c19f
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681733"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>Usar tipos de dados com funções personalizadas no Excel (visualização)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Os tipos de dados expandem a API JavaScript do Excel para dar suporte a tipos de dados além dos quatro tipos de dados originais (cadeia de caracteres, número, booleano e erro). Os tipos de dados incluem suporte para imagens da Web, valores de número formatados, valores de entidade e matrizes nos valores da entidade.

Esses tipos de dados ampliam o poder das funções personalizadas, pois as funções personalizadas aceitam tipos de dados como valores de entrada e saída. Você pode gerar tipos de dados por meio de funções personalizadas ou levar os tipos de dados existentes como argumentos de função nos cálculos. Depois que o esquema JSON de um tipo de dados for definido, esse esquema será mantido em todos os cálculos de função personalizada.

Para saber mais sobre como usar tipos de dados com um suplemento do Excel, confira [Visão geral de tipos de dados nos suplementos do Excel](/excel-data-types-overview.md). Para saber mais sobre como integrar tipos de dados personalizados com suas funções personalizadas, confira [Conceitos principais de funções e tipos de dados personalizados](/custom-functions-data-types-concepts.md).

## <a name="see-also"></a>Confira também

* [Visão geral dos tipos de dados em suplementos do Excel](/excel-data-types-overview.md)
* [Conceitos básicos dos tipos de dados do Excel](/excel-data-types-concepts.md)
* [Conceitos principais de funções e tipos de dados personalizados](/custom-functions-data-types-concepts.md)
* [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
