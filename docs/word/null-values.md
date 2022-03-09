---
title: Valores nulos em complementos do Word
description: Saiba como trabalhar com valores nulos no seu complemento do Word.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e21677dafcaaaa7e9e9164ef18c82f49820298d6
ms.sourcegitcommit: 9d930b4c77c342246607aef30479e31fdbdd47f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63353853"
---
# <a name="null-values-in-word-add-ins"></a>Valores nulos em complementos do Word

`null` tem implicações especiais nas APIs JavaScript do Word. Ele é usado para representar valores padrão ou nenhuma formatação.

## <a name="null-property-values-in-the-response"></a>Valores da propriedade nula na resposta

As propriedades de formatação, como [cor](/javascript/api/word/word.font#word-word-font-color-member) , conterão `null` valores na resposta quando valores diferentes existirem no intervalo [especificado](/javascript/api/word/word.range). Por exemplo, se você recuperar um intervalo e carregar sua propriedade `range.font.color`:

- Se todo o texto no intervalo tiver a mesma cor de fonte, `range.font.color` especificará essa cor.
- Se houver várias cores de fonte dentro do intervalo, `range.font.color` será `null`.
