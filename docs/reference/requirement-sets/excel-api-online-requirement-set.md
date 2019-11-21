---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757489"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>O conjunto de requisitos somente online da API JavaScript do Excel

O `ExcelApiOnline` conjunto de requisitos é um conjunto de requisitos especiais que inclui recursos que estão disponíveis apenas para o Excel na Web. As APIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações estruturais ou comportamentais não documentadas) para o Excel no host da Web. `ExcelApiOnline`são considerados como "Preview" APIs para outras plataformas (Windows, Mac, iOS) e podem não ser compatíveis com nenhuma dessas plataformas.

Quando há suporte para `ExcelApiOnline` APIs no conjunto de requisitos em todas as plataformas, elas serão adicionadas ao próximo conjunto de`ExcelApi 1.[NEXT]`requisitos liberados (). Depois que o novo requisito for público, essas APIs serão removidas do `ExcelApiOnline`. Pense nisso como um processo de promoção semelhante à de uma API que se move da versão prévia para o lançamento.

> [!IMPORTANT]
> `ExcelApiOnline`é o superconjunto do conjunto de requisitos mais recente.

> [!IMPORTANT]
> `ExcelApiOnline 1.1`é a única versão das APIs somente online. Isso ocorre porque o Excel na Web sempre terá uma única versão disponível para os usuários que tenham a versão mais recente.

## <a name="recommended-usage"></a>Uso recomendado

Como `ExcelApiOnline` as APIs só têm suporte no Excel na Web, seu suplemento deve verificar se o conjunto de requisitos é suportado antes de chamar essas APIs. Isso evita chamar uma API somente online em uma plataforma diferente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Depois que a API estiver em um conjunto de requisitos de plataforma cruzada, você deverá remover `isSetSupported` ou editar a verificação. Isso habilitará o recurso do seu suplemento em outras plataformas. Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.

> [!IMPORTANT]
> O manifesto não pode `ExcelApiOnline 1.1` ser especificado como um requisito de ativação. Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).

## <a name="api-list"></a>Lista de APIs

No momento, não há nenhuma API somente online. Confira novamente à medida que novos recursos são adicionados ao Excel na Web e suportados pelas APIs JavaScript do Office.

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online)
- [APIs de visualização do JavaScript para Excel](./excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)