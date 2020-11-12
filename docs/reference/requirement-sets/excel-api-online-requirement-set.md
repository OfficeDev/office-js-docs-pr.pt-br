---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 16c96f413424d5fc85a21419fb72cf6580c1ac18
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996526"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>O conjunto de requisitos somente online da API JavaScript do Excel

O `ExcelApiOnline` conjunto de requisitos é um conjunto de requisitos especiais que inclui recursos que estão disponíveis apenas para o Excel na Web. As APIs neste conjunto de requisitos são consideradas APIs de produção (não sujeitas a alterações estruturais ou comportamentais não documentadas) para o Excel no aplicativo Web. `ExcelApiOnline` são considerados como "Preview" APIs para outras plataformas (Windows, Mac, iOS) e podem não ser compatíveis com nenhuma dessas plataformas.

Quando há suporte para APIs no `ExcelApiOnline` conjunto de requisitos em todas as plataformas, elas serão adicionadas ao próximo conjunto de requisitos liberados ( `ExcelApi 1.[NEXT]` ). Depois que o novo requisito for público, essas APIs serão removidas do `ExcelApiOnline` . Pense nisso como um processo de promoção semelhante à de uma API que se move da versão prévia para o lançamento.

> [!IMPORTANT]
> `ExcelApiOnline` é o superconjunto do conjunto de requisitos mais recente.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` é a única versão das APIs somente online. Isso ocorre porque o Excel na Web sempre terá uma única versão disponível para os usuários que tenham a versão mais recente.

## <a name="recommended-usage"></a>Uso recomendado

Como as `ExcelApiOnline` APIs só têm suporte no Excel na Web, seu suplemento deve verificar se o conjunto de requisitos é suportado antes de chamar essas APIs. Isso evita chamar uma API somente online em uma plataforma diferente.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Depois que a API estiver em um conjunto de requisitos de plataforma cruzada, você deverá remover ou editar a `isSetSupported` verificação. Isso habilitará o recurso do seu suplemento em outras plataformas. Certifique-se de testar o recurso nessas plataformas ao fazer essa alteração.

> [!IMPORTANT]
> O manifesto não pode ser especificado `ExcelApiOnline 1.1` como um requisito de ativação. Não é um valor válido a ser usado no [elemento Set](../manifest/set.md).

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Range](/javascript/api/excel/excel.range)|[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Retorna um objeto RangeAreas que representa as áreas mescladas neste intervalo.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [APIs de visualização do JavaScript para Excel](excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
