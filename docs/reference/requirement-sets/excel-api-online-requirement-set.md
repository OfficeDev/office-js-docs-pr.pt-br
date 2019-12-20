---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad2a3cd627552baeb449397fa917fe10e86ebbaf
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814149"
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

As seguintes APIs estão atualmente disponíveis para o Excel na Web como parte do conjunto `ExcelApiOnline 1.1` de requisitos.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[menções](/javascript/api/excel/excel.comment#mentions)|Obtém as entidades (por exemplo, pessoas) mencionadas em comentários.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtém o conteúdo de comentário avançado (por exemplo, menciona em comentários). Essa cadeia de caracteres não deve ser exibida para os usuários finais. Seu suplemento só deve usar este para analisar conteúdo de comentário avançado.|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Obtém ou define o endereço de email da entidade que é mencionada em comentário.|
||[id](/javascript/api/excel/excel.commentmention#id)|Obtém ou define a ID da entidade. Isso corresponde a uma das IDs no `CommentRichContent.richContent`.|
||[name](/javascript/api/excel/excel.commentmention#name)|Obtém ou define o nome da entidade que é mencionada em comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[menções](/javascript/api/excel/excel.commentreply#mentions)|Obtém as entidades (por exemplo, pessoas) mencionadas em comentários.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Obtém o conteúdo de comentário avançado (por exemplo, menciona em comentários). Essa cadeia de caracteres não deve ser exibida para os usuários finais. Seu suplemento só deve usar este para analisar conteúdo de comentário avançado.|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[menções](/javascript/api/excel/excel.commentrichcontent#mentions)|Uma matriz que contém todas as entidades (por exemplo, pessoas) mencionadas no comentário.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange: cadeia \| de caracteres de intervalo)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Move valores de célula, formatação e fórmulas do intervalo atual para o intervalo de destino, substituindo as informações antigas nessas células.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (valor: número)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajusta o recuo da formatação do intervalo. O valor de recuo varia de 0 a 250 e é medido em caracteres.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online)
- [APIs de visualização do JavaScript para Excel](./excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)