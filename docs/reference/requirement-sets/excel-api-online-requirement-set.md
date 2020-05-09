---
title: O conjunto de requisitos somente online da API JavaScript do Excel
description: Detalhes sobre o conjunto de requisitos ExcelApiOnline
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f177e0107de7172c350f94c3a022cb3e0db5c6f5
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170783"
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
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Especifica o ângulo no qual o texto é orientado para o título do eixo do gráfico. O valor deve ser um inteiro de-90 a 90 ou o inteiro 180 para texto orientado verticalmente.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtém o número de tabelas dinâmicas na coleção.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtém a primeira tabela dinâmica na coleção. As tabelas dinâmicas da coleção são classificadas de cima para baixo e da esquerda para a direita, de forma que a tabela superior esquerda seja a primeira tabela dinâmica na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtém uma Tabela Dinâmica por nome.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getpivotrs (fullyContained?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtém uma coleção com escopo de tabelas dinâmicas que se sobrepõe ao intervalo.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-online)
- [APIs de visualização do JavaScript para Excel](./excel-preview-apis.md)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)