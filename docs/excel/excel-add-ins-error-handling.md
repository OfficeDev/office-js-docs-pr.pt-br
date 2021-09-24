---
title: Tratamento de erros com a EXCEL JavaScript
description: Saiba mais Excel a lógica de tratamento de erros da API JavaScript para levar em conta erros de tempo de execução.
ms.date: 09/20/2021
ms.localizationpriority: medium
ms.openlocfilehash: 24daaa8dcd5256be997c8742016a9ec80b3294df
ms.sourcegitcommit: 43f20d0933d0159dd390da052187b315222b185f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/24/2021
ms.locfileid: "59502728"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Tratamento de erros com a EXCEL JavaScript

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para obter mais informações sobre o método e a natureza assíncrona da API JavaScript Excel, consulte Excel modelo de objeto `sync()` [JavaScript em Office Add-ins](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Práticas recomendadas

Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a>Erros de API

Quando uma solicitação Excel API JavaScript falha ao executar com êxito, a API retorna um objeto de erro que contém as seguintes propriedades.

- **code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.

- **message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada. A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.

- **debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.

> [!NOTE]
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do complemento ou em qualquer lugar no Office aplicativo.

## <a name="error-messages"></a>Mensagens de erro

A tabela a seguir é uma lista de erros que a API pode retornar.

|Código de erro | Mensagem de erro | Observações |
|:----------|:--------------|:------|
|`AccessDenied` |Você não pode realizar a operação solicitada.| |
|`ActivityLimitReached`|O limite de atividades foi alcançado.| |
|`ApiNotAvailable`|A API solicitada não está disponível.| |
|`ApiNotFound`|A API que você está tentando usar não foi encontrada. Ele pode estar disponível em uma versão mais recente do Excel. Consulte o [Excel de requisitos da API JavaScript para](../reference/requirement-sets/excel-api-requirement-sets.md) obter mais informações.| |
|`BadPassword`|A senha fornecida está incorreta.| |
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.| |
|`ContentLengthRequired`|Um `Content-length` cabeçalho HTTP está faltando.| |
|`FilteredRangeConflict`|A operação tentada causa um conflito com um intervalo filtrado.| |
|`FormulaLengthExceedsLimit`|O bytecode da fórmula aplicada excede o limite máximo de comprimento. Para Office em máquinas de 32 bits, o limite de comprimento do bytecode é de 16384 caracteres. Em máquinas de 64 bits, o limite de comprimento do bytecode é de 32768 caracteres.| Esse erro ocorre no Excel na Web e na área de trabalho.|
|`GeneralException`|Ocorreu um erro interno ao processar a solicitação.| |
|`InactiveWorkbook`|A operação falhou porque várias guias de trabalho estão abertas e a workbook que está sendo chamada por essa API perdeu o foco.| |
|`InsertDeleteConflict`|A tentativa de operação de exclusão ou inserção resultou em um conflito.| |
|`InvalidArgument` |O argumento é inválido, está ausente ou tem um formato incorreto.| |
|`InvalidBinding` |Esta associação de objetos não é mais válida devido às atualizações anteriores.| |
|`InvalidOperation`|A tentativa de operação é inválida no objeto.| |
|`InvalidOperationInCellEditMode`|A operação não está disponível enquanto o Excel está no modo Editar célula. Saia do modo Editar usando as **teclas Enter** ou **Tab** ou selecionando outra célula e tente novamente.| |
|`InvalidReference`|Esta referência não é válida para a operação atual.| |
|`InvalidRequest`  |Não é possível processar a solicitação.| |
|`InvalidSelection`|A seleção atual é inválida para esta operação.| |
|`ItemAlreadyExists`|O recurso que está sendo criado já existe.| |
|`ItemNotFound` |O recurso solicitado não existe.| |
|`MemoryLimitReached`|O limite de memória foi atingido. Sua ação não pôde ser concluída.| |
|`MergedRangeConflict`|Não é possível concluir a operação. Uma tabela não pode se sobrepor a outra tabela, um relatório de tabela dinâmica, resultados de consulta, células mescladas ou um mapa XML.|
|`NonBlankCellOffSheet`|Microsoft Excel não pode inserir novas células porque empurraria células não vazias do final da planilha. Essas células não vazias podem aparecer vazias, mas têm valores em branco, algumas formatação ou uma fórmula. Exclua linhas ou colunas suficientes para dar espaço ao que você deseja inserir e tente novamente.| |
|`NotImplemented`|O recurso solicitado não foi implementado.| |
|`OperationCellsExceedLimit`|A operação tentada afeta mais do que o limite de 33554000 células.| Se o gatilho disparar esse erro, confirme se não há dados não intencional dentro da planilha, mas `TableColumnCollection.add API` fora da tabela. Em particular, verifique se há dados nas colunas mais à direita da planilha. Remova os dados não intencionados para resolver esse erro. Uma maneira de verificar quantas células uma operação processa é executar o seguinte cálculo: `(number of table rows) x (16383 - (number of table columns))` . O número 16383 é o número máximo de colunas que Excel suporta. <br><br>Esse erro só ocorre em Excel na Web. |
|`PivotTableRangeConflict`|A operação tentada causa um conflito com um intervalo de tabela dinâmica.| |
|`RangeExceedsLimit`|A contagem de células no intervalo excedeu o número máximo suportado. Consulte o [artigo Limites de recursos e otimização de desempenho para Office de complementos](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.| |
|`RefreshWorkbookLinksBlocked`|A operação falhou porque o usuário não concedeu permissão para atualizar os links da agenda de trabalho externa.| |
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.| |
|`RequestPayloadSizeLimitExceeded`|O tamanho da carga de solicitação excedeu o limite. Consulte o [artigo Limites de recursos e otimização de desempenho para Office de complementos](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.| Esse erro só ocorre em Excel na Web.|
|`ResponsePayloadSizeLimitExceeded`|O tamanho da carga de resposta excedeu o limite. Consulte o [artigo Limites de recursos e otimização de desempenho para Office de complementos](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|  Esse erro só ocorre em Excel na Web.|
|`ServiceNotAvailable`|O serviço não está disponível.| |
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.| |
|`UnsupportedFeature`|A operação falhou porque a planilha de origem contém um ou mais recursos sem suporte.| |
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.| |
|`UnsupportedSheet`|Esse tipo de planilha não dá suporte a essa operação, pois é uma planilha Macro ou Gráfico.| |

> [!NOTE]
> A tabela anterior lista mensagens de erro que você pode encontrar ao usar a API JavaScript Excel javascript. Se você estiver trabalhando com a API Comum em vez da Excel API JavaScript específica do aplicativo, consulte Office códigos de erro comuns da [API](../reference/javascript-api-for-office-error-codes.md) para saber mais sobre mensagens de erro relevantes.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Códigos de erro comuns da API do Office](../reference/javascript-api-for-office-error-codes.md)
