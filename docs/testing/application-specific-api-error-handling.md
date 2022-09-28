---
title: Tratamento de erro com as APIs JavaScript específicas do aplicativo
description: Saiba mais sobre o Excel, o Word, o PowerPoint e outra lógica de tratamento de erro da API JavaScript específica do aplicativo para contabilização de erros de runtime.
ms.date: 07/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: b6f25f5740892df4729b72ee5ad87403853f45fb
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092991"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>Tratamento de erro com as APIs JavaScript específicas do aplicativo

Ao criar um suplemento usando as [APIs JavaScript do Office](../develop/application-specific-api-model.md) específicas do aplicativo, inclua a lógica de tratamento de erros para contabilização de erros de runtime. Fazer isso é fundamental, devido à natureza assíncrona das APIs.

## <a name="best-practices"></a>Práticas recomendadas

Em nossos [exemplos](https://github.com/OfficeDev/Office-Add-in-samples) de código [e Script Lab](../overview/explore-with-script-lab.md) snippets de código, `Excel.run``PowerPoint.run`você observará que cada chamada para , ou `Word.run` `catch` é acompanhada por uma instrução para capturar quaisquer erros. Recomendamos que você use o mesmo padrão ao criar um suplemento usando as APIs específicas do aplicativo.

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

```

## <a name="api-errors"></a>Erros de API

Quando uma solicitação de API JavaScript do Office falha ao ser executada com êxito, a API retorna um objeto de erro que contém as propriedades a seguir.

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

- **message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada. A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.

- **debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.

> [!NOTE]
> Se você usar para `console.log()` imprimir mensagens de erro no console, essas mensagens só serão visíveis no servidor. Os usuários finais não veem essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo do Office. Para relatar erros ao usuário, consulte [Notificações de erro](#error-notifications).

## <a name="error-codes-and-messages"></a>Códigos de erro e mensagens

As tabelas a seguir listam os erros que apIs específicas do aplicativo podem retornar.

> [!NOTE]
> A tabela anterior lista as mensagens de erro que você pode encontrar ao usar as APIs específicas do aplicativo. Se você estiver trabalhando com a API Comum, consulte os códigos de [erro da API Comum do Office](../reference/javascript-api-for-office-error-codes.md) para saber mais sobre mensagens de erro relevantes.

|Código de erro | Mensagem de erro | Notas |
|:----------|:--------------|:------|
|`AccessDenied` |Você não pode realizar a operação solicitada.|*Nenhum.* |
|`ActivityLimitReached`|O limite de atividades foi alcançado.|*Nenhum.* |
|`ApiNotAvailable`|A API solicitada não está disponível.|*Nenhum.* |
|`ApiNotFound`|Não foi possível encontrar a API que você está tentando usar. Ele pode estar disponível em uma versão mais recente do Excel. Consulte o [artigo de conjuntos de requisitos da API JavaScript do Excel](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) para obter mais informações.|*Nenhum.* |
|`BadPassword`|A senha fornecida está incorreta.|*Nenhum.* |
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.|*Nenhum.* |
|`ContentLengthRequired`|Um `Content-length` cabeçalho HTTP está ausente.|*Nenhum.* |
|`GeneralException`|Ocorreu um erro interno ao processar a solicitação.|*Nenhum.* |
|`InsertDeleteConflict`|A tentativa de operação de exclusão ou inserção resultou em um conflito.|*Nenhum.* |
|`InvalidArgument` |O argumento é inválido, está ausente ou tem um formato incorreto.|*Nenhum.* |
|`InvalidBinding` |Esta associação de objetos não é mais válida devido às atualizações anteriores.|*Nenhum.* |
|`InvalidOperation`|A tentativa de operação é inválida no objeto.|*Nenhum.* |
|`InvalidOperationInCellEditMode`|A operação não está disponível enquanto o Excel está no modo editar célula. Saia do modo de edição usando as **teclas Enter** ou **Tab** ou selecionando outra célula e tente novamente.|*Nenhum.* |
|`InvalidReference`|Esta referência não é válida para a operação atual.|*Nenhum.* |
|`InvalidRequest`  |Não é possível processar a solicitação.|*Nenhum.* |
|`InvalidSelection`|A seleção atual é inválida para esta operação.|*Nenhum.* |
|`ItemAlreadyExists`|O recurso que está sendo criado já existe.|*Nenhum.* |
|`ItemNotFound` |O recurso solicitado não existe.|*Nenhum.* |
|`MemoryLimitReached`|O limite de memória foi atingido. Não foi possível concluir sua ação.|*Nenhum.* |
|`NotImplemented`|O recurso solicitado não foi implementado.| Isso pode significar que a API está em versão prévia ou só tem suporte em uma plataforma específica (como somente online). Confira [o aplicativo cliente do Office e a disponibilidade da plataforma para Suplementos do Office](/javascript/api/requirement-sets) para obter mais informações.|
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.|*Nenhum.* |
|`RequestPayloadSizeLimitExceeded`|O tamanho da carga da solicitação excedeu o limite. Consulte o [artigo Limites de recursos e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.| Esse erro ocorre somente em Office na Web.|
|`ResponsePayloadSizeLimitExceeded`|O tamanho da carga de resposta excedeu o limite. Consulte o [artigo Limites de recursos e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|  Esse erro ocorre somente em Office na Web.|
|`ServiceNotAvailable`|O serviço não está disponível.|*Nenhum.* |
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.|*Nenhum.* |
|`UnsupportedFeature`|A operação falhou porque a planilha de origem contém um ou mais recursos sem suporte.|*Nenhum.* |
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.|*Nenhum.* |

### <a name="excel-specific-error-codes-and-messages"></a>Mensagens e códigos de erro específicos do Excel

|Código de erro | Mensagem de erro | Notas |
|:----------|:--------------|:------|
|`EmptyChartSeries`|A operação tentada falhou porque a série de gráficos está vazia.|*Nenhum.* |
|`FilteredRangeConflict`|A operação tentada causa um conflito com um intervalo filtrado.|*Nenhum.* |
|`FormulaLengthExceedsLimit`|O código de bytes da fórmula aplicada excede o limite máximo de comprimento. Para o Office em computadores de 32 bits, o limite de comprimento do código de bytes é de 16384 caracteres. Em computadores de 64 bits, o limite de comprimento do código de bytes é de 32768 caracteres.| Esse erro ocorre no Excel na Web e na área de trabalho.|
|`InactiveWorkbook`|A operação falhou porque várias pastas de trabalho estão abertas e a pasta de trabalho que está sendo chamada por essa API perdeu o foco.|*Nenhum.* |
|`MergedRangeConflict`|Não é possível concluir a operação. Uma tabela não pode se sobrepor a outra tabela, um relatório de tabela dinâmica, resultados de consulta, células mescladas ou um mapa XML.|*Nenhum.* |
|`NonBlankCellOffSheet`|O Microsoft Excel não pode inserir novas células porque ele efetuaria push de células não vazias para fora do final da planilha. Essas células não vazias podem parecer vazias, mas têm valores em branco, alguma formatação ou uma fórmula. Exclua linhas ou colunas suficientes para abrir espaço para o que você deseja inserir e tente novamente.|*Nenhum.* |
|`OperationCellsExceedLimit`|A operação tentada afeta mais do que o limite de 33554000 células.| Se o `TableColumnCollection.add API` erro for disparado, confirme se não há dados não intencionais dentro da planilha, mas fora da tabela. Em particular, verifique se há dados nas colunas mais à direita da planilha. Remova os dados não intencionais para resolver esse erro. Uma maneira de verificar quantas células uma operação processa é executar o seguinte cálculo: `(number of table rows) x (16383 - (number of table columns))`. O número 16383 é o número máximo de colunas compatíveis com o Excel. <br><br>Esse erro ocorre somente em Excel na Web. |
|`PivotTableRangeConflict`|A operação tentada causa um conflito com um intervalo de Tabela Dinâmica.|*Nenhum.* |
|`RangeExceedsLimit`|A contagem de células no intervalo excedeu o número máximo com suporte. Consulte o [artigo Limites de recursos e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|*Nenhum.* |
|`RefreshWorkbookLinksBlocked`|A operação falhou porque o usuário não concedeu permissão para atualizar links de pasta de trabalho externa.|*Nenhum.* |
|`UnsupportedSheet`|Esse tipo de planilha não dá suporte a essa operação, pois é uma folha Macro ou Gráfico.|*Nenhum.* |

## <a name="error-notifications"></a>Notificações de erro

A maneira como você relata erros aos usuários depende do sistema de interface do usuário que você está usando. Se você estiver usando o React como o sistema de interface do usuário, use os componentes e elementos de design do Fluent UI. Escolha um controle apropriado nesta página [da interface do usuário do Fluent](https://developer.microsoft.com/fluentui#/controls/web). Recomendamos que as mensagens de erro sejam transmitidas com uma barra de mensagens, caixa de diálogo ou modal. Se o erro estiver na entrada do usuário, exiba o erro em negrito vermelho próximo ao controle de entrada. O exemplo [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) usa um elemento MessageBar e o modifica para levar em conta o menu de personalidade em um painel de tarefas do suplemento.

Se você não estiver usando o React para a interface do usuário, considere usar os componentes mais antigos da interface do usuário do Fabric implementados diretamente em HTML e JavaScript. Alguns modelos de exemplo estão no repositório [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) . Dê uma olhada especialmente nas subpastas de caixa de diálogo e navegação. O [excel-add-in-SalesLeads de exemplo](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) usa uma faixa de mensagem.

## <a name="see-also"></a>Confira também

- [Objeto OfficeExtension.Error (API JavaScript para Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Códigos de erro comuns da API do Office](../reference/javascript-api-for-office-error-codes.md)
