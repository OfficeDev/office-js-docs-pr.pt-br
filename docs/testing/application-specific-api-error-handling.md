---
title: Tratamento de erros com as APIs JavaScript específicas do aplicativo
description: Saiba mais sobre Excel, Word, PowerPoint e outra lógica de tratamento de erros da API JavaScript específica do aplicativo para considerar erros de runtime.
ms.date: 10/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21d8d3eef36f919f95459fd8e0b3037c1d5ae1b1
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767151"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>Tratamento de erros com as APIs JavaScript específicas do aplicativo

Ao criar um suplemento usando as [APIs JavaScript específicas do aplicativo](../develop/application-specific-api-model.md), inclua a lógica de tratamento de erros para considerar erros de runtime. Fazer isso é fundamental, devido à natureza assíncrona das APIs.

## <a name="best-practices"></a>Práticas recomendadas

Em nossos [exemplos de código](https://github.com/OfficeDev/Office-Add-in-samples) e [Script Lab](../overview/explore-with-script-lab.md) snippets, você observará que cada chamada para `Excel.run`, `PowerPoint.run`ou `Word.run` é acompanhada por uma `catch` instrução para capturar quaisquer erros. Recomendamos que você use o mesmo padrão ao criar um suplemento usando as APIs específicas do aplicativo.

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

Quando uma solicitação de API JavaScript do Office não é executada com êxito, a API retorna um objeto de erro que contém as propriedades a seguir.

- **código**: a `code` propriedade de uma mensagem de erro contém uma cadeia de caracteres que faz parte ou `{application}.ErrorCodes` de `OfficeExtension.ErrorCodes` onde *{application}* representa Excel, PowerPoint ou Word. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.

- **message**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada. A mensagem de erro não se destina ao consumo por usuários finais; você deve usar o código de erro e a lógica comercial apropriada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.

- **debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.

> [!NOTE]
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens só ficarão visíveis no servidor. Os usuários finais não veem essas mensagens de erro no painel de tarefas de suplemento ou em qualquer lugar no aplicativo do Office. Para relatar erros ao usuário, consulte [Notificações de erro](#error-notifications).

## <a name="error-codes-and-messages"></a>Códigos e mensagens de erro

As tabelas a seguir listam os erros que as APIs específicas do aplicativo podem retornar.

> [!NOTE]
> As seguintes tabelas listam mensagens de erro que você pode encontrar ao usar as APIs específicas do aplicativo. Se você estiver trabalhando com a API Comum, consulte [códigos de erro da API Comum do Office](../reference/javascript-api-for-office-error-codes.md) para saber mais sobre mensagens de erro relevantes.

|Código de erro | Mensagem de erro | Notas |
|:----------|:--------------|:------|
|`AccessDenied` |Você não pode realizar a operação solicitada.|*Nenhum.* |
|`ActivityLimitReached`|O limite de atividades foi alcançado.|*Nenhum.* |
|`ApiNotAvailable`|A API solicitada não está disponível.|*Nenhum.* |
|`ApiNotFound`|A API que você está tentando usar não pôde ser encontrada. Ele pode estar disponível em uma versão mais recente do aplicativo do Office. Consulte [Disponibilidade do aplicativo cliente do Office e da plataforma para suplementos do Office](/javascript/api/requirement-sets) para obter mais informações.|*Nenhum.* |
|`BadPassword`|A senha fornecida está incorreta.|*Nenhum.* |
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.|*Nenhum.* |
|`ContentLengthRequired`|Um `Content-length` cabeçalho HTTP está ausente.|*Nenhum.* |
|`GeneralException`|Ocorreu um erro interno ao processar a solicitação.|*Nenhum.* |
|`InsertDeleteConflict`|A tentativa de operação de exclusão ou inserção resultou em um conflito.|*Nenhum.* |
|`InvalidArgument` |O argumento é inválido, está ausente ou tem um formato incorreto.|*Nenhum.* |
|`InvalidBinding` |Esta associação de objetos não é mais válida devido às atualizações anteriores.|*Nenhum.* |
|`InvalidOperation`|A tentativa de operação é inválida no objeto.|*Nenhum.* |
|`InvalidReference`|Esta referência não é válida para a operação atual.|*Nenhum.* |
|`InvalidRequest`  |Não é possível processar a solicitação.|*Nenhum.* |
|`InvalidSelection`|A seleção atual é inválida para esta operação.|*Nenhum.* |
|`ItemAlreadyExists`|O recurso que está sendo criado já existe.|*Nenhum.* |
|`ItemNotFound` |O recurso solicitado não existe.|*Nenhum.* |
|`MemoryLimitReached`|O limite de memória foi atingido. Sua ação não pôde ser concluída.|*Nenhum.* |
|`NotImplemented`|O recurso solicitado não foi implementado.| Isso pode significar que a API está em versão prévia ou só tem suporte em uma plataforma específica (como somente online). Consulte [Disponibilidade do aplicativo cliente do Office e da plataforma para suplementos do Office](/javascript/api/requirement-sets) para obter mais informações.|
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.|*Nenhum.* |
|`RequestPayloadSizeLimitExceeded`|O tamanho da carga de solicitação excedeu o limite. Consulte o artigo [Limites de recursos e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.| Esse erro ocorre apenas em Office na Web.|
|`ResponsePayloadSizeLimitExceeded`|O tamanho da carga de resposta excedeu o limite. Consulte o artigo [Limites de recursos e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|  Esse erro ocorre apenas em Office na Web.|
|`ServiceNotAvailable`|O serviço não está disponível.|*Nenhum.* |
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.|*Nenhum.* |
|`UnsupportedFeature`|A operação falhou porque a planilha de origem contém um ou mais recursos sem suporte.|*Nenhum.* |
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.|*Nenhum.* |

### <a name="excel-specific-error-codes-and-messages"></a>Códigos de erro e mensagens específicos do Excel

|Código de erro | Mensagem de erro | Notas |
|:----------|:--------------|:------|
|`EmptyChartSeries`|A tentativa de operação falhou porque a série de gráficos está vazia.|*Nenhum.* |
|`FilteredRangeConflict`|A tentativa de operação causa um conflito com um intervalo filtrado.|*Nenhum.* |
|`FormulaLengthExceedsLimit`|O bytecode da fórmula aplicada excede o limite máximo de comprimento. Para o Office em computadores de 32 bits, o limite de comprimento do bytecode é de 16384 caracteres. Em computadores de 64 bits, o limite de comprimento do bytecode é de 32768 caracteres.| Esse erro ocorre em Excel na Web e na área de trabalho.|
|`GeneralException`|*Vários.*|As APIs de tipos de dados retornam `GeneralException` erros com mensagens de erro dinâmicas. Essas mensagens fazem referência à célula que é a origem do erro e ao problema que está causando o erro, como: "A célula A1 está faltando a propriedade `type`necessária."|
|`InactiveWorkbook`|A operação falhou porque várias pastas de trabalho estão abertas e a pasta de trabalho que está sendo chamada por essa API perdeu o foco.|*Nenhum.* |
|`InvalidOperationInCellEditMode`|A operação não está disponível enquanto o Excel estiver no modo de célula Editar. Saia do modo Editar usando as **teclas Enter** ou **Tab** ou selecionando outra célula e tente novamente.|*Nenhum.* |
|`MergedRangeConflict`|Não é possível concluir a operação. Uma tabela não pode se sobrepor a outra tabela, um relatório de Tabela Dinâmica, resultados de consulta, células mescladas ou um Mapa XML.|*Nenhum.* |
|`NonBlankCellOffSheet`|O Microsoft Excel não pode inserir novas células porque ele empurraria células não vazias para fora do final da planilha. Essas células não vazias podem parecer vazias, mas têm valores em branco, alguma formatação ou uma fórmula. Exclua linhas ou colunas suficientes para abrir espaço para o que você deseja inserir e tente novamente.|*Nenhum.* |
|`OperationCellsExceedLimit`|A tentativa de operação afeta mais do que o limite de 33554000 células.| Se o `TableColumnCollection.add API` disparar esse erro, confirme se não há dados não intencionais na planilha, mas fora da tabela. Em particular, verifique se há dados na maioria das colunas à direita da planilha. Remova os dados não intencionais para resolver esse erro. Uma maneira de verificar quantas células uma operação processa é executar o seguinte cálculo: `(number of table rows) x (16383 - (number of table columns))`. O número 16383 é o número máximo de colunas compatíveis com o Excel. <br><br>Esse erro ocorre apenas em Excel na Web. |
|`PivotTableRangeConflict`|A operação tentada causa um conflito com um intervalo de Tabela Dinâmica.|*Nenhum.* |
|`RangeExceedsLimit`|A contagem de células no intervalo excedeu o número máximo com suporte. Consulte o artigo [Limites de recursos e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|*Nenhum.* |
|`RefreshWorkbookLinksBlocked`|A operação falhou porque o usuário não concedeu permissão para atualizar links externos da pasta de trabalho.|*Nenhum.* |
|`UnsupportedSheet`|Esse tipo de planilha não dá suporte a essa operação, pois é uma planilha Macro ou Gráfico.|*Nenhum.* |

### <a name="word-specific-error-codes-and-messages"></a>Códigos de erro e mensagens específicos do Word

|Código de erro | Mensagem de erro | Notas |
|:----------|:--------------|:------|
|`SearchDialogIsOpen`|A caixa de diálogo de pesquisa está aberta.|*Nenhum.* |
|`SearchStringInvalidOrTooLong`|A cadeia de caracteres de pesquisa é inválida ou muito longa.| O máximo de cadeia de caracteres de pesquisa é de 255 caracteres. |

## <a name="error-notifications"></a>Notificações de erro

A forma como você relata erros aos usuários depende do sistema de interface do usuário que você está usando. Se você estiver usando React como o sistema de interface do usuário, use os componentes da interface do usuário fluente e os elementos de design. Escolha um controle apropriado nesta [página da interface do usuário fluente](https://developer.microsoft.com/fluentui#/controls/web). Recomendamos que as mensagens de erro sejam transmitidas com uma barra de mensagens, caixa de diálogo ou modal. Se o erro estiver na entrada do usuário, exiba o erro em vermelho em negrito próximo ao controle de entrada. O exemplo [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React) usa um elemento MessageBar e o modifica para levar em conta o menu de personalidade em um painel de tarefas de suplemento.

Se você não estiver usando React para a interface do usuário, considere usar os componentes mais antigos da interface do usuário do Fabric implementados diretamente em HTML e JavaScript. Alguns modelos de exemplo estão no repositório [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) . Dê uma olhada especialmente nas subpastas de diálogo e navegação. O exemplo [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) usa um banner de mensagem.

## <a name="see-also"></a>Confira também

- [Objeto OfficeExtension.Error](/javascript/api/office/officeextension.error)
- [Códigos de erro comuns da API do Office](../reference/javascript-api-for-office-error-codes.md)
