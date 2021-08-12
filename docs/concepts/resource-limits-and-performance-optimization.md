---
title: Limites de recurso e otimização de desempenho para Suplementos do Office
description: Saiba mais sobre os limites de recursos da plataforma de Office de complemento, incluindo CPU e memória.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 43902dcf3a7703a763e1268d5b5695c48c59e0fcacf3ae7d2740b54e31e3057e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082763"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites de recurso e otimização de desempenho para Suplementos do Office

Para criar a melhor experiência para os usuários, verifique se o desempenho do Suplemento do Office está dentro dos limites específicos para uso de memória e núcleo de CPU, confiabilidade e, para suplementos do Outlook, tempo de resposta para avaliar expressões regulares. Esses limites de uso de recursos de tempo de execução aplicam-se aos suplementos em execução em clientes do Office para Windows e OS X, mas não a aplicativos móveis ou a um navegador.

Também é possível otimizar o desempenho dos suplementos em dispositivos móveis e para área de trabalho aprimorando o uso de recursos no design e na implementação de suplementos.

## <a name="resource-usage-limits-for-add-ins"></a>Limites de uso de recursos para suplementos

Os limites de uso de recursos em tempo de executar se aplicam a todos os tipos de Office Desem. Esses limites ajudam a garantir o desempenho dos usuários e reduzir os ataques de negação de serviço. Certifique-se de testar seu Office add-in em seu aplicativo de Office de destino usando um intervalo de dados possíveis e mede seu desempenho em relação aos seguintes limites de uso em tempo de execução.

- **Uso de núcleo de CPU**: um limite de uso de núcleo de CPU único de 90%, observado três vezes em intervalos padrão de cinco segundos.

   O intervalo padrão para um cliente Office verificar o uso principal da CPU é a cada 5 segundos. Se o cliente Office detectar o uso principal da CPU de um complemento acima do valor limite, ele exibirá uma mensagem perguntando se o usuário deseja continuar executando o complemento. Se o usuário optar por continuar, o cliente Office não perguntará ao usuário novamente durante a sessão de edição. Os administradores podem querer usar a chave de registro **AlertInterval** para elevar o limite caso os usuários executem suplementos que consomem muita CPU, a fim de reduzir a exibição desta mensagem de aviso.

- **Uso de memória**: um limite de uso de memória padrão que é determinado dinamicamente com base na memória física disponível do dispositivo.

   Por padrão, quando um cliente Office detecta que o uso de memória física em um dispositivo excede 80% da memória disponível, o cliente começa a monitorar o uso de memória do complemento, em um nível de documento para conteúdo e complementos do painel de tarefas e em um nível de caixa de correio para Outlook de complementos. Em um intervalo padrão de 5 segundos, o cliente avisa o usuário se o uso de memória física para um conjunto de complementos no nível de documento ou caixa de correio exceder 50%. Esse limite de uso de memória usa memória física e não virtual para garantir o desempenho em dispositivos com RAM limitada, como tablets. Os administradores podem substituir essa configuração dinâmica por um limite explícito usando a chave de registro **MemoryAlertThreshold** Windows como uma configuração global, ir ajustar o intervalo de alerta usando a chave **AlertInterval** como uma configuração global.

- **Tolerância a falhas**: um limite padrão de quatro falhas para um suplemento.

   Os administradores podem ajustar o limite para casos de falha usando a chave de registro **RestartManagerRetryLimit**.

- **Bloqueio de aplicativo**: um limite prolongado de falta de resposta de cinco segundos para um suplemento.

   Isso afeta as experiências do usuário do aplicativo de Office. Quando isso ocorre, o aplicativo Office reinicia automaticamente todos os complementos ativos para um documento ou caixa de correio (quando aplicável) e avisa o usuário sobre qual add-in se tornou não responsivo. Suplementos podem atingir esse limite quando não têm rendimento do processamento regularmente ao realizar tarefas de execução demorada. Há técnicas para garantir que não ocorra bloqueio. Os administradores não podem substituir esse limite.

### <a name="outlook-add-ins"></a>Suplementos do Outlook

Se qualquer suplemento do Outlook exceder os limites anteriores para núcleo da CPU, uso de memória ou limite de tolerância a falhas, o Outlook desativa o suplemento. O Centro de Administração do Exchange exibe o status de desativação do aplicativo.

> [!NOTE]
> Mesmo que apenas clientes avançados do Outlook, e não o Outlook Online ou dispositivos móveis, monitorarem o uso de recursos, se um cliente avançado desativar um suplemento do Outlook, o suplemento também é desativado para uso no Outlook Online e dispositivos móveis.

Além das regras de núcleo da CPU, memória e confiabilidade, os Outlook deverão observar as seguintes regras na ativação.

- **Tempo de resposta de expressões regulares**: um limite padrão de 1.000 milissegundos para que o Outlook avalie todas as expressões regulares no manifesto de um suplemento do Outlook. Exceder o limite faz com que o Outlook repita a avaliação posteriormente.

    Usando uma política de grupo ou configuração específica de aplicativo no registro Windows, os administradores podem ajustar esse valor limite padrão de 1.000 milissegundos na configuração **OutlookActivationAlertThreshold.**

- **Reavaliação de expressões regulares**: um limite padrão de três vezes para que o Outlook reavalie todas as expressões regulares em um manifesto. Se a avaliação falhar todas as três vezes excedendo o limite aplicável (que é o padrão de 1.000 milissegundos ou um valor especificado pelo **OutlookActivationAlertThreshold**, se essa configuração existir no registro Windows), o Outlook desabilitará o Outlook add-in. O Exchange Admin Center exibe o status desabilitado, e o complemento está desabilitado para uso nos clientes Outlook clientes Outlook na Web e dispositivos móveis.

    Usando uma política de grupo ou configuração específica do aplicativo no registro Windows, os administradores podem ajustar esse número de vezes para repetir a avaliação na **configuração OutlookActivationManagerRetryLimit.**

### <a name="excel-add-ins"></a>Suplementos do Excel

Se você estiver criando um Excel, esteja ciente das seguintes limitações de tamanho ao interagir com a workbook.

- O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB. `RichAPI.Error` será lançado se esse limite for excedido.
- Um intervalo é limitado a cinco milhões de células para obter operações.

Se você espera que a entrada do usuário exceda esses limites, verifique os dados antes de chamar `context.sync()` . Divida a operação em partes menores conforme necessário. Certifique-se de chamar cada sub-operação para evitar que essas operações `context.sync()` sejam reunidas em lote novamente.

Essas limitações geralmente são excedida por intervalos grandes. Seu complemento pode ser capaz de usar [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar estrategicamente células dentro de um intervalo maior. Consulte [Trabalhar com vários intervalos simultaneamente em Excel de complementos](../excel/excel-add-ins-multiple-ranges.md) para obter mais informações.

### <a name="task-pane-and-content-add-ins"></a>Suplementos do painel de tarefas e de conteúdo

Se qualquer conteúdo ou um complemento do painel de tarefas exceder os limites anteriores no uso do núcleo da CPU ou da memória ou no limite de tolerância para falhas, o aplicativo Office correspondente exibirá um aviso para o usuário. Nesse momento, o usuário poderá executar uma destas ações:

- Reiniciar o suplemento.
- Cancelar outros alertas sobre a ultrapassagem desse limite. O ideal é que o usuário exclua o suplemento do documento. Continuar a usar o suplemento poderia causar ainda mais problemas de desempenho e estabilidade.  

## <a name="verify-resource-usage-issues-in-the-telemetry-log"></a>Verificar problemas de uso de recursos no Log de Telemetria

O Office fornece um Log de Telemetria que mantém um registro de determinados eventos (carregar, abrir, fechar e erros) de soluções do Office em execução no computador local, incluindo problemas de uso de recursos em um Suplemento do Office. Se você tiver o Log de Telemetria definido, poderá usar o Excel para abrir o Log de Telemetria no local padrão a seguir em sua unidade local.

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

Para cada evento que o Log de Telemetria acompanha para um suplemento, há a data/hora de ocorrência, a ID do evento, a severidade e o título descritivo curto do evento, o nome amigável e a ID exclusiva do suplemento, e o aplicativo que registrou em log o evento. Você pode atualizar o Log de Telemetria para ver os eventos atualmente acompanhados. A tabela a seguir mostra exemplos de suplementos do Outlook que foram acompanhados no log de Telemetria.

|**Data/Hora**|**ID do Evento**|**Severidade**|**Título**|**Arquivo**|**ID**|**Aplicativo**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7 ||manifesto de suplemento baixado com êxito|Quem é quem|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|8/10/2012 17:57:01|7 ||manifesto de suplemento baixado com êxito|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

A tabela a seguir lista os eventos que o Log de Telemetria acompanha para os Suplementos do Office em geral.

|**ID do Evento**|**Título**|**Severidade**|**Descrição**|
|:-----|:-----|:-----|:-----|
|7 |Manifesto de suplemento baixado com êxito||O manifesto do Office de Office foi carregado e lido com êxito pelo aplicativo Office.|
|8 |Manifesto de suplemento não baixado|Crítico|O Office aplicativo não pôde carregar o arquivo de manifesto do Office do SharePoint, catálogo corporativo ou AppSource.|
|9 |Não foi possível analisar a marcação do suplemento|Crítico|O Office o aplicativo carregou o manifesto de Office de complemento, mas não conseguiu ler a marcação HTML do aplicativo.|
|10 |O suplemento usou CPU em excesso|Crítico|O suplemento do Office usou mais de 90% dos recursos da CPU em um período de tempo finito.|
|15|Suplemento desabilitado porque esgotou o tempo limite na pesquisa de cadeia de caracteres||Os suplementos do Outlook pesquisam a linha de assunto e a mensagem de um e-mail para determinar se devem ser exibidas usando uma expressão regular. O Outlook de dados listado na  coluna Arquivo foi desabilitado por Outlook porque ele temporizou repetidamente ao tentar corresponder a uma expressão regular.|
|18 |Suplemento fechado com êxito||O Office aplicativo foi capaz de fechar o Office Add-in com êxito.|
|19|O suplemento encontrou um erro de tempo de execução|Crítico|O suplemento do Office teve um problema que causou sua falha. Para obter mais detalhes, consulte o log **Microsoft Office Alertas** usando o visualizador de eventos Windows no computador que encontrou o erro.|
|20|Falha ao verificar a licença do suplemento|Crítico|As informações de licenciamento do suplemento do Office não puderam ser verificadas e podem ter expirado. Para obter mais detalhes, consulte o log **Microsoft Office Alertas** usando o visualizador de eventos Windows no computador que encontrou o erro.|

Saiba mais em [Implantar o Painel de Telemetria](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) e [Solução de problemas de arquivos do Office e soluções personalizadas com o log de telemetria](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log).

## <a name="design-and-implementation-techniques"></a>Técnicas de design e implementação

Embora os limites de recursos para o uso de CPU e memória, a tolerância a falhas e a capacidade de resposta da interface do usuário se apliquem a suplementos do Office executados somente em clientes avançados, otimizar o uso desses recursos e da bateria deve ter prioridade se você quer que o suplemento tenha desempenho satisfatório em todos os dispositivos e clientes compatíveis. A otimização é particularmente importante se o suplemento efetua operações de longa duração ou lida com grandes conjuntos de dados. A lista a seguir sugere algumas técnicas para separar operações intensas de CPU ou de uso intenso de dados em partes menores para que o seu complemento possa evitar o consumo excessivo de recursos e o aplicativo Office pode permanecer responsivo.

- Em um cenário em que o suplemento precisa ler um grande volume de dados de um conjunto de dados não associado, você pode aplicar a paginação ao ler os dados de uma tabela ou reduzir o tamanho dos dados em cada operação de leitura mais curta, em vez de tentar concluir a leitura em uma única operação. Você pode fazer isso por meio do [método setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) do objeto global para limitar a duração da entrada e da saída. Também lida com os dados em blocos definidos, em vez dos dados não associados aleatoriamente. Outra opção é usar [o assíncrono](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) para lidar com suas Promessas.

- Se o suplemento usa um algoritmo com uso intensivo de CPU para processar um grande volume de dados, você pode usar os web workers para executar a tarefa demorada em segundo plano enquanto executa um script separado em primeiro plano, como exibir o progreso na interface do usuário. Os Web workers não bloqueiam atividades do usuário e permitem que a página HTML continue respondendo. Para obter um exemplo de Web workers, consulte [Noções básicas de Web workers](https://www.html5rocks.com/tutorials/workers/basics/). Confira [Web workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API) para saber mais sobre a API Web workers.

- Se o suplemento usa um algoritmo com uso intensivo de CPU, mas é possível dividir a entrada ou a saída de dados em conjuntos menores, considere criar um serviço Web passando os dados para o serviço Web para aliviar a carga da CPU e aguarde um retorno de chamada assíncrono.

- Teste o suplemento em relação ao maior volume de dados esperado e restrinja o suplemento a processar até esse limite.

### <a name="performance-improvements-with-the-application-specific-apis"></a>Melhorias de desempenho com as APIs específicas do aplicativo

As dicas de desempenho em Usar o modelo de [API](../develop/application-specific-api-model.md) específico do aplicativo fornecem orientações ao usar as APIs específicas do aplicativo para Excel, OneNote, Visio e Word. Em resumo, você deve:

- [Carregar somente as propriedades necessárias](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended).
- [Minimize o número de chamadas sync()](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls). Leia [Evite usar o método context.sync em loops](correlated-objects-pattern.md) para obter mais informações sobre como gerenciar chamadas em seu `sync` código.
- [Minimize o número de objetos proxy criados](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created). Você também pode desconscrevir objetos proxy, conforme descrito na próxima seção.

#### <a name="untrack-unneeded-proxy-objects"></a>Objetos proxy não desanenhados

[Os objetos proxy](../develop/application-specific-api-model.md#proxy-objects) persistem na memória até `RequestContext.sync()` que seja chamado. Grandes operações em lote podem gerar muitos objetos de proxy que são necessários apenas uma vez pelo suplemento e podem ser liberados da memória antes da execução do lote.

O `untrack()` método libera o objeto da memória. Esse método é implementado em muitos objetos proxy de API específicos do aplicativo. Chamar depois que o seu complemento for feito com o objeto deve gerar um benefício de desempenho perceptível ao usar um grande número de objetos `untrack()` proxy.

> [!NOTE]
> `Range.untrack()` é um atalho para [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove_object_). Qualquer objeto de proxy pode ser não-rastreado, removendo-o da lista de objetos rastreados no contexto.

O exemplo Excel código a seguir preenche um intervalo selecionado com dados, uma célula por vez. Depois que o valor é adicionado à célula, o intervalo que representa a célula é não-rastreado. Execute esse código em um intervalo selecionado de 20.000 de 10.000 células, primeiro, com a linha `cell.untrack()` e, em seguida, sem ela. Você deve observar que o código é executado mais rapidamente com a linha `cell.untrack()` do que sem ela. Você também poderá observar um tempo de resposta mais rápido posteriormente, porque a etapa de limpeza leva menos tempo.

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

Observe que a necessidade de desemanar objetos só se torna importante quando você está lidando com milhares deles. A maioria dos complementos não precisará gerenciar o controle de objeto proxy.

## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
- [Limites de ativação e da API do JavaScript para Suplementos do Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Otimização de desempenho usando a API do JavaScript para Excel](../excel/performance.md)
