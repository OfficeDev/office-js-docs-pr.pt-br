---
title: Limites de recurso e otimização de desempenho para Suplementos do Office
description: Saiba mais sobre os limites de recursos da plataforma de Suplementos do Office, incluindo CPU e memória.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8465eb654795b538182e01d33b2fc57ddb35eaa0
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092900"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites de recurso e otimização de desempenho para Suplementos do Office

To create the best experience for your users, ensure that your Office Add-in performs within specific limits for CPU core and memory usage, reliability, and, for Outlook add-ins, the response time for evaluating regular expressions. These run-time resource usage limits apply to add-ins running in Office clients on Windows and OS X, but not on mobile apps or in a browser.

Também é possível otimizar o desempenho dos suplementos em dispositivos móveis e para área de trabalho aprimorando o uso de recursos no design e na implementação de suplementos.

## <a name="resource-usage-limits-for-add-ins"></a>Limites de uso de recursos para suplementos

Os limites de uso de recursos em tempo de execução se aplicam a todos os tipos de Suplementos do Office. Esses limites ajudam a garantir o desempenho para seus usuários e atenuar ataques de negação de serviço. Teste seu Suplemento do Office no aplicativo do Office de destino usando uma variedade de dados possíveis e meça seu desempenho em relação aos limites de uso em tempo de execução a seguir.

- **Uso de núcleo de CPU**: um limite de uso de núcleo de CPU único de 90%, observado três vezes em intervalos padrão de cinco segundos.

   O intervalo padrão para um cliente do Office verificar o uso do núcleo da CPU é a cada 5 segundos. Se o cliente do Office detectar que o uso principal da CPU de um suplemento está acima do valor limite, ele exibirá uma mensagem perguntando se o usuário deseja continuar executando o suplemento. Se o usuário optar por continuar, o cliente do Office não perguntará ao usuário novamente durante essa sessão de edição. Os administradores podem querer usar a chave de registro **AlertInterval** para elevar o limite caso os usuários executem suplementos que consomem muita CPU, a fim de reduzir a exibição desta mensagem de aviso.

- **Uso de memória**: um limite de uso de memória padrão que é determinado dinamicamente com base na memória física disponível do dispositivo.

   Por padrão, quando um cliente do Office detecta que o uso de memória física em um dispositivo excede 80% da memória disponível, o cliente começa a monitorar o uso de memória do suplemento, em um nível de documento para suplementos de conteúdo e painel de tarefas e em um nível de caixa de correio para suplementos do Outlook. Em um intervalo padrão de 5 segundos, o cliente avisa o usuário se o uso de memória física para um conjunto de suplementos no nível do documento ou da caixa de correio exceder 50%. Esse limite de uso de memória usa memória física em vez de virtual para garantir o desempenho em dispositivos com RAM limitada, como tablets. Os administradores podem substituir essa configuração dinâmica por um limite explícito usando a chave do Registro do **Windows MemoryAlertThreshold** como uma configuração global, ir ajustar o intervalo de alerta usando a **chave AlertInterval** como uma configuração global.

- **Tolerância a falhas**: um limite padrão de quatro falhas para um suplemento.

   Os administradores podem ajustar o limite para casos de falha usando a chave de registro **RestartManagerRetryLimit**.

- **Bloqueio de aplicativo**: um limite prolongado de falta de resposta de cinco segundos para um suplemento.

   Isso afeta as experiências do usuário do suplemento e do aplicativo do Office. Quando isso ocorre, o aplicativo do Office reinicia automaticamente todos os suplementos ativos para um documento ou caixa de correio (quando aplicável) e avisa o usuário sobre qual suplemento ficou sem resposta. Suplementos podem atingir esse limite quando não têm rendimento do processamento regularmente ao realizar tarefas de execução demorada. Há técnicas para garantir que não ocorra bloqueio. Os administradores não podem substituir esse limite.

### <a name="outlook-add-ins"></a>Suplementos do Outlook

If any Outlook add-in exceeds the preceding thresholds for CPU core or memory usage, or tolerance limit for crashes, Outlook disables the add-in. The Exchange Admin Center displays the disabled status of the app.

> [!NOTE]
> Mesmo que apenas clientes avançados do Outlook, e não o Outlook Online ou dispositivos móveis, monitorarem o uso de recursos, se um cliente avançado desativar um suplemento do Outlook, o suplemento também é desativado para uso no Outlook Online e dispositivos móveis.

Além das regras de núcleo de CPU, memória e confiabilidade, os suplementos do Outlook devem observar as regras a seguir na ativação.

- **Regular expressions response time** - A default threshold of 1,000 milliseconds for Outlook to evaluate all regular expressions in the manifest of an Outlook add-in. Exceeding the threshold causes Outlook to retry evaluation at a later time.

    Usando uma política de grupo ou uma configuração específica do aplicativo no Registro do Windows, os administradores podem ajustar esse valor de limite padrão de 1.000 milissegundos na configuração **OutlookActivationAlertThreshold** .

- **Reavaliação de expressões regulares**: um limite padrão de três vezes para que o Outlook reavalie todas as expressões regulares em um manifesto. Se a avaliação falhar todas as três vezes excedendo o limite aplicável (que é o padrão de 1.000 milissegundos ou um valor especificado pelo **OutlookActivationAlertThreshold**, se essa configuração existir no Registro do Windows), o Outlook desabilitará o suplemento do Outlook. O Exchange Administração Center exibe o status desabilitado e o suplemento está desabilitado para uso nos clientes avançados do Outlook e Outlook na Web dispositivos móveis.

    Usando uma política de grupo ou uma configuração específica do aplicativo no Registro do Windows, os administradores podem ajustar esse número de vezes para repetir a avaliação na configuração **OutlookActivationManagerRetryLimit** .

### <a name="excel-add-ins"></a>Suplementos do Excel

Se você estiver criando um suplemento do Excel, esteja ciente das seguintes limitações de tamanho ao interagir com a pasta de trabalho.

- O Excel na Web tem um limite de tamanho de conteúdo para solicitações e respostas de 5 MB. `RichAPI.Error` será lançado se esse limite for excedido.
- Um intervalo é limitado a cinco milhões de células para obter operações.

Se você espera que a entrada do usuário exceda esses limites, verifique os dados antes de chamar `context.sync()`. Divida a operação em partes menores, conforme necessário. Certifique-se de chamar `context.sync()` cada sub-operação para evitar que essas operações sejam agrupadas em lote novamente.

Essas limitações normalmente são excedida por intervalos grandes. Seu suplemento pode ser capaz de usar [RangeAreas](/javascript/api/excel/excel.rangeareas) para atualizar estrategicamente as células dentro de um intervalo maior. Para obter mais informações sobre como trabalhar `RangeAreas`com, [consulte Trabalhar com vários intervalos simultaneamente em suplementos do Excel](../excel/excel-add-ins-multiple-ranges.md). Para obter informações adicionais sobre como otimizar o tamanho da carga no Excel, confira as práticas recomendadas [de limite de tamanho de carga](../excel/performance.md#payload-size-limit-best-practices).

### <a name="task-pane-and-content-add-ins"></a>Suplementos do painel de tarefas e de conteúdo

Se qualquer suplemento de conteúdo ou painel de tarefas exceder os limites anteriores no uso do núcleo da CPU ou da memória ou no limite de tolerância a falhas, o aplicativo do Office correspondente exibirá um aviso para o usuário. Nesse momento, o usuário poderá executar uma destas ações:

- Reiniciar o suplemento.
- Cancel further alerts about exceeding that threshold. Ideally, the user should then delete the add-in from the document; continuing the add-in would risk further performance and stability issues.  

## <a name="verify-resource-usage-issues-in-the-telemetry-log"></a>Verificar problemas de uso de recursos no Log de Telemetria

O Office fornece um Log de Telemetria que mantém um registro de determinados eventos (carregar, abrir, fechar e erros) de soluções do Office em execução no computador local, incluindo problemas de uso de recursos em um Suplemento do Office. Se você tiver o Log de Telemetria configurado, poderá usar o Excel para abrir o Log de Telemetria no seguinte local padrão na unidade local.

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

For each event that the Telemetry Log tracks for an add-in, there is a date/time of the occurrence, event ID, severity, and short descriptive title for the event, the friendly name and unique ID of the add-in, and the application that logged the event. You can refresh the Telemetry Log to see the current tracked events. The following table shows examples of Outlook add-ins that were tracked in the Telemetry log.

|Data/Hora|ID do Evento|Severity|Título|Arquivo|ID|Application|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7 |*Não aplicável*|manifesto de suplemento baixado com êxito|Quem é quem|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|8/10/2012 17:57:01|7 |*Não aplicável*|manifesto de suplemento baixado com êxito|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

A tabela a seguir lista os eventos que o Log de Telemetria acompanha para os Suplementos do Office em geral.

|ID do Evento|Título|Severity|Descrição|
|:-----|:-----|:-----|:-----|
|7 |Manifesto de suplemento baixado com êxito|*Não aplicável*|O manifesto do Suplemento do Office foi carregado e lido com êxito pelo aplicativo do Office.|
|8 |Manifesto de suplemento não baixado|Crítico|O aplicativo do Office não pôde carregar o arquivo de manifesto para o Suplemento do Office do catálogo do SharePoint, catálogo corporativo ou AppSource.|
|9 |Não foi possível analisar a marcação do suplemento|Crítico|O aplicativo do Office carregou o manifesto do Suplemento do Office, mas não pôde ler a marcação HTML do aplicativo.|
|10|O suplemento usou CPU em excesso|Crítico|O suplemento do Office usou mais de 90% dos recursos da CPU em um período de tempo finito.|
|15|Suplemento desabilitado porque esgotou o tempo limite na pesquisa de cadeia de caracteres|*Não aplicável*|Os suplementos do Outlook pesquisam a linha de assunto e a mensagem de um e-mail para determinar se devem ser exibidas usando uma expressão regular. O suplemento do Outlook listado na coluna Arquivo  foi desabilitado pelo Outlook porque ele tempo limiteu repetidamente ao tentar corresponder a uma expressão regular.|
|18 |Suplemento fechado com êxito|*Não aplicável*|O aplicativo do Office pôde fechar o Suplemento do Office com êxito.|
|19|O suplemento encontrou um erro de tempo de execução|Crítico|O suplemento do Office teve um problema que causou sua falha. Para obter mais detalhes, examine o log **de Alertas do Microsoft Office** usando o windows Visualizador de Eventos no computador que encontrou o erro.|
|20|Falha ao verificar a licença do suplemento|Crítico|As informações de licenciamento do suplemento do Office não puderam ser verificadas e podem ter expirado. Para obter mais detalhes, examine o log **de Alertas do Microsoft Office** usando o windows Visualizador de Eventos no computador que encontrou o erro.|

Saiba mais em [Implantar o Painel de Telemetria](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) e [Solução de problemas de arquivos do Office e soluções personalizadas com o log de telemetria](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log).

## <a name="design-and-implementation-techniques"></a>Técnicas de design e implementação

Embora os limites de recursos para o uso de CPU e memória, a tolerância a falhas e a capacidade de resposta da interface do usuário se apliquem a suplementos do Office executados somente em clientes avançados, otimizar o uso desses recursos e da bateria deve ter prioridade se você quer que o suplemento tenha desempenho satisfatório em todos os dispositivos e clientes compatíveis. A otimização é particularmente importante se o suplemento efetua operações de longa duração ou lida com grandes conjuntos de dados. A lista a seguir sugere algumas técnicas para dividir operações com uso intensivo de CPU ou de dados em partes menores para que seu suplemento possa evitar o consumo excessivo de recursos e o aplicativo do Office possa permanecer responsivo.

- Em um cenário em que o suplemento precisa ler um grande volume de dados de um conjunto de dados não associado, você pode aplicar a paginação ao ler os dados de uma tabela ou reduzir o tamanho dos dados em cada operação de leitura mais curta, em vez de tentar concluir a leitura em uma única operação. Você pode fazer isso por meio do [método setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) do objeto global para limitar a duração da entrada e da saída. Também lida com os dados em blocos definidos, em vez dos dados não associados aleatoriamente. Outra opção é usar [o assíncrono](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) para lidar com suas Promessas.

- If your add-in uses a CPU-intensive algorithm to process a large volume of data, you can use web workers to perform the long-running task in the background while running a separate script in the foreground, such as displaying progress in the user interface. Web workers do not block user activities and allow the HTML page to remain responsive. For an example of web workers, see [The Basics of Web Workers](https://www.html5rocks.com/tutorials/workers/basics/). See [Web Workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API) for more information about the Web Workers API.

- Se o suplemento usa um algoritmo com uso intensivo de CPU, mas é possível dividir a entrada ou a saída de dados em conjuntos menores, considere criar um serviço Web passando os dados para o serviço Web para aliviar a carga da CPU e aguarde um retorno de chamada assíncrono.

- Teste o suplemento em relação ao maior volume de dados esperado e restrinja o suplemento a processar até esse limite.

### <a name="performance-improvements-with-the-application-specific-apis"></a>Melhorias de desempenho com as APIs específicas do aplicativo

As dicas de desempenho em Usar o modelo de [API](../develop/application-specific-api-model.md) específico do aplicativo fornecem diretrizes ao usar as APIs específicas do aplicativo para Excel, OneNote, Visio e Word. Em resumo, você deve:

- [Carregue apenas as propriedades necessárias](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended).
- [Minimize o número de chamadas sync()](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls). Leia [Evite usar o método context.sync em loops](correlated-objects-pattern.md) para obter mais informações sobre como gerenciar `sync` chamadas em seu código.
- [Minimize o número de objetos proxy criados](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created). Você também pode desviá-lo de objetos proxy, conforme descrito na próxima seção.

#### <a name="untrack-unneeded-proxy-objects"></a>Objetos de proxy desnecessários

[Os objetos proxy](../develop/application-specific-api-model.md#proxy-objects) persistem na memória até `RequestContext.sync()` que sejam chamados. Grandes operações em lote podem gerar muitos objetos de proxy que são necessários apenas uma vez pelo suplemento e podem ser liberados da memória antes da execução do lote.

O `untrack()` método libera o objeto da memória. Esse método é implementado em muitos objetos de proxy de API específicos do aplicativo. Chamar `untrack()` depois que o suplemento for concluído com o objeto deve gerar um benefício de desempenho perceptível ao usar um grande número de objetos proxy.

> [!NOTE]
> `Range.untrack()` é um atalho para [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#office-officeextension-trackedobjects-remove-member(1)). Qualquer objeto de proxy pode ser não-rastreado, removendo-o da lista de objetos rastreados no contexto.

O exemplo de código do Excel a seguir preenche um intervalo selecionado com dados, uma célula por vez. Depois que o valor é adicionado à célula, o intervalo que representa a célula é não-rastreado. Execute esse código em um intervalo selecionado de 20.000 de 10.000 células, primeiro, com a linha `cell.untrack()` e, em seguida, sem ela. Você deve observar que o código é executado mais rapidamente com a linha `cell.untrack()` do que sem ela. Você também poderá observar um tempo de resposta mais rápido posteriormente, porque a etapa de limpeza leva menos tempo.

```js
Excel.run(async (context) => {
    const largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (let i = 0; i < largeRange.rowCount; i++) {
        for (let j = 0; j < largeRange.columnCount; j++) {
            let cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

Observe que a necessidade de descompstalar objetos só se torna importante quando você está lidando com milhares deles. A maioria dos suplementos não precisará gerenciar o acompanhamento de objetos de proxy.

## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
- [Limites de ativação e da API do JavaScript para Suplementos do Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Otimização de desempenho usando a API do JavaScript para Excel](../excel/performance.md)
