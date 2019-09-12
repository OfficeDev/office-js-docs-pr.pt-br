---
title: Limites de recurso e otimização de desempenho para Suplementos do Office
description: ''
ms.date: 09/09/2019
localization_priority: Priority
ms.openlocfilehash: 33d97d36128a32f50e0689d8ac58644f83bf604f
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838582"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites de recurso e otimização de desempenho para Suplementos do Office

Para criar a melhor experiência para os usuários, verifique se o desempenho do Suplemento do Office está dentro dos limites específicos para uso de memória e núcleo de CPU, confiabilidade e, para suplementos do Outlook, tempo de resposta para avaliar expressões regulares. Esses limites de uso de recursos de tempo de execução aplicam-se aos suplementos em execução em clientes do Office para Windows e OS X, mas não a aplicativos móveis ou a um navegador.

Também é possível otimizar o desempenho dos suplementos em dispositivos móveis e para área de trabalho aprimorando o uso de recursos no design e na implementação de suplementos.

## <a name="resource-usage-limits-for-add-ins"></a>Limites de uso de recursos para suplementos

Os limites de uso de recursos de tempo de execução aplicam-se a todos os tipos de Suplementos do Office. Esses limites ajudam a garantir o desempenho para os usuários e a reduzir ataques de negação de serviço. Teste o Suplemento do Office no aplicativo de host de destino usando o intervalo de dados possíveis e meça o desempenho em relação aos seguintes limites de uso de tempo de execução:

- **Uso de núcleo de CPU**: um limite de uso de núcleo de CPU único de 90%, observado três vezes em intervalos padrão de cinco segundos.

   O intervalo padrão para um cliente avançado de host verificar o uso do núcleo da CPU é a cada 5 segundos. Se o cliente host detectar que o uso do núcleo da CPU de um suplemento está acima do valor limite, ele exibe uma mensagem perguntando se o usuário deseja continuar a executar o suplemento. Se o usuário optar por continuar, o cliente host não pergunta novamente durante aquela sessão de edição. Os administradores podem querer usar a chave de registro **AlertInterval** para elevar o limite caso os usuários executem suplementos que consomem muita CPU, a fim de reduzir a exibição desta mensagem de aviso.

- **Uso de memória**: um limite de uso de memória padrão que é determinado dinamicamente com base na memória física disponível do dispositivo.

   Por padrão, quando um cliente avançado de host detecta que o uso da memória física em um dispositivo excedeu 80% da memória disponível, o cliente começa a monitorar o uso de memória do suplemento, no âmbito de um documento para suplementos de conteúdo e de painel de tarefas e no âmbito de caixa de correio para suplementos do Outlook. Com um intervalo padrão de 5 segundos, o cliente avisa o usuário se o uso da memória física exceder os 50% em um conjunto de suplementos de documento ou de caixa de correio. Esse limite de uso da memória utiliza a memória física, e não a virtual, para garantir o desempenho em dispositivos com RAM limitada, como tablets. Os administradores podem sobrepor esta configuração dinâmica com um limite explícito usando a chave de registro do Windows **MemoryAlertThreshold** como configuração global, ou ajustando o intervalo de alerta usando a chave **AlertInterval** como configuração global.

- **Tolerância a falhas**: um limite padrão de quatro falhas para um suplemento.

   Os administradores podem ajustar o limite para casos de falha usando a chave de registro **RestartManagerRetryLimit**.

- **Bloqueio de aplicativo**: um limite prolongado de falta de resposta de cinco segundos para um suplemento.

   Isso afeta a experiência do usuário no suplemento e no aplicativo host. Quado isso ocorre, o aplicativo host automaticamente reinicia todos os suplementos ativos de um documento ou caixa de correio (quando for aplicável) e avisa o usuário sobre qual suplemento parou de responder. Os suplementos podem atingir este limite quando não produzirem regularmente velocidade de processamento ao realizar tarefas com longa execução. Há técnicas para garantir que o bloqueio não ocorra. Os administradores não podem sobrepor esse limite.

### <a name="outlook-add-ins"></a>Suplementos do Outlook

Se qualquer suplemento do Outlook exceder os limites anteriores para núcleo da CPU, uso de memória ou limite de tolerância a falhas, o Outlook desativa o suplemento. O Centro de Administração do Exchange exibe o status de desativação do aplicativo.

> [!NOTE]
> Mesmo que apenas clientes avançados do Outlook, e não o Outlook Online ou dispositivos móveis, monitorarem o uso de recursos, se um cliente avançado desativar um suplemento do Outlook, o suplemento também é desativado para uso no Outlook Online e dispositivos móveis.

Além do núcleo da CPU, da memória e de regras de confiabilidade, os suplementos do Outlook devem estar de acordo com as seguintes regras durante a ativação:

- **Tempo de resposta de expressões regulares**: um limite padrão de 1.000 milissegundos para que o Outlook avalie todas as expressões regulares no manifesto de um suplemento do Outlook. Exceder o limite faz com que o Outlook repita a avaliação posteriormente.

    Usando uma política de grupo ou uma configuração específica do aplicativo no registro do Windows, os administradores podem ajustar esse valor limite padrão de 1.000 milissegundos na configuração **OutlookActivationAlertThreshold**.

- **Reavaliação de expressões regulares**: um limite padrão de três vezes para que o Outlook reavalie todas as expressões regulares em um manifesto. Se a avaliação falhar todas as três vezes excedendo o limite aplicável (que é o padrão de 1.000 milissegundos ou um valor especificado por **OutlookActivationAlertThreshold**, se essa configuração existir no Registro do Windows), o Outlook desabilitará o suplemento do Outlook. O Centro de Administração do Exchange exibe o status desabilitado e o suplemento é desabilitado para uso nos clientes avançados do Outlook, no Outlook Online e para dispositivos móveis.

    Usando uma política de grupo ou uma configuração específica do aplicativo no registro do Windows, os administradores podem ajustar esse número de novas tentativas de avaliação na configuração **OutlookActivationManagerRetryLimit**.

### <a name="task-pane-and-content-add-ins"></a>Suplementos do painel de tarefas e de conteúdo

Se qualquer suplemento de painel de tarefas ou de conteúdo exceder os limites anteriores no uso de núcleo da CPU, de memória ou no limite de tolerância a falhas, o aplicativo host correspondente exibe um aviso ao usuário. Neste momento, o usuário pode tomar uma das seguintes ações:

- Reiniciar o suplemento.
- Cancelar outros alertas sobre a ultrapassagem desse limite. O ideal é que o usuário exclua o suplemento do documento. Continuar a usar o suplemento poderia causar ainda mais problemas de desempenho e estabilidade.  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>Verificar problemas de uso de recursos no Log de Telemetria

O Office fornece um Log de Telemetria que mantém um registro de determinados eventos (carregar, abrir, fechar e erros) de soluções do Office em execução no computador local, incluindo problemas de uso de recursos em um Suplemento do Office. Se tiver o Log de Telemetria configurado, é possível usar o Excel para abri-lo no seguinte local padrão na unidade local:

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

Para cada evento que o Log de Telemetria acompanha para um suplemento, há a data/hora de ocorrência, a ID do evento, a severidade e o título descritivo curto do evento, o nome amigável e a ID exclusiva do suplemento, e o aplicativo que registrou em log o evento. Você pode atualizar o Log de Telemetria para ver os eventos atualmente acompanhados. A tabela a seguir mostra exemplos de suplementos do Outlook que foram acompanhados no log de Telemetria. 

|**Data/Hora**|**ID do Evento**|**Severidade**|**Título**|**Arquivo**|**ID**|**Aplicativo**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7||manifesto de suplemento baixado com êxito|Quem é quem|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|8/10/2012 17:57:01|7||manifesto de suplemento baixado com êxito|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

A tabela a seguir lista os eventos que o Log de Telemetria acompanha para os Suplementos do Office em geral.

|**ID do Evento**|**Título**|**Severidade**|**Descrição**|
|:-----|:-----|:-----|:-----|
|7|Manifesto de suplemento baixado com êxito||O manifesto do Suplemento do Office foi carregado e lido com êxito pelo aplicativo host.|
|8|Manifesto de suplemento não baixado|Crítico|O aplicativo host não pôde carregar o arquivo de manifesto do suplemento do Office do catálogo do SharePoint, do catálogo corporativo ou do AppSource.|
|9|Não foi possível analisar a marcação do suplemento|Crítico|O aplicativo host carregou o manifesto do suplemento do Office, mas não pôde ler a marcação HTML do aplicativo.|
|10|O suplemento usou CPU em excesso|Crítico|O suplemento do Office usou mais de 90% dos recursos da CPU em um período de tempo finito.|
|15|Suplemento desabilitado porque esgotou o tempo limite na pesquisa de cadeia de caracteres||Os suplementos do Outlook pesquisam a linha de assunto e a mensagem de um e-mail para determinar se devem ser exibidas usando uma expressão regular. O suplemento do Outlook listado na coluna **Arquivo** foi desabilitado pelo Outlook porque atingiu o tempo limite repetidamente ao tentar fazer a correspondência de uma expressão regular.|
|18|Suplemento fechado com êxito||O aplicativo host conseguiu fechar o suplemento do Office com êxito.|
|19|O suplemento encontrou um erro de tempo de execução|Crítico|O suplemento do Office teve um problema que causou sua falha. Para saber mais, examine o log de **Alertas do Microsoft Office** usando o Visualizador de Eventos do Windows no computador que encontrou o erro.|
|20|Falha ao verificar a licença do suplemento|Crítico|As informações de licenciamento do suplemento do Office não puderam ser verificadas e podem ter expirado. Para saber mais, examine o log de **Alertas do Microsoft Office** usando o Visualizador de Eventos do Windows no computador que encontrou o erro.|

Saiba mais em [Implantar o Painel de Telemetria](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15)) e [Solução de problemas de arquivos do Office e soluções personalizadas com o log de telemetria](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log).


## <a name="design-and-implementation-techniques"></a>Técnicas de design e implementação

Embora os limites de recursos para o uso de CPU e memória, a tolerância a falhas e a capacidade de resposta da interface do usuário se apliquem a suplementos do Office executados somente em clientes avançados, otimizar o uso desses recursos e da bateria deve ter prioridade se você quer que o suplemento tenha desempenho satisfatório em todos os dispositivos e clientes compatíveis. A otimização é particularmente importante se o suplemento efetua operações de longa duração ou lida com grandes conjuntos de dados. A lista a seguir sugere algumas técnicas para dividir operações com uso intensivo da CPU ou com muitos dados em partes menores, para que o suplemento possa evitar o consumo excessivo de recursos e o aplicativo host possa continuar a responder:

- Em um cenário em que o suplemento precisa ler um grande volume de dados de um conjunto de dados não associado, você pode aplicar a paginação ao ler os dados de uma tabela ou reduzir o tamanho dos dados em cada operação de leitura mais curta, em vez de tentar concluir a leitura em uma única operação. 

   Para obter exemplos de códigos JavaScript e jQuery que mostram a divisão de uma série de operações de entrada e saída em dados não associados (que possivelmente consumiria muitos recursos de CPU e demoraria em demasiado), consulte [Como posso passar o controle de volta (brevemente) ao navegador durante um processamento de JavaScript que consome muitos recursos?](https://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). Este exemplo usa o método [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) do objeto global para limitar a duração da entrada e da saída. Também manipula os dados em pedaços definidos, ao invés de dados não associados de forma aleatória.

- Se o suplemento usa um algoritmo com uso intensivo de CPU para processar um grande volume de dados, você pode usar os web workers para executar a tarefa demorada em segundo plano enquanto executa um script separado em primeiro plano, como exibir o progreso na interface do usuário. Os Web workers não bloqueiam atividades do usuário e permitem que a página HTML continue respondendo. Para obter um exemplo de Web workers, consulte [Noções básicas de Web workers](https://www.html5rocks.com/en/tutorials/workers/basics/). Confira [Web workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API) para saber mais sobre a API Web workers.

- Se o suplemento usa um algoritmo com uso intensivo de CPU, mas é possível dividir a entrada ou a saída de dados em conjuntos menores, considere criar um serviço Web passando os dados para o serviço Web para aliviar a carga da CPU e aguarde um retorno de chamada assíncrono.

- Teste o suplemento em relação ao maior volume de dados esperado e restrinja o suplemento a processar até esse limite.


## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
- [Limites de ativação e da API do JavaScript para Suplementos do Outlook](/outlook/add-ins/limits-for-activation-and-javascript-api-for-outlook-add-ins)
- [Otimização de desempenho usando a API do JavaScript para Excel](../excel/performance.md)
