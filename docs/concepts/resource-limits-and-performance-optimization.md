---
title: Limites de recurso e otimiza??o de desempenho para Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 1f352cfe07b114a7c2622e68a0bf41fb5878d982
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Limites de recurso e otimiza??o de desempenho para Suplementos do Office

Para criar a melhor experi?ncia para os usu?rios, verifique se o desempenho do Suplemento do Office est? dentro dos limites espec?ficos para uso de mem?ria e n?cleo de CPU, confiabilidade e, para suplementos do Outlook, tempo de resposta para avaliar express?es regulares. Esses limites de uso de recursos de tempo de execu??o aplicam-se aos suplementos em execu??o em clientes do Office para Windows e OS X, mas n?o Office Online, Outlook Web App ou OWA para Dispositivos. 

Tamb?m ? poss?vel otimizar o desempenho dos suplementos em dispositivos m?veis e para ?rea de trabalho aprimorando o uso de recursos no design e na implementa??o de suplementos.

## <a name="resource-usage-limits-for-add-ins"></a>Limites de uso de recursos para suplementos

Os limites de uso de recursos de tempo de execu??o aplicam-se a todos os tipos de Suplementos do Office. Esses limites ajudam a garantir o desempenho para os usu?rios e a reduzir ataques de nega??o de servi?o. Teste o Suplemento do Office no aplicativo de host de destino usando o intervalo de dados poss?veis e me?a o desempenho em rela??o aos seguintes limites de uso de tempo de execu??o:

- **Uso de n?cleo de CPU**: um limite de uso de n?cleo de CPU ?nico de 90%, observado tr?s vezes em intervalos padr?o de cinco segundos.
    
   O intervalo padr?o para um cliente avan?ado de host verificar o uso do n?cleo da CPU ? a cada 5 segundos. Se o cliente host detectar que o uso do n?cleo da CPU de um suplemento est? acima do valor limite, ele exibe uma mensagem perguntando se o usu?rio deseja continuar a executar o suplemento. Se o usu?rio optar por continuar, o cliente host n?o pergunta novamente durante aquela sess?o de edi??o. Os administradores podem querer usar a chave de registro **AlertInterval** para elevar o limite caso os usu?rios executem suplementos que consomem muita CPU, a fim de reduzir a exibi??o desta mensagem de aviso.
    
- **Uso de mem?ria**: um limite de uso de mem?ria padr?o que ? determinado dinamicamente com base na mem?ria f?sica dispon?vel do dispositivo.
    
   Por padr?o, quando um cliente avan?ado de host detecta que o uso da mem?ria f?sica em um dispositivo excedeu 80% da mem?ria dispon?vel, o cliente come?a a monitorar o uso de mem?ria do suplemento, no ?mbito de um documento para suplementos de conte?do e de painel de tarefas e no ?mbito de caixa de correio para suplementos do Outlook. Com um intervalo padr?o de 5 segundos, o cliente avisa o usu?rio se o uso da mem?ria f?sica exceder os 50% em um conjunto de suplementos de documento ou de caixa de correio. Esse limite de uso da mem?ria utiliza a mem?ria f?sica, e n?o a virtual, para garantir o desempenho em dispositivos com RAM limitada, como tablets. Os administradores podem sobrepor esta configura??o din?mica com um limite expl?cito usando a chave de registro do Windows **MemoryAlertThreshold** como configura??o global, ou ajustando o intervalo de alerta usando a chave **AlertInterval** como configura??o global.
    
- **Toler?ncia a falhas**: um limite padr?o de quatro falhas para um suplemento.
    
   Os administradores podem ajustar o limite para casos de falha usando a chave de registro **RestartManagerRetryLimit**.
    
- **Bloqueio de aplicativo**: um limite prolongado de falta de resposta de cinco segundos para um suplemento.
    
   Isso afeta a experi?ncia do usu?rio no suplemento e no aplicativo host. Quado isso ocorre, o aplicativo host automaticamente reinicia todos os suplementos ativos de um documento ou caixa de correio (quando for aplic?vel) e avisa o usu?rio sobre qual suplemento parou de responder. Os suplementos podem atingir este limite quando n?o produzirem regularmente velocidade de processamento ao realizar tarefas com longa execu??o. H? t?cnicas para garantir que o bloqueio n?o ocorra. Os administradores n?o podem sobrepor esse limite.
    
### <a name="outlook-add-ins"></a>Suplementos do Outlook
    
Se qualquer suplemento do Outlook exceder os limites anteriores para n?cleo da CPU, uso de mem?ria ou limite de toler?ncia a falhas, o Outlook desativa o suplemento. O Centro de Administra??o do Exchange exibe o status de desativa??o do aplicativo.

> [!NOTE]
> Mesmo que apenas clientes avan?ados do Outlook, e n?o o Outlook Web App ou o OWA para dispositivos, monitorarem o uso de recursos, se um cliente avan?ado desativar um suplemento do Outlook, o suplemento tamb?m ? desativado para uso no Outlook Web App e no OWA para dispositivos.

Al?m do n?cleo da CPU, da mem?ria e de regras de confiabilidade, os suplementos do Outlook devem estar de acordo com as seguintes regras durante a ativa??o:

- **Tempo de resposta de express?es regulares**: um limite padr?o de 1.000 milissegundos para que o Outlook avalie todas as express?es regulares no manifesto de um suplemento do Outlook. Exceder o limite faz com que o Outlook repita a avalia??o posteriormente.

    Usando uma pol?tica de grupo ou uma configura??o espec?fica para um aplicativo no registro do Windows, os administradores podem ajustar esse valor limite padr?o de 1.000 milissegundos na configura??o **OutlookActivationAlertThreshold**. Para saber mais, consulte [Substituir as configura??es de uso de recursos para desempenho de suplementos do Office](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).

- **Reavalia??o de express?es regulares**: um limite padr?o de tr?s vezes para que o Outlook reavalie todas as express?es regulares em um manifesto. Se a avalia??o falhar todas as tr?s vezes excedendo o limite aplic?vel (que ? o padr?o de 1.000 milissegundos ou um valor especificado por **OutlookActivationAlertThreshold**, se essa configura??o existir no Registro do Windows), o Outlook desabilitar? o suplemento do Outlook. O Centro de Administra??o do Exchange exibe o status desabilitado, e o suplemento ? desabilitado para uso nos clientes avan?ados do Outlook, no Outlook Web App e no OWA para Dispositivos.

    Usando uma pol?tica de grupo ou uma configura??o espec?fica para um aplicativo no registro do Windows, os administradores podem ajustar esse n?mero de tentativas de avalia??o na configura??o **OutlookActivationManagerRetryLimit**. Para saber mais, consulte [Substituir as configura??es de uso de recursos para desempenho de suplementos do Office](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).

### <a name="task-pane-and-content-add-ins"></a>Suplementos de painel de tarefas e de conte?do
    
Se qualquer suplemento de painel de tarefas ou de conte?do exceder os limites anteriores no uso de n?cleo da CPU, de mem?ria ou no limite de toler?ncia a falhas, o aplicativo host correspondente exibe um aviso ao usu?rio. Neste momento, o usu?rio pode tomar uma das seguintes a??es:

- Reiniciar o suplemento.
- Cancelar outros alertas sobre a ultrapassagem desse limite. O ideal ? que o usu?rio exclua o suplemento do documento. Continuar a usar o suplemento poderia causar ainda mais problemas de desempenho e estabilidade.  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>Verificar problemas de uso de recursos no Log de Telemetria

O Office fornece um Log de Telemetria que mant?m um registro de determinados eventos (carregar, abrir, fechar e erros) de solu??es do Office em execu??o no computador local, incluindo problemas de uso de recursos em um Suplemento do Office. Se tiver o Log de Telemetria configurado, ? poss?vel usar o Excel para abri-lo no seguinte local padr?o na unidade local:

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

Para cada evento que o Log de Telemetria acompanha para um suplemento, h? a data/hora de ocorr?ncia, a ID do evento, a severidade e o t?tulo descritivo curto do evento, o nome amig?vel e a ID exclusiva do suplemento, e o aplicativo que registrou em log o evento. Voc? pode atualizar o Log de Telemetria para ver os eventos atualmente acompanhados. A tabela a seguir mostra exemplos de suplementos do Outlook que foram acompanhados no log de Telemetria. 

|**Data/Hora**|**ID do Evento**|**Severidade**|**T?tulo**|**Arquivo**|**ID**|**Aplicativo**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08/10/2012 17:57:10|7||manifesto de suplemento baixado com ?xito|Quem ? quem|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|8/10/2012 17:57:01|7||manifesto de suplemento baixado com ?xito|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

A tabela a seguir lista os eventos que o Log de Telemetria acompanha para os Suplementos do Office em geral.

|**ID do Evento**|**T?tulo**|**Severidade**|**Descri??o**|
|:-----|:-----|:-----|:-----|
|7|Manifesto de suplemento baixado com ?xito||O manifesto do Suplemento do Office foi carregado e lido com ?xito pelo aplicativo host.|
|8|Manifesto de suplemento n?o baixado|Cr?tico|O aplicativo host n?o p?de carregar o arquivo de manifesto do suplemento do Office do cat?logo do SharePoint, do cat?logo corporativo ou do AppSource.|
|9|N?o foi poss?vel analisar a marca??o do suplemento|Cr?tico|O aplicativo host carregou o manifesto do suplemento do Office, mas n?o p?de ler a marca??o HTML do aplicativo.|
|10|O suplemento usou CPU em excesso|Cr?tico|O suplemento do Office usou mais de 90% dos recursos da CPU em um per?odo de tempo finito.|
|15|Suplemento desabilitado porque esgotou o tempo limite na pesquisa de cadeia de caracteres||Os suplementos do Outlook pesquisam a linha de assunto e a mensagem de um e-mail para determinar se devem ser exibidas usando uma express?o regular. O suplemento do Outlook listado na coluna **Arquivo** foi desabilitado pelo Outlook porque atingiu o tempo limite repetidamente ao tentar fazer a correspond?ncia de uma express?o regular.|
|18|Suplemento fechado com ?xito||O aplicativo host conseguiu fechar o suplemento do Office com ?xito.|
|19|O suplemento encontrou um erro de tempo de execu??o|Cr?tico|O suplemento do Office teve um problema que causou sua falha. Para saber mais, examine o log de **Alertas do Microsoft Office** usando o Visualizador de Eventos do Windows no computador que encontrou o erro.|
|20|Falha ao verificar a licen?a do suplemento|Cr?tico|As informa??es de licenciamento do suplemento do Office n?o puderam ser verificadas e podem ter expirado. Para saber mais, examine o log de **Alertas do Microsoft Office** usando o Visualizador de Eventos do Windows no computador que encontrou o erro.|

Saiba mais em [Implantar o Painel de Telemetria](http://msdn.microsoft.com/en-us/library/f69cde72-689d-421f-99b8-c51676c77717%28Office.15%29.aspx) e [Solu??o de problemas de arquivos do Office e solu??es personalizadas com o log de telemetria](http://msdn.microsoft.com/library/ef88e30e-7537-488e-bc72-8da29810f7aa%28Office.15%29.aspx).


## <a name="design-and-implementation-techniques"></a>T?cnicas de design e implementa??o

Embora os limites de recursos para o uso de CPU e mem?ria, a toler?ncia a falhas e a capacidade de resposta da interface do usu?rio se apliquem a suplementos do Office executados somente em clientes avan?ados, otimizar o uso desses recursos e da bateria deve ter prioridade se voc? quer que o suplemento tenha desempenho satisfat?rio em todos os dispositivos e clientes compat?veis. A otimiza??o ? particularmente importante se o suplemento efetua opera??es de longa dura??o ou lida com grandes conjuntos de dados. A lista a seguir sugere algumas t?cnicas para dividir opera??es com uso intensivo da CPU ou com muitos dados em partes menores, para que o suplemento possa evitar o consumo excessivo de recursos e o aplicativo host possa continuar a responder:

- Em um cen?rio em que o suplemento precisa ler um grande volume de dados de um conjunto de dados n?o associado, voc? pode aplicar a pagina??o ao ler os dados de uma tabela ou reduzir o tamanho dos dados em cada opera??o de leitura mais curta, em vez de tentar concluir a leitura em uma ?nica opera??o. 
    
   Para obter exemplos de c?digos JavaScript e jQuery que mostram a divis?o de uma s?rie de opera??es de entrada e sa?da em dados n?o associados (que possivelmente consumiria muitos recursos de CPU e demoraria em demasiado), consulte [Como posso passar o controle de volta (brevemente) ao navegador durante um processamento de JavaScript que consome muitos recursos?](http://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). Este exemplo usa o m?todo [setTimeout](http://msdn.microsoft.com/en-us/library/ie/ms536753%28v=vs.85%29.aspx) do objeto global para limitar a dura??o da entrada e da sa?da. Tamb?m manipula os dados em peda?os definidos, ao inv?s de dados n?o associados de forma aleat?ria.
    
- Se o suplemento usa um algoritmo com uso intensivo de CPU para processar um grande volume de dados, voc? pode usar os web workers para executar a tarefa demorada em segundo plano enquanto executa um script separado em primeiro plano, como exibir o andamento na interface do usu?rio. Os Web workers n?o bloqueiam atividades do usu?rio e permitem que a p?gina HTML continue respondendo. Para obter um exemplo de Web workers, confira [No??es b?sicas de Web workers](https://www.html5rocks.com/en/tutorials/workers/basics/). Confira [Web workers](http://msdn.microsoft.com/en-us/library/IE/hh772807%28v=vs.85%29.aspx) para saber mais sobre a API Web workers do Internet Explorer.
    
- Se o suplemento usa um algoritmo com uso intensivo de CPU, mas ? poss?vel dividir a entrada ou a sa?da de dados em conjuntos menores, considere criar um servi?o Web passando os dados para o servi?o Web para aliviar a carga da CPU e aguarde um retorno de chamada ass?ncrono.
    
- Teste o suplemento em rela??o ao maior volume de dados esperado e restrinja o suplemento a processar at? esse limite.
    

## <a name="see-also"></a>Veja tamb?m

- [Privacidade e seguran?a para Suplementos do Office](../concepts/privacy-and-security.md)
- [Limites de ativa??o e da API do JavaScript API para suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/limits-for-activation-and-javascript-api-for-outlook-add-ins)
    
