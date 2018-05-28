---
title: Privacidade e seguran?a para suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 326c8095b6ced105cc21492dc290a443212b3d3f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="privacy-and-security-for-office-add-ins"></a>Privacidade e seguran?a para suplementos do Office

## <a name="understanding-the-add-in-runtime"></a>No??es b?sicas sobre o tempo de execu??o do suplemento

Os suplementos do Office s?o protegidos por um ambiente de tempo de execu??o de suplemento, um modelo de permiss?es com v?rias camadas e administradores de desempenho. Essa estrutura protege a experi?ncia do usu?rio das seguintes maneiras: 

- O acesso ao quadro da interface do usu?rio do aplicativo host ? gerenciado.

- ? permitido somente o acesso indireto ao thread da interface do usu?rio do aplicativo host.

- As intera??es modais n?o s?o permitidas. Por exemplo, chamadas ?s fun??es **alert**, **confirm** e **prompt** do JavaScript n?o s?o permitidas porque s?o modais.

Al?m disso, a estrutura de tempo de execu??o fornece os seguintes benef?cios para garantir que um suplemento do Office n?o possa danificar o ambiente do usu?rio:

- Isola o processo no qual o suplemento ? executado.

- N?o exige substitui??o de .dll ou de .exe ou de componentes ActiveX.

- Facilita a instala??o e a desinstala??o do suplemento.

E o uso de mem?ria, CPU e recursos de rede por suplementos do Office ? govern?vel para garantir que o bom desempenho e a confiabilidade sejam mantidos. 

As se??es a seguir descrevem brevemente como a arquitetura de tempo de execu??o d? suporte a suplementos em execu??o em clientes do Office em dispositivos Windows, em dispositivos Mac OS X e em clientes do Office Online na Web.

> **OBSERVA??O:** para saber mais sobre como usar WIP e Intune com os Suplementos do Office, confira [Usar o WIP e o Intune para proteger dados empresariais em documentos executando suplementos do Office](https://docs.microsoft.com/en-us/microsoft-365-enterprise/office-add-ins-wip).

### <a name="clients-for-windows-and-os-x-devices"></a>Clientes para dispositivos Windows e OS X

Em clientes com suporte para dispositivos de ?rea de trabalho e de tablet, como Excel, Outlook e Outlook para Mac, h? suporte a suplementos do Office por meio da integra??o de um componente no processo, o tempo de execu??o de Suplementos do Office, que gerencia o ciclo de vida do suplemento e habilita a interoperabilidade entre o suplemento e o aplicativo cliente. A p?gina da Web do suplemento em si ? hospedada fora do processo. Como mostrado na Figura 1, em um dispositivo Windows para ?rea de trabalho ou tablet, a p?gina da Web do suplemento ? hospedada em um controle do Internet Explorer que, por sua vez, ? hospedado em um processo de tempo de execu??o de suplemento que fornece seguran?a e isolamento de desempenho.

No Windows Desktop, o Modo Protegido no Internet Explorer deve ser ativado para a Zona de Site Restrito. Ele geralmente est? habilitado por padr?o. Se estiver desabilitado, um [erro ocorrer?](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) quando voc? tentar iniciar um suplemento.

*Figura 1. Ambiente de execu??o dos Suplementos do Office nos clientes Windows para ?rea de trabalho e tablet*

![Infraestrutura de cliente avan?ado](../images/dk2-agave-overview-02.png)

Como mostrado na figura a seguir, em um computador de mesa Mac OS X, a p?gina da Web do suplemento ? hospedada em um processo de host de tempo de execu??o WebKit em ?rea restrita que ajuda a fornecer um n?vel semelhante de seguran?a e prote??o de desempenho. 

*Figura 2. Ambiente de execu??o dos Suplementos do Office nos clientes Mac OS X*

![Aplicativos para o ambiente de execu??o do Office no Mac OS X](../images/dk2-agave-overview-mac-02.png)

O tempo de execu??o de Suplementos do Office gerencia a comunica??o entre processos, a convers?o de eventos e chamadas ? API JavaScript em itens nativos, bem como o suporte de comunica??o remota da interface do usu?rio para habilitar o suplemento a ser processado dentro do documento, em um painel de tarefas ou de forma adjacente a uma mensagem de e-mail, solicita??o de reuni?o ou compromisso.

### <a name="web-clients"></a>Clientes Web

Em clientes Web com suporte, como o Excel Online e o Outlook Web App, os Suplementos do Office s?o hospedados em um **iframe** que ? executado usando o atributo **sandbox** do HTML5. N?o s?o permitidos componentes ActiveX nem a navega??o na p?gina principal do cliente Web. O suporte a Suplementos do Office ? habilitado em clientes Web por meio da integra??o da API JavaScript para Office. De maneira semelhante aos aplicativos cliente de ?rea de trabalho, a API JavaScript gerencia o ciclo de vida do suplemento e a interoperabilidade entre o suplemento e o cliente Web. Essa interoperabilidade ? implementada por meio de uma infraestrutura especial de comunica??o de mensagens de publica??o entre quadros. A mesma biblioteca JavaScript (Office.js) que ? usada em clientes de ?rea de trabalho, est? dispon?vel para interagir com o cliente Web. A figura a seguir ilustra a infraestrutura que d? suporte aos Suplementos do Office no Office Online (em execu??o no navegador) e os componentes relevantes (o cliente Web, o **iframe**, o tempo de execu??o de Suplementos do Office e a API JavaScript para o Office) que s?o necess?rios para dar suporte a eles.

*Figura 3. Infraestrutura que d? suporte aos Suplementos do Office nos clientes Web do Office*

![Infraestrutura do cliente Web](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a>Integridade do suplemento no AppSource

Voc? pode disponibilizar os Suplementos do Office para o p?blico publicando-os no AppSource, que imp?e as seguintes medidas para manter a integridade dos suplementos:


- Requer que o servidor host de um Suplemento do Office sempre use o protocolo SSL para se comunicar.

- Requer que um desenvolvedor forne?a uma prova de identidade, um acordo contratual e uma pol?tica de privacidade compat?vel para enviar suplementos.

- Garante que a origem dos suplementos seja acess?vel no modo somente leitura.

- D? suporte a um sistema de revis?o pelo usu?rio para os suplementos dispon?veis para promover uma comunidade autovigilante.

## <a name="addressing-end-users-privacy-concerns"></a>Lidar com as preocupa??es de privacidade dos usu?rios finais

Esta se??o descreve a prote??o oferecida pela plataforma de Suplementos do Office da perspectiva do cliente (usu?rio final) e fornece as diretrizes sobre como dar suporte ?s expectativas dos usu?rios e como manipular com seguran?a as PII (informa??es de identifica??o pessoal) dos usu?rios.

### <a name="end-users-perspective"></a>Perspectiva dos usu?rios finais

Os Suplementos do Office s?o criados usando tecnologias da Web que s?o executadas em um controle de navegador ou em um **iframe**. Por isso, o uso de suplementos ? semelhante ? navega??o em sites na Internet ou na intranet. Os suplementos podem ser externos ? organiza??o (se voc? adquire o suplemento do AppSource) ou internos (se voc? adquire o suplemento de um cat?logo de suplementos do Exchange Server, de um cat?logo de suplementos do SharePoint ou de um compartilhamento de arquivos na rede da organiza??o). Os suplementos t?m acesso limitado ? rede, e a maioria dos suplementos pode ler ou gravar no documento ou item de email ativo. A plataforma de suplementos aplica certas restri??es antes que um usu?rio ou administrador instale ou inicie um suplemento. Por?m, como ocorre com qualquer modelo de extensibilidade, os usu?rios devem ser cuidadosos antes de iniciar um suplemento desconhecido.

A plataforma de suplementos lida com as preocupa??es com privacidade dos usu?rios finais das seguintes maneiras:

- Os dados comunicados com o servidor Web que hospeda um suplemento de conte?do, do Outlook ou de painel de tarefas, bem como a comunica??o entre o suplemento e quaisquer servi?os Web que ele usa, devem ser criptografados usando o protocolo SSL.

- Antes de instalar um suplemento do AppSource, o usu?rio pode exibir a pol?tica de privacidade e os requisitos desse suplemento. Al?m disso, os suplementos do Outlook que interagem com caixas de correio dos usu?rios exp?em as permiss?es espec?ficas das quais precisam. O usu?rio pode examinar os termos de uso, as permiss?es solicitadas e a pol?tica de privacidade antes de instalar um suplemento do Outlook.

- Ao compartilhar um documento, os usu?rios tamb?m compartilham suplementos que foram inseridos no documento ou associados a ele. Se um usu?rio abrir um documento que contenha um suplemento que o usu?rio n?o usou antes, o aplicativo host solicitar? que o usu?rio conceda permiss?o para que o suplemento seja executado no documento. Em um ambiente empresarial, o aplicativo host do Office tamb?m consultar? o usu?rio se o documento for proveniente de uma fonte externa.

- Os usu?rios podem habilitar ou desabilitar o acesso ao AppSource. Para os suplementos do conte?do e do painel de tarefas, os usu?rios gerenciam o acesso aos suplementos e cat?logos confi?veis na **Central de Confiabilidade** no cliente host do Office (aberto com **Arquivo** > **Op??es** > **Central de Confiabilidade** > **Configura??es da Central de Confiabilidade** > **Cat?logos de Suplementos Confi?veis**). Para os suplementos do Outlook, os usu?rios podem gerenciar os suplementos escolhendo o bot?o **Gerenciar Suplementos**: no Outlook para Windows, escolha **Arquivo** > **Gerenciar Suplementos**. No Outlook para Mac, escolha o bot?o **Gerenciar Suplementos** na barra de suplementos. No Outlook Web App, escolha o menu **Configura??es** (?cone de engrenagem) > **Gerenciar suplementos**. Os administradores tamb?m podem gerenciar esse acesso [usando a pol?tica de grupo](http://technet.microsoft.com/en-us/library/jj219429.aspx#BKMK_Managing).

- O design da plataforma do suplemento fornece seguran?a e desempenho aos usu?rios finais das seguintes maneiras:

  - Um Suplemento do Office ? executado em um controle de navegador da Web hospedado em um ambiente de tempo de execu??o de suplementos separado do aplicativo host do Office. Esse design fornece seguran?a e isolamento de desempenho do aplicativo host.

  - A execu??o em um controle de navegador da Web permite que o suplemento fa?a quase tudo que uma p?gina da Web regular em execu??o em um navegador pode fazer, mas, ao mesmo tempo, restringe o suplemento a observar a pol?tica de mesma origem para o isolamento de dom?nio e as zonas seguran?a.

Os suplementos do Outlook fornecem recursos adicionais de seguran?a e desempenho por meio do monitoramento de uso de recursos espec?ficos do suplemento do Outlook. Para saber mais, consulte [Privacidade, permiss?es e seguran?a de suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/privacy-and-security).

### <a name="developer-guidelines-to-handle-pii"></a>Diretrizes de desenvolvedor para lidar com PII

Voc? pode ler as diretrizes gerais de prote??o de PII para administradores de TI e desenvolvedores em [Proteger a privacidade no desenvolvimento e teste de aplicativos de recursos humanos](http://technet.microsoft.com/en-us/library/gg447064.aspx). A seguir s?o listadas algumas diretrizes de prote??o espec?ficas de PII para voc?, como desenvolvedor de Suplementos do Office:

- O objeto [Settings](https://dev.office.com/reference/add-ins/shared/settings) destina-se a persistir configura??es e dados de estado de suplementos entre sess?es para um suplemento de conte?do ou de painel de tarefas, mas n?o armazena senhas e outros itens de PII confidenciais no objeto **Settings**. Os dados no objeto **Settings** n?o ficam vis?veis para os usu?rios finais, mas s?o armazenados como parte do formato de arquivo do documento, que est? prontamente acess?vel. Voc? deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necess?rios ao suplemento no servidor que hospeda o suplemento como um recurso protegido pelo usu?rio.

- O uso de alguns aplicativos pode revelar itens de PII. Armazene com seguran?a os dados de identidade, local, horas de acesso e outras credenciais dos usu?rios para que os dados n?o sejam disponibilizados para outros usu?rios do suplemento.

- Se o suplemento estiver dispon?vel no AppSource, o requisito do AppSource por HTTPS proteger? os itens de PII transmitidos entre o servidor Web e o dispositivo ou computador cliente. No entanto, se voc? retransmitir esses dados para outros servidores, observe o mesmo n?vel de prote??o.

- Se voc? armazenar itens de PII dos usu?rios, revele esse fato e forne?a uma maneira para que os usu?rios os inspecionem e excluam. Se voc? enviar o suplemento ao AppSource, poder? indicar na pol?tica de privacidade os dados que coleta e como eles s?o usados.

## <a name="developers-permission-choices-and-security-practices"></a>Op??es de permiss?o e pr?ticas de seguran?a de desenvolvedores

Siga estas diretrizes gerais para dar suporte ao modelo de seguran?a de Suplementos do Office e analisar detalhadamente cada tipo de suplemento.

### <a name="permissions-choices"></a>Op??es de permiss?es

A plataforma de suplemento fornece um modelo de permiss?es que o seu suplemento usa para declarar o n?vel de acesso aos dados de um usu?rio que ele exige para seus recursos. Cada n?vel de permiss?o corresponde ao subconjunto da API JavaScript para Office que seu suplemento pode usar em seus recursos. Por exemplo, a permiss?o **WriteDocument** para os suplementos do conte?do e do painel de tarefas permite acesso ao m?todo [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync), que permite que um suplemento grave no documento do usu?rio, mas n?o permite acesso a qualquer um dos m?todos para ler dados do documento. Esse n?vel de permiss?o faz sentido para suplementos que precisam apenas gravar em um documento, como um suplemento no qual o usu?rio pode consultar dados para inserir em seus documentos.

Como pr?tica recomendada, voc? deve solicitar permiss?es com base no princ?pio de _menor privil?gio_. Ou seja, voc? deve solicitar permiss?o para acessar apenas o subconjunto m?nimo da API que o suplemento requer para funcionar corretamente. Por exemplo, se o suplemento precisa apenas ler dados no documento de um usu?rio para seus recursos, voc? n?o deve solicitar mais do que a permiss?o **ReadDocument**. (Por?m, lembre-se de que a solicita??o de permiss?es insuficientes far? com que a plataforma de suplementos bloqueie o uso de algumas APIs pelo suplemento e gerar? erros em tempo de execu??o.)

Voc? especifica permiss?es no manifesto do suplemento, conforme mostrado no exemplo abaixo nesta se??o, e os usu?rios finais podem ver o n?vel de permiss?o solicitado de um suplemento antes de decidirem instalar ou ativar o suplemento pela primeira vez. Al?m disso, os suplementos do Outlook que solicitam a permiss?o **ReadWriteMailbox** exigem o privil?gio de administrador expl?cito para serem instalados.

O exemplo a seguir mostra como um suplemento de painel de tarefas especifica a permiss?o **ReadDocument** em seu manifesto. Para manter as permiss?es em destaque, outros elementos no manifesto n?o s?o exibidos.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:ver="http://schemas.microsoft.com/office/appforoffice/1.0"
           xsi:type="TaskPaneApp">

... <!-- To keep permissions as the focus, not displaying other elements. -->
  <Permissions>ReadDocument</Permissions>
...
</OfficeApp>
```

Para saber mais sobre permiss?es para suplementos de painel de tarefas e de conte?do, consulte [Solicitar permiss?es para uso da API em suplementos de conte?do e de painel de tarefas](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins).

Para saber mais sobre permiss?es para suplementos do Outlook, confira os t?picos a seguir:

- [Privacidade, permiss?es e seguran?a de suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/privacy-and-security)

- [No??es b?sicas sobre permiss?es de suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)

### <a name="same-origin-policy"></a>Pol?tica de mesma origem

Como os suplementos do Office s?o p?ginas da Web executadas em um controle de navegador da Web, eles devem seguir a pol?tica de mesma origem imposta pelo navegador: por padr?o, uma p?gina da Web em um dom?nio n?o pode fazer chamadas ao servi?o Web [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) para outro dom?nio que n?o aquele em que est? hospedada.

Uma maneira de superar essa limita??o ? usar JSON/P: forne?a um proxy para o servi?o Web incluindo uma marca **script** com um atributo **src** que aponte para algum script hospedado em outro dom?nio. Voc? pode criar as marcas**script** via programa??o gerando de forma din?mica a URL para a qual apontar o atributo **src** e passando par?metros ? URL por meio de par?metros da consulta de URI. Os provedores de servi?os Web criam e hospedam o c?digo JavaScript em URLs espec?ficas e retornam scripts diferentes, dependendo dos par?metros de consulta de URI. Em seguida, esses scripts s?o executados onde est?o inseridos e funcionam como esperado.

A seguir h? um exemplo de JSON/P no exemplo de suplemento do Outlook. 

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}
```

O Exchange e o SharePoint fornecem proxies do lado do cliente para habilitar o acesso de dom?nio cruzado. Em geral, a pol?tica de mesma origem em uma intranet n?o ? t?o estrita como na Internet. Para saber mais, confira [Pol?tica de mesma origem, parte 1: sem exibi??o](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx) e [Como lidar com limita??es de pol?tica de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md).

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a>Dicas para evitar scripts mal-intencionados entre sites

Um usu?rio mal-intencionado pode atacar a origem de um suplemento inserindo um script mal-intencionado no documento ou nos campos do suplemento. Um desenvolvedor deve processar a entrada do usu?rio para evitar a execu??o do JavaScript de um usu?rio mal-intencionado em seu dom?nio. A seguir est?o algumas boas pr?ticas a serem seguidas para manipular a entrada do usu?rio a partir de um documento ou mensagem de e-mail, ou por meio de campos em um suplemento:


- Em vez da propriedade DOM [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx), use as propriedades [innerText](https://msdn.microsoft.com/library/ms533899.aspx) e [textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent) quando apropriado. Fa?a o seguinte para o suporte entre navegadores do Internet Explorer e do Firefox:

    ```js
     var text = x.innerText || x.textContent
    ```

    Para saber mais sobre as diferen?as entre **innerText** e **textContent**, confira [Node.textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent). Para saber mais sobre a compatibilidade de DOM entre navegadores comuns, consulte [Compatibilidade de DOM W3C ? HTML](http://www.quirksmode.org/dom/w3c_html.html#t07).

- Se precisar usar **innerHTML**, verifique se a entrada do usu?rio n?o tem conte?do mal-intencionado antes de pass?-la para **innerHTML**. Para saber mais e obter um exemplo de como usar **innerHTML** com seguran?a, confira a propriedade [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx).

- Se estiver usando jQuery, use o m?todo [.text()](http://api.jquery.com/text/) em vez do m?todo [.html()](http://api.jquery.com/html/).

- Use o m?todo [toStaticHTML](http://msdn.microsoft.com/en-us/library/ie/cc848922.aspx) para remover atributos e elementos HTML din?micos da entrada dos usu?rios antes de pass?-la para **innerHTML**.

- Use a fun??o [encodeURIComponent](http://msdn.microsoft.com/en-us/library/8202bce6-1342-40dc-a5ef-ac6d210a7d15.aspx) ou [encodeURI](http://msdn.microsoft.com/en-us/library/17bab5a2-bcd4-46c2-8b52-b2b5a0ed98a3.aspx) para codificar texto que se destina a ser uma URL que vem da entrada do usu?rio ou a cont?m.

- Consulte [Desenvolver suplementos seguros](http://msdn.microsoft.com/en-us/library/windows/apps/hh849625.aspx) para obter mais pr?ticas recomendadas para criar solu??es Web mais seguras.

### <a name="tips-to-prevent-clickjacking"></a>Dicas para impedir "clickjacking"

Como os suplementos do Office s?o processados em um iframe durante a execu??o em um navegador com aplicativos de host do Office Online, use as dicas a seguir para reduzir o risco de [clickjacking](http://en.wikipedia.org/wiki/Clickjacking), uma t?cnica explorada por hackers para induzir os usu?rios a revelarem informa??es confidenciais.

Em primeiro lugar, identifique a??es confidenciais que o suplemento pode executar. Elas incluem a??es que um usu?rio n?o autorizado pode usar de forma mal-intencionada, como iniciar uma transa??o financeira ou publicar dados confidenciais. Por exemplo, o suplemento pode permitir que o usu?rio envie um pagamento a um destinat?rio definido pelo usu?rio.

Segundo, para a??es confidenciais, o suplemento deve confirmar com o usu?rio antes de executar a a??o. A confirma??o deve detalhar o efeito que a a??o ter?. Tamb?m deve detalhar como o usu?rio pode impedir a a??o, se necess?rio, escolhendo um bot?o espec?fico marcado como "N?o Permitir" ou ignorando a confirma??o.

Terceiro, para garantir que nenhum poss?vel hacker possa ocultar ou mascarar a confirma??o, voc? deve exibi-la fora do contexto do suplemento (ou seja, n?o em uma caixa de di?logo HTML).

Aqui est?o alguns exemplos de como obter uma confirma??o:

- Envie um e-mail ao usu?rio com um link de confirma??o.

- Envie uma mensagem de texto ao usu?rio com um c?digo de confirma??o para ele inserir no suplemento.

- Abra um di?logo de confirma??o em uma nova janela do navegador para uma p?gina que n?o possa ser exibida em iframe. Geralmente, esse ? o padr?o usado por p?ginas de login. Use a [API de di?logo](../develop/dialog-api-in-office-add-ins.md) para criar um novo di?logo.

Verifique tamb?m se o endere?o usado para entrar em contato com o usu?rio n?o pode ter sido fornecido por um poss?vel hacker. Por exemplo, para confirma??es de pagamento, use o endere?o arquivado na conta autorizada do usu?rio.

### <a name="other-security-practices"></a>Outras pr?ticas de seguran?a

Os desenvolvedores tamb?m devem observar as seguintes pr?ticas de seguran?a:


- Os desenvolvedores n?o devem usar controles ActiveX em Suplementos do Office, pois os controles ActiveX n?o d?o suporte ? natureza de plataforma cruzada da plataforma de suplementos.

- Os suplementos de conte?do e de painel de tarefas presumem o uso das mesmas configura??es de SSL que o Internet Explorer usa por padr?o e permitem que a maioria do conte?do seja fornecida apenas por SSL. Os suplementos do Outlook exigem que todo o conte?do seja fornecido por SSL. Os desenvolvedores devem especificar no elemento **SourceLocation** do manifesto do suplemento uma URL que use HTTPS, para identificar o local do arquivo HTML do suplemento.

    Para garantir que os suplementos n?o estejam fornecendo conte?do usando HTTP, ao test?-los os desenvolvedores devem se certificar que as seguintes configura??es est?o selecionadas no Internet Explorer e que n?o h? avisos de seguran?a aparecendo em seus cen?rios de teste:

    - Verifique se a configura??o de seguran?a **Exibir conte?do misto** da zona **Internet** est? definida para **Perguntar**. Voc? pode fazer isso selecionando o seguinte no Internet Explorer: na guia **Seguran?a** da caixa de di?logo **Op??es da Internet**, selecione a zona **Internet**, escolha **N?vel personalizado**, role at? **Exibir conte?do misto** e marque **Perguntar** se essa op??o n?o estiver marcada.

    - Verifique se a op??o **Avisar ao alterar o modo de seguran?a** est? marcada na guia **Avan?ado** da caixa de di?logo **Op??es da Internet**.

- Para garantir que os suplementos n?o usem excessivamente os recursos de mem?ria ou do n?cleo da CPU e causem a nega??o de servi?o em um computador cliente, a plataforma de suplementos estabelece limites de uso de recursos. Como parte dos testes, os desenvolvedores devem verificar se o desempenho de um suplemento est? dentro dos limites de uso de recursos.

- Antes de publicar um suplemento, os desenvolvedores devem verificar se as informa??es de identifica??o pessoal expostas nos arquivos do suplemento est?o seguras.

- Os desenvolvedores n?o devem inserir chaves usadas para acessar APIs ou servi?os de terceiros (como o Bing, Google ou Facebook) diretamente nas p?ginas HTML do suplemento. Em vez disso, devem criar um servi?o Web personalizado ou armazenar as chaves em alguma outra forma de armazenamento seguro na Web, que podem ent?o chamar para passar o valor de chave ao suplemento.

- Os desenvolvedores devem fazer o seguinte ao enviar um suplemento ? AppSource:

  - Hospedar o suplemento que est?o enviando em um servidor Web que d? suporte a SSL.
  - Produzir uma declara??o com uma pol?tica de privacidade compat?vel.
  - Estar preparados para assinar um acordo contratual ao enviar o suplemento.

Al?m das regras de uso de recursos, os desenvolvedores de suplementos do Outlook tamb?m devem verificar se os suplementos est?o de acordo com os limites para a especifica??o de regras de ativa??o e se usam a API JavaScript. Para saber mais, confira [Limites de ativa??o e API JavaScript para suplementos do Outlook](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx).

## <a name="it-administrators-control"></a>Controle de administradores de TI

Em uma configura??o corporativa, os administradores de TI t?m autoridade final para habilitar ou desabilitar o acesso ao AppSource e a cat?logos particulares.

## <a name="see-also"></a>Confira tamb?m

- [Solicitar permiss?es para uso da API em suplementos de painel de tarefas e de conte?do](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd.aspx)
- [Privacidade, permiss?es e seguran?a de suplementos do Outlook](http://msdn.microsoft.com/library/44208fc4-05d4-42d8-ab20-faa89624de1c.aspx)
- [No??es b?sicas sobre permiss?es de suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/understanding-outlook-add-in-permissions)
- [Limites de ativa??o e da API do JavaScript API para suplementos do Outlook](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx)
- [Como lidar com limita??es de pol?tica de mesma origem nos suplementos do Office](http://msdn.microsoft.com/library/36c800ae-1dda-4ea8-a558-37c89ffb161b.aspx)
- [Pol?tica de Mesma Origem](http://www.w3.org/Security/wiki/Same_Origin_Policy)
- [Pol?tica de Mesma Origem Parte 1: Sem Inspecionar](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx)
- [Pol?tica de mesma origem para JavaScript](https://developer.mozilla.org/En/Same_origin_policy_for_JavaScript)
- [Modo Protegido do IE](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
