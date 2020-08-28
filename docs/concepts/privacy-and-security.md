---
title: Privacidade e segurança para suplementos do Office
description: Saiba mais sobre os aspectos de privacidade e segurança da plataforma de suplementos do Office.
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 6707e94b53eaf714699ab666200e2c2e089b128a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293013"
---
# <a name="privacy-and-security-for-office-add-ins"></a>Privacidade e segurança para suplementos do Office

## <a name="understanding-the-add-in-runtime"></a>Noções básicas sobre o tempo de execução do suplemento

Os suplementos do Office são protegidos por um ambiente de tempo de execução de suplemento, um modelo de permissões com várias camadas e administradores de desempenho. Essa estrutura protege a experiência do usuário das seguintes maneiras: 

- O acesso ao quadro da interface do usuário do aplicativo cliente do Office é gerenciado.

- Só é permitido o acesso indireto ao thread de interface do usuário do aplicativo cliente do Office.

- As interações modais não são permitidas-por exemplo, as chamadas para JavaScript `alert` , `confirm` e `prompt` funções não são permitidas porque são modais.

Além disso, a estrutura de tempo de execução fornece os seguintes benefícios para garantir que um suplemento do Office não possa danificar o ambiente do usuário:

- Isola o processo no qual o suplemento é executado.

- Não exige substituição de .dll ou de .exe ou de componentes ActiveX.

- Facilita a instalação e a desinstalação do suplemento.

E o uso de memória, CPU e recursos de rede por suplementos do Office é governável para garantir que o bom desempenho e a confiabilidade sejam mantidos.

As seções a seguir descrevem brevemente como a arquitetura de tempo de execução dá suporte a suplementos em execução em clientes do Office em dispositivos Windows, em dispositivos Mac OS X e navegadores da web.

### <a name="clients-on-windows-and-os-x-devices"></a>Clientes para dispositivos Windows e OS X

Em clientes com suporte para dispositivos de área de trabalho e de tablet, como Excel no Windows e Outlook no Windows e no Mac, há suporte aos suplementos do Office por meio da integração de um componente no processo, o tempo de execução de suplementos do Office, que gerencia o ciclo de vida do suplemento e habilita a interoperabilidade entre o suplemento e o aplicativo cliente. A página da web do suplemento em si é hospedada fora do processo. Como mostrado na figura 1, em um dispositivo Windows para área de trabalho ou tablet, [a página da Web do suplemento é hospedada em um controle do Internet Explorer ou Microsoft Edge](browsers-used-by-office-web-add-ins.md) que, por sua vez, é hospedado em um processo de tempo de execução de suplemento que fornece segurança e isolamento de desempenho.

Nos computadores com o Windows, o Modo Protegido no Internet Explorer deve estar habilitado para a Zona de Site Restrito. Isso é normalmente ativado por padrão. Se estiver desabilitado, [ocorrerá um erro](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) ao tentar iniciar um suplemento.

*Figura 1. Ambiente de execução dos Suplementos do Office nos clientes Windows para área de trabalho e tablet*

![Infraestrutura de cliente avançado](../images/dk2-agave-overview-02.png)

Como mostrado na figura a seguir, em um computador de mesa Mac OS X, a página da Web do suplemento é hospedada em um processo de host de tempo de execução WebKit em área restrita que ajuda a fornecer um nível semelhante de segurança e proteção de desempenho. 

*Figura 2. Ambiente de execução dos Suplementos do Office nos clientes Mac OS X*

![Apps for Office runtime environment on OS X Mac](../images/dk2-agave-overview-mac-02.png)

O tempo de execução de Suplementos do Office gerencia a comunicação entre processos, a conversão de eventos e chamadas à API JavaScript em itens nativos, bem como o suporte de comunicação remota da interface do usuário para habilitar o suplemento a ser processado dentro do documento, em um painel de tarefas ou de forma adjacente a uma mensagem de e-mail, solicitação de reunião ou compromisso.

### <a name="web-clients"></a>Clientes Web

Em clientes Web com suporte, os suplementos do Office são hospedados em um **iframe** que é executado usando o atributo de **área restrita** do HTML5. Não são permitidos componentes ActiveX nem a navegação na página principal do cliente Web. O suporte a Suplementos do Office é habilitado em clientes Web por meio da integração da API JavaScript para Office. De maneira semelhante aos aplicativos cliente de área de trabalho, a API JavaScript gerencia o ciclo de vida do suplemento e a interoperabilidade entre o suplemento e o cliente Web. Essa interoperabilidade é implementada por meio de uma infraestrutura especial de comunicação de mensagens de publicação entre quadros. A mesma biblioteca JavaScript (Office.js) que é usada em clientes de área de trabalho está disponível para interagir com o cliente Web. A figura a seguir mostra a infraestrutura que oferece suporte a suplementos no Office em execução no navegador, e os componentes relevantes (o cliente da Web, o **iframe**, o tempo de execução de suplementos do Office e a API JavaScript para Office) que são necessários para dar suporte a eles.


*Figura 3. Infraestrutura que dá suporte aos Suplementos do Office nos clientes Web do Office*

![Infraestrutura do cliente Web](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a>Integridade do suplemento no AppSource

Você pode disponibilizar os Suplementos do Office para o público publicando-os no AppSource, que impõe as seguintes medidas para manter a integridade dos suplementos:


- Requer que o servidor host de um Suplemento do Office sempre use o protocolo SSL para se comunicar.

- Requer que um desenvolvedor forneça uma prova de identidade, um acordo contratual e uma política de privacidade compatível para enviar suplementos.

- Garante que a origem dos suplementos seja acessível no modo somente leitura.

- Dá suporte a um sistema de revisão pelo usuário para os suplementos disponíveis para promover uma comunidade autovigilante.

## <a name="addressing-end-users-privacy-concerns"></a>Lidar com as preocupações de privacidade dos usuários finais

Esta seção descreve a proteção oferecida pela plataforma de Suplementos do Office da perspectiva do cliente (usuário final) e fornece as diretrizes sobre como dar suporte às expectativas dos usuários e como manipular com segurança as PII (informações de identificação pessoal) dos usuários.

### <a name="end-users-perspective"></a>Perspectiva dos usuários finais

Os Suplementos do Office são criados usando tecnologias da Web que são executadas em um controle de navegador ou **iframe**. Por causa disso, o uso de suplementos é semelhante à navegação em sites da Web na Internet ou intranet. Os suplementos podem ser externos a uma organização (se você adquirir o suplemento do AppSource) ou internos (se você adquirir o suplemento de um catálogo de suplementos do Exchange Server, catálogo de aplicativos do SharePoint ou compartilhamento de arquivos na rede de uma organização). Os suplementos têm acesso limitado à rede, e a maioria dos suplementos pode ler ou gravar no documento ativo ou no item de email. A plataforma de suplemento aplica certas restrições antes de um usuário ou administrador instalar ou iniciar um suplemento. Mas, como acontece com qualquer modelo de extensibilidade, os usuários devem ser cautelosos antes de iniciar um complemento desconhecido.

A plataforma de suplementos lida com as preocupações com privacidade dos usuários finais das seguintes maneiras:

- Os dados comunicados com o servidor Web que hospeda um suplemento de conteúdo, do Outlook ou de painel de tarefas, bem como a comunicação entre o suplemento e quaisquer serviços Web que ele usa, devem ser criptografados usando o protocolo SSL.

- Antes de instalar um suplemento do AppSource, o usuário pode exibir a política de privacidade e os requisitos desse suplemento. Além disso, os suplementos do Outlook que interagem com caixas de correio dos usuários expõem as permissões específicas das quais precisam. O usuário pode examinar os termos de uso, as permissões solicitadas e a política de privacidade antes de instalar um suplemento do Outlook.

- Ao compartilhar um documento, os usuários também compartilham suplementos que foram inseridos no documento ou associados a ele. Se um usuário abrir um documento que contenha um suplemento que o usuário não tenha usado anteriormente, o aplicativo cliente do Office solicitará que o usuário Conceda permissão para que o suplemento seja executado no documento. Em um ambiente organizacional, o aplicativo cliente do Office também solicita o usuário se o documento vier de uma fonte externa.

- Os usuários podem habilitar ou desabilitar o acesso ao AppSource. Para suplementos de conteúdo e de painel de tarefas, os usuários gerenciam o acesso a suplementos e catálogos confiáveis da **central de confiabilidade** no cliente do Office do host (aberto nas opções de **arquivo**configurações da central de confiabilidade da central de confiabilidade dos  >  **Options**  >  **Trust Center**  >  **Trust Center Settings**  >  **catálogos de suplementos confiáveis**). Para os suplementos do Outlook, os usos podem gerenciar suplementos escolhendo o botão **gerenciar suplementos** : no Outlook no Windows, escolha **arquivo**  >  **gerenciar suplementos**. No Outlook no Mac, escolha o botão **gerenciar suplementos** na barra de suplementos. No Outlook na Web, escolha o menu **Configurações** (ícone de engrenagem) > **Gerenciar suplementos**. Os administradores também podem gerenciar este acesso [usando a política de grupo](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).

- O design da plataforma do suplemento fornece segurança e desempenho aos usuários finais das seguintes maneiras:

  - Um suplemento do Office é executado em um controle de navegador da Web hospedado em um ambiente de tempo de execução do suplemento separado do aplicativo cliente do Office. Esse design oferece isolamento de segurança e desempenho do aplicativo cliente.

  - A execução em um controle de navegador da Web permite que o suplemento faça quase tudo que uma página da Web regular em execução em um navegador pode fazer, mas, ao mesmo tempo, restringe o suplemento a observar a política de mesma origem para o isolamento de domínio e as zonas segurança.

Os suplementos do Outlook fornecem recursos adicionais de segurança e desempenho por meio do monitoramento de uso de recursos específicos do suplemento do Outlook. Para saber mais, consulte [Privacidade, permissões e segurança de suplementos do Outlook](../outlook/privacy-and-security.md).

### <a name="developer-guidelines-to-handle-pii"></a>Diretrizes de desenvolvedor para lidar com PII

A seguir são listadas algumas diretrizes de proteção específicas de PII para desenvolvedores de Suplementos do Office:

- O objeto [Settings](/javascript/api/office/office.settings) destina-se a persistir configurações e dados de estado de suplementos entre sessões para um suplemento de conteúdo ou de painel de tarefas, mas não armazena senhas e outros itens de PII confidenciais no objeto **Settings**. Os dados no objeto **Settings** não ficam visíveis para os usuários finais, mas são armazenados como parte do formato de arquivo do documento, que está prontamente acessível. Você deve limitar o uso de PII pelo suplemento e armazenar quaisquer itens de PII necessários ao suplemento no servidor que hospeda o suplemento como um recurso protegido pelo usuário.

- O uso de alguns aplicativos pode revelar itens de PII. Armazene com segurança os dados de identidade, local, horas de acesso e outras credenciais dos usuários para que os dados não sejam disponibilizados para outros usuários do suplemento.

- Se o suplemento estiver disponível no AppSource, o requisito do AppSource por HTTPS protegerá os itens de PII transmitidos entre o servidor Web e o dispositivo ou computador cliente. No entanto, se você retransmitir esses dados para outros servidores, observe o mesmo nível de proteção.

- Se você armazenar itens de PII dos usuários, revele esse fato e forneça uma maneira para que os usuários os inspecionem e excluam. Se você enviar o suplemento ao AppSource, poderá indicar na política de privacidade os dados que coleta e como eles são usados.

## <a name="developers-permission-choices-and-security-practices"></a>Opções de permissão e práticas de segurança de desenvolvedores

Siga estas diretrizes gerais para dar suporte ao modelo de segurança de Suplementos do Office e analisar detalhadamente cada tipo de suplemento.

### <a name="permissions-choices"></a>Opções de permissões

A plataforma de suplementos fornece um modelo de permissões que o suplemento usa para declarar o nível de acesso aos dados de um usuário de que necessita para seus recursos. Cada nível de permissão corresponde ao subconjunto da API JavaScript para Office que o suplemento tem permissão para usar para seus recursos. Por exemplo, a permissão **WriteDocument** para suplementos de painel de tarefas e de conteúdo permite o acesso ao método [Document. setSelectedDataAsync](/javascript/api/office/office.document) que permite que um suplemento grave no documento do usuário, mas não permite o acesso a qualquer um dos métodos de leitura de dados do documento. Esse nível de permissão faz sentido para suplementos que só precisam gravar em um documento, como um suplemento em que o usuário pode consultar dados para inserir em seu documento.

Como prática recomendada, você deve solicitar permissões com base no princípio de _menor privilégio_. Ou seja, você deve solicitar permissão para acessar apenas o subconjunto mínimo da API que o suplemento requer para funcionar corretamente. Por exemplo, se o suplemento precisa apenas ler dados no documento de um usuário para seus recursos, você não deve solicitar mais do que a permissão **ReadDocument**. (Porém, lembre-se de que a solicitação de permissões insuficientes fará com que a plataforma de suplementos bloqueie o uso de algumas APIs pelo suplemento e gerará erros em tempo de execução.)

Você especifica permissões no manifesto do suplemento, conforme mostrado no exemplo abaixo nesta seção, e os usuários finais podem ver o nível de permissão solicitado de um suplemento antes de decidirem instalar ou ativar o suplemento pela primeira vez. Além disso, os suplementos do Outlook que solicitam a permissão **ReadWriteMailbox** exigem privilégio de administrador explícito para instalar o.

O exemplo a seguir mostra como um suplemento de painel de tarefas especifica a permissão **ReadDocument** em seu manifesto. Para manter as permissões em destaque, outros elementos no manifesto não são exibidos.

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

Para saber mais sobre permissões para suplementos de painel de tarefas e de conteúdo, consulte [Solicitar permissões para uso da API em suplementos ](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

Para saber mais sobre permissões para suplementos do Outlook, confira os tópicos a seguir:

- [Privacidade, permissões e segurança de suplementos do Outlook](../outlook/privacy-and-security.md)

- [Noções básicas sobre permissões de suplemento do Outlook](../outlook/understanding-outlook-add-in-permissions.md)

### <a name="same-origin-policy"></a>Política de mesma origem

Como os suplementos do Office são páginas da Web executadas em um controle de navegador da Web, eles devem seguir a política de mesma origem imposta pelo navegador: por padrão, uma página da Web em um domínio não pode fazer chamadas ao serviço Web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) para outro domínio que não aquele em que está hospedada.

Uma maneira de superar essa limitação é usar JSON/P--fornecer um proxy para o serviço Web, incluindo uma marca de **script** com um atributo **src** que aponta para um script hospedado em outro domínio. Você pode criar as marcas**script** via programação gerando de forma dinâmica a URL para a qual apontar o atributo **src** e passando parâmetros à URL por meio de parâmetros da consulta de URI. Os provedores de serviços Web criam e hospedam o código JavaScript em URLs específicas e retornam scripts diferentes, dependendo dos parâmetros de consulta de URI. Em seguida, esses scripts serão executados onde estiverem inseridos e funcionarão como esperado.

A seguir há um exemplo de JSON/P no exemplo de suplemento do Outlook. 

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

O Exchange e o SharePoint fornecem proxies do lado do cliente para habilitar o acesso de domínio cruzado. Em geral, a política de mesma origem em uma intranet não é tão estrita como na Internet. Para saber mais, confira [Política de mesma origem, parte 1: sem exibição](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking) e [Como lidar com limitações de política de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md).

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a>Dicas para evitar scripts mal-intencionados entre sites

Um usuário mal-intencionado pode atacar a origem de um suplemento inserindo um script mal-intencionado por meio do documento ou de campos no suplemento. Um desenvolvedor deve processar a entrada do usuário para evitar a execução de JavaScript de um usuário mal-intencionado em seu domínio. Estas são algumas práticas recomendadas a seguir para manipular a entrada do usuário em um documento ou uma mensagem de e-mail ou por meio de campos em um suplemento:


- Em vez da propriedade DOM [innerHTML](https://developer.mozilla.org/docs/Web/API/Element/innerHTML), use as propriedades [innerText](https://developer.mozilla.org/docs/Web/API/Node/innerText) e [textContent](https://developer.mozilla.org/docs/DOM/Node.textContent) quando apropriado. Faça o seguinte para o suporte entre navegadores do Internet Explorer e do Firefox:

    ```js
     var text = x.innerText || x.textContent
    ```

    Para obter informações sobre as diferenças entre **InnerText** e **textcontent**, consulte [node. textcontent](https://developer.mozilla.org/docs/DOM/Node.textContent). Para saber mais sobre a compatibilidade de DOM entre navegadores comuns, consulte [Compatibilidade de DOM W3C ‒ HTML](https://www.quirksmode.org/dom/w3c_html.html#t07).

- Se você precisar usar **InnerHtml**, certifique-se de que a entrada do usuário não contenha conteúdo mal-intencionado antes de passá-la para **InnerHtml**. Para obter mais informações e um exemplo de como usar o **InnerHtml** com segurança, consulte Propriedade [InnerHtml](https://developer.mozilla.org/docs/Web/API/Element/innerHTML) .

- Se estiver usando jQuery, use o método [.text()](https://api.jquery.com/text/) em vez do método [.html()](https://api.jquery.com/html/).

- Use o método [toStaticHTML](https://developer.mozilla.org/en-US/docs/Web/HTML/Reference) para remover atributos e elementos HTML dinâmicos da entrada dos usuários antes de passá-la para **innerHTML**.

- Use a função [encodeURIComponent](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuricomponent) ou [encodeURI](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuri) para codificar texto que se destina a ser uma URL que vem da entrada do usuário ou a contém.

- Consulte [Desenvolver suplementos seguros](/previous-versions/windows/apps/hh849625(v=win.10)) para obter mais práticas recomendadas para criar soluções Web mais seguras.

### <a name="tips-to-prevent-clickjacking"></a>Dicas para impedir "clickjacking"

Como os suplementos do Office são renderizados em um iframe ao executar em um navegador com aplicativos cliente do Office, use as seguintes dicas para minimizar o risco de [clickjacking](https://en.wikipedia.org/wiki/Clickjacking) --uma técnica usada por hackers para enganar os usuários na revelação de informações confidenciais.

Em primeiro lugar, identifique ações confidenciais que o suplemento pode executar. Elas incluem ações que um usuário não autorizado pode usar de forma mal-intencionada, como iniciar uma transação financeira ou publicar dados confidenciais. Por exemplo, o suplemento pode permitir que o usuário envie um pagamento a um destinatário definido pelo usuário.

Segundo, para ações confidenciais, o suplemento deve confirmar com o usuário antes de executar a ação. A confirmação deve detalhar o efeito que a ação terá. Também deve detalhar como o usuário pode impedir a ação, se necessário, escolhendo um botão específico marcado como "Não Permitir" ou ignorando a confirmação.

Terceiro, para garantir que nenhum possível hacker possa ocultar ou mascarar a confirmação, você deve exibi-la fora do contexto do suplemento (ou seja, não em uma caixa de diálogo HTML).

Aqui estão alguns exemplos de como obter uma confirmação:

- Envie um e-mail ao usuário com um link de confirmação.

- Envie uma mensagem de texto ao usuário com um código de confirmação para ele inserir no suplemento.

- Abra um diálogo de confirmação em uma nova janela do navegador para uma página que não possa ser exibida em iframe. Geralmente, esse é o padrão usado por páginas de login. Use a [API de diálogo](../develop/dialog-api-in-office-add-ins.md) para criar um novo diálogo.

Verifique também se o endereço usado para entrar em contato com o usuário não pode ter sido fornecido por um possível hacker. Por exemplo, para confirmações de pagamento, use o endereço arquivado na conta autorizada do usuário.

### <a name="other-security-practices"></a>Outras práticas de segurança

Os desenvolvedores também devem observar as seguintes práticas de segurança:


- Os desenvolvedores não devem usar controles ActiveX em Suplementos do Office, pois os controles ActiveX não dão suporte à natureza de plataforma cruzada da plataforma de suplementos.

- Os suplementos de conteúdo e de painel de tarefas assumem o uso das mesmas configurações de SSL que o navegador usa por padrão e permitem que a maioria do conteúdo seja fornecido apenas por SSL. Os suplementos do Outlook exigem que todo o conteúdo seja fornecido por SSL. Os desenvolvedores devem especificar no elemento **SourceLocation** do manifesto do suplemento uma URL que use HTTPS, para identificar o local do arquivo HTML para o suplemento.

    Para garantir que os suplementos não estejam fornecendo conteúdo usando HTTP, ao testá-los, os desenvolvedores devem se certificar que as seguintes configurações estão selecionadas nas **Opções de Internet** no **Painel de Controle** e que não ocorra avisos de segurança em seus cenários de teste:

    - Certifique-se de que a configuração de segurança, **Exibir conteúdo misto**, para a zona da **Internet** esteja definida como **avisar**. Você pode fazer isso selecionando o seguinte em **Opções da Internet**: na **guia Segurança** , selecione a zona da **Internet** , selecione **nível personalizado**, rolar para procurar por **Exibir conteúdo misto**e selecione **avisar** se ainda não estiver selecionado.

    - Verifique se a opção **Avisar ao alterar o modo de segurança** está marcada na guia **Avançado** da caixa de diálogo **Opções da Internet**.

- Para garantir que os suplementos não usem excessivamente os recursos de memória ou do núcleo da CPU e causem a negação de serviço em um computador cliente, a plataforma de suplementos estabelece limites de uso de recursos. Como parte dos testes, os desenvolvedores devem verificar se o desempenho de um suplemento está dentro dos limites de uso de recursos.

- Antes de publicar um suplemento, os desenvolvedores devem verificar se as informações de identificação pessoal expostas nos arquivos do suplemento estão seguras.

- Os desenvolvedores não devem inserir chaves usadas para acessar APIs ou serviços de terceiros (como o Bing, Google ou Facebook) diretamente nas páginas HTML do suplemento. Em vez disso, devem criar um serviço Web personalizado ou armazenar as chaves em alguma outra forma de armazenamento seguro na Web, que podem então chamar para passar o valor de chave ao suplemento.

- Os desenvolvedores devem fazer o seguinte ao enviar um suplemento à AppSource:

  - Hospedar o suplemento que estão enviando em um servidor Web que dê suporte a SSL.
  - Produzir uma declaração com uma política de privacidade compatível.
  - Estar preparados para assinar um acordo contratual ao enviar o suplemento.

Além das regras de uso de recursos, os desenvolvedores de suplementos do Outlook também devem verificar se os suplementos estão de acordo com os limites para a especificação de regras de ativação e se usam a API JavaScript. Para saber mais, confira [Limites de ativação e API JavaScript para suplementos do Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

## <a name="it-administrators-control"></a>Controle de administradores de TI

Em uma configuração corporativa, os administradores de TI têm autoridade final para habilitar ou desabilitar o acesso ao AppSource e a catálogos particulares.

O gerenciamento e a execução das configurações do Office são feitos com as configurações de política de grupo. Eles são configuráveis através da [Ferramenta de Implantação do Office](/deployoffice/overview-of-the-office-2016-deployment-tool), em conjunto com a [Ferramenta de Personalização do Office](/deployoffice/overview-of-the-office-customization-tool-for-click-to-run).

| Nome da configuração | Descrição |
|--------------|-------------|
| Permitir catálogos e suplementos da Web inseguros | Permite com que os usuários executem suplementos não seguros, que são suplementos que contêm uma página da Web ou locais de catálogo não protegidos por SSL (https://) e não estão nas zonas de Internet dos usuários. |
| Bloquear Suplementos da Web | Permite com que você impeça usuários de utilizar suplementos da Web. |
| Bloquear a Office Store |  Permite com que você impeça usuários de utilizar ou inserir suplementos da Web que vierem da Office Store.  |

> [!IMPORTANT]
> Se os seus grupos de trabalho estiverem usando diversas edições do Office, os parâmetros da política de grupo devem ser configurados para cada edição. Consulte o [Usando Política de Grupo para gerenciar como os usuários podem instalar e usar aplicativos do Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) do artigo [Visão geral dos aplicativos do Office 2013](/previous-versions/office/office-2013-resource-kit/jj219429(v%3doffice.15)) para detalhes sobre as configuração da política de grupo do Office 2013.

## <a name="see-also"></a>Confira também

- [Solicitar permissões para uso da API em suplementos ](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [Privacidade, permissões e segurança de suplementos do Outlook](../outlook/privacy-and-security.md)
- [Noções básicas sobre permissões de suplemento do Outlook](../outlook/understanding-outlook-add-in-permissions.md)
- [Limites de ativação e da API do JavaScript API para suplementos do Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Como lidar com limitações de política de mesma origem nos suplementos do Office](../develop/addressing-same-origin-policy-limitations.md)
- [Política de Mesma Origem](https://www.w3.org/Security/wiki/Same_Origin_Policy)
- [Política de Mesma Origem Parte 1: Sem Inspecionar](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking)
- [Política de mesma origem para JavaScript](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)
- [Modo Protegido do IE](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
