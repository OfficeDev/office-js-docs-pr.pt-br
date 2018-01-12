# <a name="ux-design-pattern-templates-for-office-add-ins"></a>Modelos de padrão de design da experiência do usuário para Suplementos do Office 

O [projeto de padrões de design da experiência do usuário para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projeto de padrões de design da experiência do usuário para Suplementos do Office") contém arquivos HTML, JavaScript e CSS que você pode usar para criar a experiência de usuário para seu suplemento.   

Use o projeto de padrões de design da experiência do usuário para:

* Aplicar soluções a cenários comuns de clientes.
* Aplicar as práticas recomendadas de design.
* Incorporar componentes e estilos do [Office UI Fabric](https://dev.office.com/fabric#/get-started).
* Criar suplementos que se integram visualmente à interface do usuário padrão do Office.  

## <a name="using-the-ux-design-patterns"></a>Usar os padrões de design da experiência do usuário

Você pode usar as [especificações de padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns) como um guia para projetar seu próprio Suplemento do Office ou pode adicionar o [código-fonte](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) diretamente ao seu projeto.

Para usar as especificações para criar uma simulação da interface do usuário do seu suplemento:

1. Baixe arquivos de ativos de design e comece a criar sua própria interface do usuário:
    * [Componentes de design da experiência do usuário do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_components.ai) (arquivo do Adobe Illustrator)
    * [Padrões de design da experiência do usuário do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_patterns.ai) (arquivo do Adobe Illustrator) ou o
    * [Protótipo de design da experiência do usuário do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_prototype.xd) (arquivo do Adobe Experience Design)
2. Você encontra orientações nos seguintes artigos:
    * [Padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/README.md)
    * Práticas recomendadas para [projetar seus Suplementos do Office](https://dev.office.com/docs/add-ins/design/add-in-design)
    * [Kits de ferramentas do Office UI Fabric](https://developer.microsoft.com/pt-BR/fabric#/resources)

Para adicionar o código-fonte:

1. Clone o [repositório do projeto de padrões de design da experiência do usuário para suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projeto de padrões de design da experiência do usuário para suplementos do Office"). 
2. Copie a [pasta de ativos](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) e a pasta de código do padrão específico que você escolheu para o projeto do seu suplemento.  
3. Incorpore o padrão individual ao suplemento. Por exemplo:
    - Edite o local de origem ou a URL de comando do suplemento no manifesto.
    - Use o padrão de design da experiência do usuário como modelo para outras páginas.
    - Crie um link de ou para o padrão de design da experiência do usuário.

> **Observação:** algumas especificações do padrão da experiência do usuário não correspondem ao código-fonte. Estamos trabalhando arduamente para alinhar todos os ativos. Observe também que algumas especificações são apresentadas como arquivadas. Estamos avaliando essas especificações arquivadas para verificar o valor para a plataforma. Cada padrão pretende representar um modelo exclusivo e um padrão de interação. Os padrões não devem se sobrepor entre si e devem ser diferenciados dos componentes do Office Fabric UI.

## <a name="types-of-ux-design-patterns"></a>Tipos de padrões de design da experiência do usuário
### <a name="generic-pages"></a>Páginas genéricas

Modelos de páginas genéricas podem ser aplicados a qualquer página no seu suplemento e não têm uma finalidade especial. Um exemplo de uma página com finalidade especial seria qualquer um dos padrões de apresentação. A lista a seguir descreve os tipos de páginas genéricas disponíveis:

* **Página de chegada**: uma página de suplemento padrão, por exemplo, a página a que um usuário chega após uma tela de apresentação ou processo de login. 
    * Saiba mais sobre as diretrizes para adotar a [linguagem de design do Office](https://dev.office.com/docs/add-ins/design/add-in-design-language) no seu suplemento.
    * [Código da página de chegada](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* **Imagem da marca na barra da marca**: a página de chegada com uma imagem no rodapé que representa sua marca. 
    * [Especificação da barra da marca](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [Código da barra da marca](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th>Chegada</th><th>Barra da marca</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../../images/landing.page.PNG" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../../images/brand.bar.PNG" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a>Tela de apresentação

Uma tela de apresentação é a experiência que o usuário tem ao abrir o suplemento pela primeira vez. Estão disponíveis os seguintes modelos de padrão de design de apresentação: 

* **Etapas para iniciar**: fornece aos usuários uma lista ordenada de etapas a serem executadas para começar a usar o suplemento. 
    * [Especificação das etapas para iniciar](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Conforme avaliamos seu valor, confira também a [especificação do valor da Tela de Apresentação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)  
    * [Código das etapas para iniciar](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* **Valor**: comunica a proposta de valor do suplemento.
    * [Especificação do valor](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [Código do valor](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* **Vídeo**: mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.
    * [Especificação do vídeo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [Código do vídeo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* **Passo a passo**: apresenta aos usuários uma série de recursos ou informações antes que eles comecem a usar o suplemento.
    * [Especificação do carrossel](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md)
        * Observação: Esse padrão de design da experiência do usuário foi renomeado para “Carrossel”. Especificações anteriores o chamavam de “Painel de Paginação”. Os ativos do código o chamam de "Passo a Passo da Tela de Apresentação". 
    * [Código do passo a passo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)

A [Office Store](https://msdn.microsoft.com/pt-BR/library/office/jj220033.aspx) possui um sistema que gerencia versões de avaliação de um suplemento, mas se você deseja controlar a interface do usuário da experiência de avaliação para o seu suplemento, use os seguintes padrões:

* **Avaliação**: mostra aos usuários como começar a usar uma versão de avaliação do suplemento.
    * [Especificação da avaliação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF.
    * [Código da avaliação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* **Recurso de avaliação**: avisa aos usuários que o recurso que eles estão tentando usar não está disponível na versão de avaliação do suplemento. Como alternativa, se o seu suplemento for gratuito, mas houver um recurso que exija uma assinatura, considere usar esse padrão. Você também pode usar esse padrão para fornecer uma experiência de versão limitada após o término do período de avaliação.
    * [Especificação do recurso de avaliação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
    * [Código do recurso de avaliação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> **Importante:** Se você decidir gerenciar sua própria avaliação e não usar a Office Store, certifique-se de incluir a marca **Compra adicional pode ser necessária** nas anotações da avaliação no painel do vendedor.

Considere se mostrar aos usuários a tela de apresentação uma ou muitas vezes é importante para seu cenário. Por exemplo, se seu suplemento for usado periodicamente, talvez os usuários se esqueçam de como usá-lo, e pode ser útil ver a tela de apresentação mais de uma vez. 

 <table>
 <tr><th>Etapas para iniciar</th><th>Valor</th><th>Vídeo</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../../images/instruction.step.PNG" alt="instruction steps" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../../images/value.placemat.PNG" alt="value placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../../images/video.placemat.PNG" alt="video placemat" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Página inicial de explicação passo a passo</th><th>Avaliação</th><th>Recurso de avaliação</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../../images/walkthrough1.PNG" alt="walkthrough 1" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../../images/trial.placemat.PNG" alt="trial placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../../images/trial.placemat.feature.PNG" alt="trial placemat feature" style="width: 264px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a>Navegação

Os usuários precisam navegar entre as diferentes páginas do seu suplemento. Os seguintes modelos de navegação mostram diferentes opções que você pode usar para organizar páginas e comandos no seu suplemento.

* **Botões Voltar e Próxima página**: mostra um painel de tarefas com os botões Voltar e Próxima página. Use esse padrão para garantir que os usuários sigam uma série de etapas ordenadas.
    * [Especificação do Botão Voltar e Próxima Página](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [Código do Botão Voltar e Próxima Página](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* **Navegação**: mostra um menu, comumente conhecido como menu vertical clicável, com itens de menu da página em um painel de tarefas. 
    * [Especificação da navegação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [Código da navegação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* **Navegação com comandos**: mostra o menu vertical clicável com botões de comando (ou de ação) em um painel de tarefas. Use este padrão quando quiser fornecer as opções de navegação e comando juntas. 
    * [Especificação da navegação com comandos](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [Código da navegação com comandos](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* **Navegação dinâmica**: mostra a navegação dinâmica dentro de um painel de tarefas. Use a navegação dinâmica para permitir que os usuários naveguem entre diferentes conteúdos.
    * [Especificação da navegação dinâmica](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [Código da navegação dinâmica](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* **Barra de guias**: mostra a navegação usando botões com texto empilhado na vertical e ícones. Use a barra de guias para fornecer a navegação usando guias com títulos curtos e descritivos.
    * [Especificação da barra de guias](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [Código da barra de guias](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th>Botão Voltar</th><th>Navegação</th><th>Navegação com comandos</th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../../images/back.button.png" alt="back button" style="width: 264px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../../images/navigation.png" alt="navigation" style="width: 264px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../../images/navigation.commands.png" alt="navigation with commands" style="width: 264px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th>Navegação dinâmica</th><th>Barra de guias</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../../images/pivot.png" alt="pivot navigation" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../../images/tab.bar.png" alt="tab bar" style="width: 264px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a>Notificações

Seu suplemento pode notificar os usuários sobre eventos, como erros, ou sobre o progresso de várias maneiras. Estão disponíveis os seguintes modelos de notificação: 

* **Caixa de diálogo incorporada**: mostra uma caixa de diálogo dentro do painel de tarefas que fornece informações e, opcionalmente, uma experiência interativa, usando botões ou outros controles. Considere usar uma para solicitar que o usuário confirme uma ação. Use o padrão da caixa de diálogo incorporada quando quiser manter a experiência do usuário no painel de tarefas.
    * [Especificação da caixa de diálogo incorporada](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [Código da caixa de diálogo incorporada](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* **Mensagem embutida**: indica o erro, o êxito ou as informações, e pode ser exibida em um local especificado no painel de tarefas. Por exemplo, se um usuário inserir um endereço de email com formato incorreto em uma caixa de texto, uma mensagem de erro aparecerá logo abaixo da caixa. 
    * [Especificação da mensagem embutida](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
    * [Código da mensagem embutida](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* **Faixa de mensagem**: fornece informações e, opcionalmente, uma chamada simples à ação em uma faixa que pode ser recolhida para uma única linha, expandida para várias linhas ou ignorada. Use faixas de mensagem para relatar uma atualização de serviço ou uma dica útil quando o suplemento inicia. 
    * [Especificação da faixa de mensagem](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
    * [Código da faixa de mensagem](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* **Barra de progresso**: indica o andamento de um processo demorado e síncrono, como uma tarefa de configuração que deve ser concluída antes do usuário poder executar outras ações. É uma página de transição separada que também reforça a marca do suplemento. Use uma barra de progresso quando o processo puder enviar medições periódicas de seu progresso ao suplemento.
    * [Especificação da barra de progresso](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [Código da barra de progresso](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* **Controle giratório**: indica que um processo demorado e síncrono está em andamento, mas não fornece uma indicação do progresso. É uma página de transição separada que também reforça a marca do suplemento. Use um controle giratório quando o suplemento não puder saber com segurança o progresso de um processo. 
    * [Especificação do controle giratório](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [Código do controle giratório](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* **Notificação do sistema**: fornece uma breve mensagem que desaparece após alguns segundos. Como o usuário pode não ver a mensagem, use a notificação do sistema somente para informações não essenciais. É uma boa opção notificar os usuários sobre um evento em um sistema remoto, como o recebimento de um email.
    * [Especificação da notificação do sistema](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [Código da notificação do sistema](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th>Caixa de diálogo incorporada</th><th>Mensagem embutida</th><th>Faixa de mensagem</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../../images/embedded.dialog.PNG" alt="embedded dialog" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../../images/inline.message.PNG" alt="inline message" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../../images/message.banner.PNG" alt="message banner" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Barra de progresso</th><th>Controle giratório</th><th>Notificação do sistema</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../../images/progress.bar.PNG" alt="progress bar" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../../images/spinner.PNG" alt="spinner" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../../images/toast.PNG" alt="toast" style="width: 264px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a>Componentes gerais

A seguir apresentamos os componentes gerais que você pode usar em seus suplementos em vários cenários.  

#### <a name="client-dialog-boxes"></a>Caixas de diálogo de cliente

As caixas de diálogo de cliente oferecem outra maneira para seus usuários trabalharem com seu suplemento fora de um painel de tarefas. Estão disponíveis os seguintes modelos de caixa de diálogo:

* **Caixa de diálogo typeramp**: mostra uma caixa de diálogo com conteúdo textual. Use a caixa de diálogo typeramp para exibir informações elaborativas para os usuários. 
    * Saiba mais sobre como projetar [caixas de diálogo em Suplementos do Office](https://dev.office.com/docs/add-ins/design/dialog-boxes). Além disso, siga nossa diretrizes para [Tipografia em Suplementos do Office](https://dev.office.com/docs/add-ins/design/add-in-design-language#typography).
    * [Código da caixa de diálogo typeramp](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* **Caixa de diálogo de alerta**: mostra uma caixa de alerta com informações importantes aos usuários, como erros ou notificações.  
    * [Especificação da caixa de diálogo de alerta](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
    * [Código da caixa de diálogo de alerta](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* **Caixa de diálogo de navegação**: mostra uma caixa de diálogo com navegação. Use a caixa de diálogo de navegação para permitir que os usuários naveguem entre conteúdos diferentes. 
    * Saiba mais sobre como projetar [caixas de diálogo em Suplementos do Office](https://dev.office.com/docs/add-ins/design/dialog-boxes). Além disso, saiba como usar os [componentes dinâmicos em Suplementos do Office](https://dev.office.com/docs/add-ins/design/pivot) usando o Office UI Fabric.
    * [Código da caixa de diálogo de navegação](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th>Caixa de diálogo typeramp</th><th>Caixa de diálogo de alerta</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../../images/typeramp.dialog.png" alt="typeramp dialog" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../../images/alert.dialog.png" alt="alert dialog" style="width: 264px;"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>Caixa de diálogo de navegação</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../../images/navigation.dialog.png" alt="navigation dialog" style="width: 300px;"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a>Comentários e classificações

Para melhorar a visibilidade e a adoção do seu suplemento, você deve fornecer aos usuários a capacidade de classificar e comentar sobre seu suplemento na Office Store. Este padrão mostra dois métodos para apresentar comentários e classificações de dentro do suplemento:

- Comentários iniciados pelo usuário: um usuário opta por enviar comentários usando o menu de navegação (por exemplo, usando o link **Enviar comentários**) ou um ícone no rodapé.
- Comentários iniciados pelo sistema: depois que o suplemento é executado três vezes, o usuário é solicitado a fornecer comentários por meio de uma faixa de mensagem.

Os dois métodos abrem uma caixa de diálogo que contém a página Office Store para o suplemento.

* [Especificação dos comentários e classificações](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf)
    * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
* [Código dos comentários e classificações](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

>**Importante:** esse padrão atualmente aponta para a página inicial da Office Store. Certifique-se de atualizar essa URL para a URL da página do seu suplemento na Office Store.

 <table>
 <tr><th>Comentários e classificações</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../../images/feedback.ratings.PNG" alt="Feedback and Ratings" style="width: 264px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a>Configurações e privacidade

Suplementos podem precisar de uma página de configurações que permita aos usuários definir as configurações que controlam o comportamento do suplemento. Além disso, convém fornecer aos usuários as políticas de privacidade a que seu suplemento está sujeito. 

* **Configurações**: mostra um painel de tarefas com componentes de configuração que controlam o comportamento do suplemento. Uma página de configurações fornece opções para o usuário escolher.
    * [Especificação das configurações](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [Código das configurações](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* **Política de privacidade**: mostra o painel de tarefas com informações importantes sobre as políticas de privacidade. 
    * [Especificação da política de privacidade](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf)
        * Esse padrão de design da experiência do usuário foi arquivado. Enquanto avaliamos seu valor, confira o PDF acima.
    * [Código da política de privacidade](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th>Configurações</th><th>Política de privacidade</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../../images/privacy.policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="additional-resources"></a>Recursos adicionais

* [Práticas recomendadas para desenvolvimento de suplementos do Office](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
