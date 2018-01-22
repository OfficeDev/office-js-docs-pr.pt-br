# <a name="office-add-in-design-language"></a>Linguagem design de suplemento do Office

A linguagem de design do Office é um sistema visual claro e simples que garante a consistência nas experiências. Ela contém um conjunto de elementos visuais que definem as interfaces do Office, incluindo: 

- Um tipo de fonte padrão
- Uma paleta de cores comuns
- Um conjunto de pesos e tamanhos tipográficos
- Diretrizes de ícones
- Ativos de ícones compartilhados
- Definições de animação
- Componentes comuns

O [Office UI Fabric](https://dev.office.com/fabric) é a estrutura de front-end oficial para criação com a linguagem de design do Office. O uso do Fabric é opcional, mas é a maneira mais rápida de garantir que os suplementos sejam como uma extensão natural do Office. Tire proveito do Fabric para projetar e criar suplementos que complementam o Office.

Vários suplementos do Office estão associados a uma marca pré-existente. Você pode manter uma marca forte e sua linguagem visual ou de componente no suplemento. Procure oportunidades para manter sua própria linguagem visual durante a integração ao Office. Considere maneiras de substituir cores, tipografia, ícones ou outros elementos estilísticos pelos elementos de sua própria marca do Office. Considere maneiras de seguir layouts comuns de suplemento ou padrões de design da experiência do usuário durante a inserção de controles e componentes que são familiares para seus clientes.

Inserir uma interface do usuário baseada em HTML com uma forte identidade visual no Office pode criar dissonâncias para os clientes. Encontre um equilíbrio que se ajuste perfeitamente ao Office, mas também se alinhe claramente à sua marca pai ou serviço. Quando um suplemento não se ajusta ao Office, normalmente é porque elementos estilísticos estão em conflito. Por exemplo, a tipografia é muito grande e está fora da grade, as cores são contrastantes ou particularmente fortes ou as animações são supérfluas e se comportam de maneira diferente do Office. A aparência e o comportamento de controles ou componentes se desviam demasiadamente dos padrões do Office.

## <a name="typography"></a>Tipografia
Segoe é o tipo de fonte padrão para o Office. Use-a no suplemento para alinhar objetos de conteúdo, caixas de diálogo e painéis de tarefas do Office. O Office UI Fabric lhe dá acesso à fonte Segoe. Ele fornece um conjunto completo da fonte Segoe com muitas variações (incluindo espessura e tamanho da fonte) em classes CSS convenientes. Nem todos os tamanhos e espessuras do Office UI Fabric terão boa aparência em um suplemento do Office. Para obter um ajuste harmonioso ou evitar conflitos, considere o uso de um subconjunto do conjunto de fontes do Fabric. Aqui está uma lista de classes base do Fabric que recomendamos para uso em suplementos do Office.

|Exemplo |Classe |Tamanho |Peso |Uso recomendado |
|------ |----- |---- |------ |----------------- |
|![Imagem de Texto Hero](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Essa classe é maior do que todos os outros elementos tipográficos no Office. Use-a com moderação para não prejudicar o ajuste na hierarquia visual.</li><li>Evite o uso de cadeias de caracteres longas em espaços restritos.</li><li>Deixe bastante espaço em branco ao redor do texto ao usar esta classe.</li><li>Comumente usada para mensagens da tela de apresentação, elementos Hero ou outras chamadas à ação.</li></ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Essa classe corresponde ao título do painel de tarefas dos aplicativos do Office.</li><li>Use-a com moderação para evitar uma hierarquia tipográfica monótona.</li><li>Comumente usado como o elemento de nível superior, como títulos de conteúdo, página ou caixa de diálogo.</li><li></ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Essa classe é a primeira abaixo de títulos.</li><li>Comumente usada como um subtítulo, um elemento de navegação ou um cabeçalho de grupo.</li><ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |*Comumente usada como corpo de texto dentro de suplementos. |
|![Imagem de Texto Hero](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |*Comumente usada para texto secundário ou terciário, como carimbos de data/hora, linhas, títulos ou rótulos de campo. |
|![Imagem de Texto Hero](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |*A menor etapa no painel de tipos deve ser usada raramente. Está disponível para situações em que a legibilidade não é necessária. |
> A cor do texto não está incluída nessas classes base. Use a opção "neutro principal" do Fabric para o texto na maioria das telas de fundo brancas.

## <a name="color"></a>Cor
A cor é geralmente usada para enfatizar a marca e reforçar a hierarquia visual. Ela ajuda a identificar uma interface, além de orientar os clientes em uma experiência. No Office, a cor é usada para os mesmos objetivos, mas é aplicada intencionalmente e de forma mínima. Ela nunca sobrecarrega o conteúdo do cliente. Mesmo quando cada aplicativo do Office é identificado com sua própria cor dominante, ela é usada com moderação.

O Office UI Fabric inclui um conjunto padrão de cores de tema. Quando o Fabric é aplicado a um suplemento do Office, como componentes ou em layouts, os mesmos objetivos são aplicados. A cor deve comunicar a hierarquia, levando intencionalmente os clientes à ação, sem interferir no conteúdo. As cores de tema do Fabric podem introduzir uma nova cor de ênfase para a interface geral. Esse novo elemento pode entrar em conflito com a identidade visual do aplicativo do Office e interferir na hierarquia. Em outras palavras, o Fabric pode introduzir uma nova cor de ênfase para a interface geral quando usado em um suplemento. Essa nova cor de ênfase pode desviar a atenção e interferir em toda a hierarquia. Considere maneiras de evitar conflitos e interferência. Use ênfase neutra ou substitua cores de tema do Fabric para corresponder à identidade visual do aplicativo do Office ou às cores de sua própria marca.

Os aplicativos do Office permitem que os clientes personalizem as interfaces aplicando um tema de interface do usuário do Office. Os clientes podem escolher entre quatro temas de interface do usuário para variar o estilo de telas de fundo e botões no Word, no PowerPoint, no Excel e em outros aplicativos do Office. Para que os suplementos pareçam uma parte natural do Office e reajam à personalização, use nossas APIs de Temas. Por exemplo, as cores de tela de fundo do painel de tarefas alternam para um cinza escuro em alguns temas. Nossas APIs de temas permitem que faça o mesmo e ajuste o texto de primeiro plano para garantir a [acessibilidade](add-in-design-guidelines.md#accessibility-guidelines).

>  Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](https://dev.office.com/docs/reference/shared/office.context.officetheme.htm) para combinar o tema dos aplicativos do Office. Atualmente, essa API só está disponível no Office 2016.

> Para suplementos de conteúdo do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](https://dev.office.com/docs/add-ins/powerpoint/use-document-themes-in-your-powerpoint-add-ins.htm).

Aplique as seguintes diretrizes gerais para as cores:

* Use as cores com moderação para comunicar a hierarquia e reforçar a marca.
* O uso exagerado de uma cor de realce única aplicada aos elementos interativos e não interativos pode causar confusão. Por exemplo, evite usar a mesma cor para itens selecionados e não selecionados em um menu de navegação.
* Evite conflitos desnecessários com cores de aplicativo da identidade visual do Office.
* Use as cores de sua própria marca para criar a associação com seu serviço ou empresa.
* Verifique se todo o texto é acessível. Verifique se há uma razão de contraste de 4.5:1 entre o texto de primeiro plano e a tela de fundo.
* Lembre-se do daltonismo. Use mais do que apenas a cor para indicar a interatividade e a hierarquia.
* Confira as [diretrizes de ícone](design-icons.md) para saber mais sobre a criação de ícones de comando do suplemento com a paleta de cores de ícone do Office.

## <a name="layout"></a>Layout
Cada contêiner HTML inserido no Office terá um layout. Esses layouts são das telas principais do suplemento. Nelas, você criará experiências que permitem que os clientes iniciem ações, modifiquem configurações, exibam, rolem ou naveguem pelo conteúdo. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experiência. Se você tiver um site existente com o qual ps clientes estão familiarizados, considere a reutilização de layouts de páginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em contêineres HTML do Office.

Para obter diretrizes de layout, confira [Painel de tarefas](task-pane-add-ins.md), [Conteúdo](content-add-ins.md) e [Caixa de diálogo](dialog-boxes.md). Para obter mais informações sobre como montar componentes do Office UI Fabric em layouts comuns e fluxos de experiência do usuário, confira [Modelos de padrões de design da experiência do usuário](ux-design-patterns.md).

Aplique as seguintes diretrizes gerais aos layouts:

*    Evite margens estreitas ou amplas em contêineres HTML. 20 pixels é um ótimo padrão. 
*    Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.
*    As interfaces do Office estão em uma grade de 4px. Procure manter o preenchimento entre os elementos como múltiplos de 4. 
*    Sobrecarregar a interface pode causar confusão e prejudicar a facilidade de uso com interações de toque. 
*    Mantenha layouts consistentes entre as telas. Alterações de layout inesperadas parecem bugs visuais que contribuem para a falta de confiança na solução. 
*    Siga os padrões de layout comuns. As convenções ajudam os usuários a compreender como usar uma interface.
*    Evite elementos redundantes como identidade visual ou comandos.
*    Consolide os controles e modos de exibição para evitar exigir muitos movimentos do mouse. 
*    Crie experiências ágeis que se adaptem a alturas e larguras de contêineres HTML.

## <a name="component-language"></a>Linguagem de componente

Telas e layouts são compostos de conteúdo e componentes. Os componentes são controles que ajudam os clientes a interagir com os elementos do software ou serviço. Botões, navegação, selos, alertas e menus suspensos são exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.

O Office UI Fabric renderiza componentes que têm aparência e comportamento como os de uma parte do Office. Tire proveito do Fabric para garantir a integração perfeita ao Office. Se o suplemento tiver sua própria linguagem de componente pré-existente, não será necessário descartá-lo para usar o Fabric. Procure oportunidades para mantê-lo durante a integração ao Office. Considere maneiras de trocar elementos estilísticos, remover conflitos ou adotar estilos e comportamentos que removam a confusão para o usuário.

Aplique as seguintes diretrizes gerais aos componentes:

*    Não replique a faixa de opções do Office no suplemento
*    Evite criar menus, botões ou outros componentes que se comportem de forma diferente de componentes do Office.
*    Use os componentes do [Office UI Fabric](office-ui-fabric.md) que recomendamos para suplementos.
*    Use os [modelos de padrões de design da experiência do usuário](ux-design-patterns.md) para componentes da interface do usuário do Office comuns. 

## <a name="icons"></a>Ícones
Ícones são a representação visual de um comportamento ou conceito. Eles são usados frequentemente para adicionar significado a controles e comandos. Os elementos visuais, realistas ou simbólicos, habilitam o usuário a navegar pela interface do usuário da mesma maneira como os avisos ajudam os usuários a navegar pelo ambiente. Eles devem ser simples e claros e conter apenas os detalhes necessários para permitir que os clientes analisem rapidamente a ação que ocorrerá quando eles escolherem um controle.

As interfaces de faixa de opções do Office têm um estilo visual padrão. Se você estiver criando um comando de suplemento para a faixa de opções do Office, siga nossas [diretrizes de ícone](design-icons). Isso garante a consistência e a familiaridade em aplicativos do Office. As diretrizes ajudarão você a criar um conjunto de ativos PNG para a solução que se ajustem como parte natural do Office.

Muitos contêineres HTML contêm controles com iconografia. Use a fonte personalizada do Office UI Fabric para renderizar os ícones com o estilo do Office no suplemento. A fonte de ícone do Fabric contém muitos glifos para metáforas comuns do Office que você pode dimensionar, colorir e estilizar para atender às suas necessidades. Se você tiver uma linguagem visual existente com seu próprio conjunto de ícones, fique à vontade para usá-la em telas HTML. Criar continuidade com sua própria marca com um conjunto de ícones padrão é uma parte importante de qualquer linguagem de design. Tenha cuidado para não criar confusão para os clientes entrando em conflito com as metáforas do Office.

Aplique as seguintes diretrizes gerais aos ícones:

* Não reutilize glifos do Office UI Fabric para comandos de suplemento na faixa de opções do Office ou em menus contextuais. Os ícones do Fabric são estilisticamente diferentes e não serão compatíveis.
* Use a linguagem de ícones do Office para representar comportamentos ou conceitos.
* Reutilize metáforas visuais comuns do Office, como o pincel para formatar ou a lupa para localizar.
* Não use indevidamente metáforas para ações não relacionadas. Usar o mesmo elemento visual para um comportamento ou conceito diferente pode causar confusão para os usuários.

## <a name="animation"></a>Animação
Componentes, controles e elementos da interface do usuário geralmente têm comportamentos interativos que exigem transições, movimento ou animação. Características comuns de movimento entre elementos da interface do usuário definem os aspectos de animação de uma linguagem de design. Como o Office é voltado para a produtividade, a linguagem de animação do Office dá suporte ao objetivo de ajudar os clientes a realizar tarefas. Ela permite o equilíbrio entre a resposta de alto desempenho, a coreografia confiável e a satisfação detalhada.

O Office UI Fabric inclui uma biblioteca de animação para controlar a animação em contêineres HTML. Use para ajustar perfeitamente no Office. Ele ajudará a criar experiências que são mais sentidas do que observadas. As classes CSS de animação fornecem direcionalidade, entrada/saída e especificações de duração que reforçam modelos mentais do Office e fornecem oportunidades para que os clientes saibam como interagir com o suplemento. 

Se o suplemento tem sua própria linguagem de animação, use-a. Procure oportunidades para manter sua animação de identidade visual durante a integração ao Office. Tenha cuidado para não interferir ou entrar em conflito com padrões de movimento comuns no Office. Evite criar experiências que sejam ornamentos que apenas desviam a atenção dos clientes.

Aplique as seguintes diretrizes gerais às animações:

* As animações devem ser sentidas e experimentadas de forma subconsciente, para não impedir a conclusão da tarefa.
* Evite antecipações, saltos, pulos ou outros efeitos que emulem as características físicas do mundo natural.
* Coreografe os elementos para reforçar a hierarquia e os modelos mentais.
* Use o movimento para orientar o usuário e fornecer foco composicional sobre os principais elementos para conclusão da tarefa. 
* Considere a origem do elemento de disparo. Use animação para criar um vínculo entre a ação e a interface de usuário resultante.
* Considere o tom e a finalidade do conteúdo ao escolher animações. Lide com mensagens críticas de forma diferente da navegação exploratória.
