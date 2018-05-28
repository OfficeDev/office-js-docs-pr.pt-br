---
title: Idioma de design de suplemento do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7d19714fa14fb374bcd41aa744c08929c228c94f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-design-language"></a>Idioma de design de suplemento do Office

A linguagem de design do Office ? um sistema visual claro e simples que garante a consist?ncia nas experi?ncias. Ela cont?m um conjunto de elementos visuais que definem as interfaces do Office, incluindo:

- Um tipo de fonte padr?o
- Uma paleta de cores comuns
- Um conjunto de pesos e tamanhos tipogr?ficos
- Diretrizes de ?cones
- Ativos de ?cones compartilhados
- Defini??es de anima??o
- Componentes comuns

O [Office UI Fabric](https://dev.office.com/fabric) ? a estrutura de front-end oficial para cria??o com a linguagem de design do Office. O uso do Fabric ? opcional, mas ? a maneira mais r?pida de garantir que os suplementos sejam como uma extens?o natural do Office. Tire proveito do Fabric para projetar e criar suplementos que complementam o Office.

V?rios suplementos do Office est?o associados a uma marca pr?-existente. Voc? pode manter uma marca forte e sua linguagem visual ou de componente no suplemento. Procure oportunidades para manter sua pr?pria linguagem visual durante a integra??o ao Office. Considere maneiras de substituir cores, tipografia, ?cones ou outros elementos estil?sticos pelos elementos de sua pr?pria marca do Office. Considere maneiras de seguir layouts comuns de suplemento ou padr?es de design da experi?ncia do usu?rio durante a inser??o de controles e componentes que s?o familiares para seus clientes.

Inserir uma interface do usu?rio baseada em HTML com uma forte identidade visual no Office pode criar disson?ncias para os clientes. Encontre um equil?brio que se ajuste perfeitamente ao Office, mas tamb?m se alinhe claramente ? sua marca pai ou servi?o. Quando um suplemento n?o se ajusta ao Office, normalmente ? porque elementos estil?sticos est?o em conflito. Por exemplo, a tipografia ? muito grande e est? fora da grade, as cores s?o contrastantes ou particularmente fortes ou as anima??es s?o sup?rfluas e se comportam de maneira diferente do Office. A apar?ncia e o comportamento de controles ou componentes se desviam demasiadamente dos padr?es do Office.

## <a name="typography"></a>Tipografia
Segoe ? o tipo de fonte padr?o para o Office. Use-a no suplemento para alinhar objetos de conte?do, caixas de di?logo e pain?is de tarefas do Office. O Office UI Fabric lhe d? acesso ? fonte Segoe. Ele fornece um conjunto completo da fonte Segoe com muitas varia??es (incluindo espessura e tamanho da fonte) em classes CSS convenientes. Nem todos os tamanhos e espessuras do Office UI Fabric ter?o boa apar?ncia em um suplemento do Office. Para obter um ajuste harmonioso ou evitar conflitos, considere o uso de um subconjunto do conjunto de fontes do Fabric. Aqui est? uma lista de classes base do Fabric que recomendamos para uso em suplementos do Office.

|Amostra |Classe |Tamanho |Peso |Uso recomendado |
|------ |----- |---- |------ |----------------- |
|![Imagem de Texto Hero](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Essa classe ? maior do que todos os outros elementos tipogr?ficos no Office. Use-a com modera??o para n?o prejudicar o ajuste na hierarquia visual.</li><li>Evite o uso de cadeias de caracteres longas em espa?os restritos.</li><li>Deixe bastante espa?o em branco ao redor do texto ao usar esta classe.</li><li>Comumente usada para mensagens da tela de apresenta??o, elementos Hero ou outras chamadas ? a??o.</li></ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Essa classe corresponde ao t?tulo do painel de tarefas dos aplicativos do Office.</li><li>Use-a com modera??o para evitar uma hierarquia tipogr?fica mon?tona.</li><li>Comumente usado como o elemento de n?vel superior, como t?tulos de conte?do, p?gina ou caixa de di?logo.</li></ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Essa classe ? a primeira abaixo de t?tulos.</li><li>Comumente usada como um subt?tulo, um elemento de navega??o ou um cabe?alho de grupo.</li><ul> |
|![Imagem de Texto Hero](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |<ul><li>Muito usada como corpo de texto dentro de suplementos.</li><ul>|
|![Imagem de texto Hero](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |<ul><li>Muito usada em texto secund?rio ou terci?rio, como carimbos de data/hora, linhas, t?tulos ou r?tulos de campo.</li><ul>|
|![Imagem de texto Hero](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |<ul><li>A menor etapa no painel de tipos deve ser usada raramente. Est? dispon?vel para situa??es em que a legibilidade n?o ? necess?ria.</li><ul>|

> [!NOTE]
> A cor do texto n?o est? inclu?da nessas classes base. Use a op??o "Neutro principal" do Fabric para a maioria dos textos em fundos brancos.

## <a name="color"></a>Cor
A cor ? frequentemente usada para enfatizar a marca e refor?ar a hierarquia visual. Ajuda a identificar uma interface, al?m de orientar os clientes por meio de uma experi?ncia. Dentro do Office, a cor ? usada para os mesmos objetivos, mas ? aplicada de forma proposital e m?nima. Em nenhum momento sobrecarrega o conte?do do cliente. Mesmo quando cada aplicativo do Office ? marcado com sua pr?pria cor dominante, ? usado com modera??o.

O Office UI Fabric inclui um conjunto padr?o de cores de tema. Quando o Fabric ? aplicado a um suplemento do Office, como componentes ou em layouts, os mesmos objetivos s?o aplicados. A cor deve comunicar a hierarquia, levando intencionalmente os clientes ? a??o, sem interferir no conte?do. As cores de tema do Fabric podem introduzir uma nova cor de ?nfase para a interface geral. Esse novo elemento pode entrar em conflito com a identidade visual do aplicativo do Office e interferir na hierarquia. Em outras palavras, o Fabric pode introduzir uma nova cor de ?nfase para a interface geral quando usado em um suplemento. Essa nova cor de ?nfase pode desviar a aten??o e interferir em toda a hierarquia. Considere maneiras de evitar conflitos e interfer?ncia. Use ?nfase neutra ou substitua cores de tema do Fabric para corresponder ? identidade visual do aplicativo do Office ou ?s cores de sua pr?pria marca.

Os aplicativos do Office permitem que os clientes personalizem as interfaces aplicando um tema de interface do usu?rio do Office. Os clientes podem escolher entre quatro temas de interface do usu?rio para variar o estilo de telas de fundo e bot?es no Word, no PowerPoint, no Excel e em outros aplicativos do Office. Para que os suplementos pare?am uma parte natural do Office e reajam ? personaliza??o, use nossas APIs de Temas. Por exemplo, as cores de tela de fundo do painel de tarefas alternam para um cinza escuro em alguns temas. Nossas APIs de temas permitem que fa?a o mesmo e ajuste o texto de primeiro plano para garantir a [acessibilidade](add-in-design-guidelines.md#accessibility-guidelines).

> [!NOTE]
> - Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](https://dev.office.com/reference/add-ins/shared/office.context.officetheme) para combinar o tema dos aplicativos do Office. Atualmente, essa API s? est? dispon?vel no Office 2016.
> - Para suplementos de conte?do do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Aplique as seguintes diretrizes gerais para as cores:

* Use as cores com modera??o para comunicar a hierarquia e refor?ar a marca.
* O uso exagerado de uma cor de realce ?nica aplicada aos elementos interativos e n?o interativos pode causar confus?o. Por exemplo, evite usar a mesma cor para itens selecionados e n?o selecionados em um menu de navega??o.
* Evite conflitos desnecess?rios com cores de aplicativo da identidade visual do Office.
* Use as cores de sua pr?pria marca para criar a associa??o com seu servi?o ou empresa.
* Verifique se todo o texto ? acess?vel. Verifique se h? uma raz?o de contraste de 4.5:1 entre o texto de primeiro plano e a tela de fundo.
* Lembre-se do daltonismo, use mais do que apenas cores para indicar interatividade e hierarquia.
* Confira as [diretrizes de ?cone](design-icons.md) para saber mais sobre a cria??o de ?cones de comando do suplemento com a paleta de cores de ?cone do Office.

## <a name="layout"></a>Layout
Cada cont?iner HTML inserido no Office ter? um layout. Esses layouts s?o das telas principais do suplemento. Nelas, voc? criar? experi?ncias que permitem que os clientes iniciem a??es, modifiquem configura??es, exibam, rolem ou naveguem pelo conte?do. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experi?ncia. Se voc? tiver um site existente com o qual ps clientes est?o familiarizados, considere a reutiliza??o de layouts de p?ginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em cont?ineres HTML do Office.

Para obter diretrizes de layout, confira [Painel de tarefas](task-pane-add-ins.md), [Conte?do](content-add-ins.md) e [Caixa de di?logo](dialog-boxes.md). Para obter mais informa??es sobre como montar componentes do Office UI Fabric em layouts comuns e fluxos de experi?ncia do usu?rio, confira [Modelos de padr?es de design da experi?ncia do usu?rio](ux-design-patterns.md).

Aplique as seguintes diretrizes gerais aos layouts:

*   Evite margens estreitas ou amplas em cont?ineres HTML. 20 pixels ? um ?timo padr?o.
*   Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.
*   As interfaces do Office est?o em uma grade de 4px. Procure manter o preenchimento entre os elementos como m?ltiplos de 4.
*   Sobrecarregar a interface pode causar confus?o e prejudicar a facilidade de uso com intera??es de toque.
*   Mantenha layouts consistentes entre as telas. Altera??es de layout inesperadas parecem bugs visuais que contribuem para a falta de confian?a na solu??o.
*   Siga os padr?es de layout comuns. As conven??es ajudam os usu?rios a compreender como usar uma interface.
*   Evite elementos redundantes como identidade visual ou comandos.
*   Consolide os controles e modos de exibi??o para evitar exigir muitos movimentos do mouse.
*   Crie experi?ncias ?geis que se adaptem a alturas e larguras de cont?ineres HTML.

## <a name="component-language"></a>Linguagem de componente

Telas e layouts s?o compostos de conte?do e componentes. Os componentes s?o controles que ajudam os clientes a interagir com os elementos do software ou servi?o. Bot?es, navega??o, selos, alertas e menus suspensos s?o exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.

O Office UI Fabric renderiza componentes que t?m apar?ncia e comportamento como os de uma parte do Office. Tire proveito do Fabric para garantir a integra??o perfeita ao Office. Se o suplemento tiver sua pr?pria linguagem de componente pr?-existente, n?o ser? necess?rio descart?-lo para usar o Fabric. Procure oportunidades para mant?-lo durante a integra??o ao Office. Considere maneiras de trocar elementos estil?sticos, remover conflitos ou adotar estilos e comportamentos que removam a confus?o para o usu?rio.

Aplique as seguintes diretrizes gerais aos componentes:

*   N?o replique a faixa de op??es do Office no suplemento
*   Evite criar menus, bot?es ou outros componentes que se comportem de forma diferente de componentes do Office.
*   Use os componentes do [Office UI Fabric](office-ui-fabric.md) que recomendamos para suplementos.
*   Use os [modelos de padr?es de design da experi?ncia do usu?rio](ux-design-patterns.md) para componentes da interface do usu?rio do Office comuns.

## <a name="icons"></a>?cones
?cones s?o a representa??o visual de um comportamento ou conceito. Eles s?o usados frequentemente para adicionar significado a controles e comandos. Os elementos visuais, realistas ou simb?licos, habilitam o usu?rio a navegar pela interface do usu?rio da mesma maneira como os avisos ajudam os usu?rios a navegar pelo ambiente. Eles devem ser simples e claros e conter apenas os detalhes necess?rios para permitir que os clientes analisem rapidamente a a??o que ocorrer? quando eles escolherem um controle.

As interfaces de faixa de op??es do Office t?m um estilo visual padr?o. Se voc? estiver criando um comando de suplemento para a faixa de op??es do Office, siga nossas [diretrizes de ?cone](design-icons.md). Isso garante a consist?ncia e a familiaridade em aplicativos do Office. As diretrizes ajudar?o voc? a criar um conjunto de ativos PNG para a solu??o que se ajustem como parte natural do Office.

Muitos cont?ineres HTML cont?m controles com iconografia. Use a fonte personalizada do Office UI Fabric para renderizar os ?cones com o estilo do Office no suplemento. A fonte de ?cone do Fabric cont?m muitos glifos para met?foras comuns do Office que voc? pode dimensionar, colorir e estilizar para atender ?s suas necessidades. Se voc? tiver uma linguagem visual existente com seu pr?prio conjunto de ?cones, fique ? vontade para us?-la em telas HTML. Criar continuidade com sua pr?pria marca com um conjunto de ?cones padr?o ? uma parte importante de qualquer linguagem de design. Tenha cuidado para n?o criar confus?o para os clientes entrando em conflito com as met?foras do Office.

Aplique as seguintes diretrizes gerais aos ?cones:

* N?o reutilize glifos do Office UI Fabric para comandos de suplemento na faixa de op??es do Office ou em menus contextuais. Os ?cones do Fabric s?o estilisticamente diferentes e n?o ser?o compat?veis.
* Use a linguagem de ?cones do Office para representar comportamentos ou conceitos.
* Reutilize met?foras visuais comuns do Office, como o pincel para formatar ou a lupa para localizar.
* N?o use indevidamente met?foras para a??es n?o relacionadas. Usar o mesmo elemento visual para um comportamento ou conceito diferente pode causar confus?o para os usu?rios.


## <a name="see-also"></a>Veja tamb?m

- [Diretrizes de design de suplementos do Office](add-in-design-guidelines.md)
- [Usar movimento em suplementos do Office](using-motion-office-addins.md)
