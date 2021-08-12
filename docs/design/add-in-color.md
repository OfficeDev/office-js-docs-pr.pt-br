---
title: Diretrizes de cor para Suplementos do Office
description: Saiba como usar cores na interface do usuário de um Office Add-in.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: fc22a2168a531d0f3fe50358f5d45e6052bfde6c3418f9ee13197bd48ed35101
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082727"
---
# <a name="color-guidelines-for-office-add-ins"></a>Diretrizes de cor para Suplementos do Office

A cor é geralmente usada para enfatizar a marca e reforçar a hierarquia visual. Ela ajuda a identificar uma interface, além de orientar os clientes em uma experiência. No Office, a cor é usada para os mesmos objetivos, mas é aplicada intencionalmente e de forma mínima. Ela nunca sobrecarrega o conteúdo do cliente. Mesmo quando cada aplicativo do Office é identificado com sua própria cor dominante, ela é usada com moderação.

![Diagrama mostrando o esquema de cores para Office, Excel, Word e PowerPoint. As cores principais para Office são preto e branco, e as cores secundárias são cinza claro, cinza escuro e laranja. A cor dominante para Excel é verde, Word é azul e PowerPoint laranja.](../images/office-addins-color-schemes.png)

[O Fabric Core](fabric-core.md) inclui um conjunto de cores de tema padrão. Quando o Fabric Core é aplicado a um Office de componentes ou em layouts, as mesmas metas se aplicam. A cor deve comunicar a hierarquia, levando intencionalmente os clientes à ação, sem interferir no conteúdo. As cores de tema do Fabric Core podem introduzir uma nova cor de destaque à interface geral. Esse novo elemento pode entrar em conflito com a identidade visual do aplicativo do Office e interferir na hierarquia. Em outras palavras, o Fabric Core pode introduzir uma nova cor de destaque à interface geral quando usado dentro de um complemento. Essa nova cor de ênfase pode desviar a atenção e interferir em toda a hierarquia. Considere maneiras de evitar conflitos e interferência. Use ênfases neutras ou sobrescreva cores de tema do Fabric Core para corresponder Aplicativo do Office identidade visual ou suas próprias cores de marca.

Os aplicativos do Office permitem que os clientes personalizem as interfaces aplicando um tema de interface do usuário do Office. Os clientes podem escolher entre quatro temas de interface do usuário para variar o estilo de telas de fundo e botões no Word, no PowerPoint, no Excel e em outros aplicativos do Office. Para fazer com que os seus complementos se sintam como uma parte natural da Office e respondam à personalização, use nossas APIs de Temas. Por exemplo, as cores de tela de fundo do painel de tarefas alternam para um cinza escuro em alguns temas. Nossas APIs de temas permitem que faça o mesmo e ajuste o texto de primeiro plano para garantir a [acessibilidade](../design/accessibility-guidelines.md).

> [!NOTE]
>
> - Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](/javascript/api/office/office.context) para combinar o tema dos aplicativos do Office. Esta API está disponível atualmente no Office 2016 ou posterior.
> - Para suplementos de conteúdo do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Aplique as seguintes diretrizes gerais para cor.

- Use as cores com moderação para comunicar a hierarquia e reforçar a marca.
- O uso exagerado de uma cor de realce única aplicada aos elementos interativos e não interativos pode causar confusão. Por exemplo, evite usar a mesma cor para itens selecionados e não selecionados em um menu de navegação.
- Evite conflitos desnecessários com cores de aplicativo da identidade visual do Office.
- Use as cores de sua própria marca para criar a associação com seu serviço ou empresa.
- Verifique se todo o texto é acessível. Certifique-se de que haja uma taxa de contraste de 4,5:1 entre o texto em primeiro plano e o plano de fundo.
- Lembre-se do daltonismo, use mais do que apenas cores para indicar interatividade e hierarquia.
- Consulte as [diretrizes de ícone para](../design/add-in-icons.md) saber mais sobre como projetar ícones de comando do add-in com o palete de cores Office ícone.
