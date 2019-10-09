---
title: Diretrizes de cor para Suplementos do Office
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3930cf22d40bd853c3fd6d96ade77a1a060cfc9d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448172"
---
# <a name="color"></a>Cor

A cor é geralmente usada para enfatizar a marca e reforçar a hierarquia visual. Ela ajuda a identificar uma interface, além de orientar os clientes em uma experiência. No Office, a cor é usada para os mesmos objetivos, mas é aplicada intencionalmente e de forma mínima. Ela nunca sobrecarrega o conteúdo do cliente. Mesmo quando cada aplicativo do Office é identificado com sua própria cor dominante, ela é usada com moderação.

O Office UI Fabric inclui um conjunto padrão de cores de tema. Quando o Fabric é aplicado a um suplemento do Office, como componentes ou em layouts, os mesmos objetivos são aplicados. A cor deve comunicar a hierarquia, levando intencionalmente os clientes à ação, sem interferir no conteúdo. As cores de tema do Fabric podem introduzir uma nova cor de ênfase para a interface geral. Esse novo elemento pode entrar em conflito com a identidade visual do aplicativo do Office e interferir na hierarquia. Em outras palavras, o Fabric pode introduzir uma nova cor de ênfase para a interface geral quando usado em um suplemento. Essa nova cor de ênfase pode desviar a atenção e interferir em toda a hierarquia. Considere maneiras de evitar conflitos e interferência. Use ênfase neutra ou substitua cores de tema do Fabric para corresponder à identidade visual do aplicativo do Office ou às cores de sua própria marca.

Os aplicativos do Office permitem que os clientes personalizem as interfaces aplicando um tema de interface do usuário do Office. Os clientes podem escolher entre quatro temas de interface do usuário para variar o estilo de telas de fundo e botões no Word, no PowerPoint, no Excel e em outros aplicativos do Office. Para que os suplementos pareçam uma parte natural do Office e reajam à personalização, use nossas APIs de Temas. Por exemplo, as cores de tela de fundo do painel de tarefas alternam para um cinza escuro em alguns temas. Nossas APIs de temas permitem que faça o mesmo e ajuste o texto de primeiro plano para garantir a [acessibilidade](../design/accessibility-guidelines.md).

> [!NOTE]
> - Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](/javascript/api/office/office.context) para combinar o tema dos aplicativos do Office. Essa API está atualmente disponível no Office 2016 ou posterior.
> - Para suplementos de conteúdo do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Aplique as seguintes diretrizes gerais para as cores:

* Use as cores com moderação para comunicar a hierarquia e reforçar a marca.
* O uso exagerado de uma cor de realce única aplicada aos elementos interativos e não interativos pode causar confusão. Por exemplo, evite usar a mesma cor para itens selecionados e não selecionados em um menu de navegação.
* Evite conflitos desnecessários com cores de aplicativo da identidade visual do Office.
* Use as cores de sua própria marca para criar a associação com seu serviço ou empresa.
* Verifique se todo o texto é acessível. Certifique-se de que haja uma taxa de contraste de 4,5:1 entre o texto do primeiro plano e o plano de fundo.
* Esteja ciente da cegueira de cores. Use mais do que apenas cor para indicar interatividade e hierarquia.
* Consulte as [diretrizes de ícone](../design/add-in-icons.md) para saber mais sobre a criação de ícones de comando de suplemento com a cor de ícone do Office paleta.
