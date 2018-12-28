---
title: Diretrizes de cor para Suplementos do Office
description: ''
ms.date: 06/27/2018
ms.openlocfilehash: 34e067e4f5361ca54b8e50d6b86ff42d31154f19
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458087"
---
# <a name="color"></a>Cor
A cor é geralmente usada para enfatizar a marca e reforçar a hierarquia visual. Ela ajuda a identificar uma interface, além de orientar os clientes em uma experiência. No Office, a cor é usada para os mesmos objetivos, mas é aplicada intencionalmente e de forma mínima. Ela nunca sobrecarrega o conteúdo do cliente. Mesmo quando cada aplicativo do Office é identificado com sua própria cor dominante, ela é usada com moderação.

O Office UI Fabric inclui um conjunto padrão de cores de tema. Quando o Fabric é aplicado a um suplemento do Office, como componentes ou em layouts, os mesmos objetivos são aplicados. A cor deve comunicar a hierarquia, levando intencionalmente os clientes à ação, sem interferir no conteúdo. As cores de tema do Fabric podem introduzir uma nova cor de ênfase para a interface geral. Esse novo elemento pode entrar em conflito com a identidade visual do aplicativo do Office e interferir na hierarquia. Em outras palavras, o Fabric pode introduzir uma nova cor de ênfase para a interface geral quando usado em um suplemento. Essa nova cor de ênfase pode desviar a atenção e interferir em toda a hierarquia. Considere maneiras de evitar conflitos e interferência. Use ênfase neutra ou substitua cores de tema do Fabric para corresponder à identidade visual do aplicativo do Office ou às cores de sua própria marca.

Os aplicativos do Office permitem que os clientes personalizem as interfaces aplicando um tema de interface do usuário do Office. Os clientes podem escolher entre quatro temas de interface do usuário para variar o estilo de telas de fundo e botões no Word, no PowerPoint, no Excel e em outros aplicativos do Office. Para que os suplementos pareçam uma parte natural do Office e reajam à personalização, use nossas APIs de Temas. Por exemplo, as cores de tela de fundo do painel de tarefas alternam para um cinza escuro em alguns temas. Nossas APIs de temas permitem que faça o mesmo e ajuste o texto de primeiro plano para garantir a [acessibilidade](../design/accessibility-guidelines.md).

> [!NOTE]
> - Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](https://docs.microsoft.com/javascript/api/office/office.context) para combinar o tema dos aplicativos do Office. Atualmente, essa API só está disponível no Office 2016.
> - Para suplementos de conteúdo do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Aplique as seguintes diretrizes gerais para as cores:

* Use as cores com moderação para comunicar a hierarquia e reforçar a marca.
* O uso exagerado de uma cor de realce única aplicada aos elementos interativos e não interativos pode causar confusão. Por exemplo, evite usar a mesma cor para itens selecionados e não selecionados em um menu de navegação.
* Evite conflitos desnecessários com cores de aplicativo da identidade visual do Office.
* Use as cores de sua própria marca para criar a associação com seu serviço ou empresa.
* Verifique se todo o texto é acessível. Verifique se há uma razão de contraste de 4.5:1 entre o texto de primeiro plano e a tela de fundo.
* Lembre-se do daltonismo, use mais do que apenas cores para indicar interatividade e hierarquia.
* Consulte as [diretrizes de ícone](../design/add-in-icons.md) para saber mais sobre a criação de ícones de comando do suplemento com a paleta de cores de ícones do Office.