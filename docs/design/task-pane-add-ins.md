---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: ed3f3b8fdf7cf62b6016fe8b03393de0d56dfb33
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132015"
---
# <a name="task-panes-in-office-add-ins"></a>Painéis de tarefas nos Suplementos do Office

Painéis de tarefas são superfícies de interface que normalmente são exibidas no lado direito da janela no Word, PowerPoint, Excel e Outlook. As painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefa quando não precisar inserir a funcionalidade diretamente no documento.

*Figura 1. Layout típico do painel de tarefa*

![Ilustração exibindo um layout de painel de tarefas típico com guias de seção na parte superior, logotipo da empresa e nome da empresa na parte inferior esquerda e um ícone de configurações no canto inferior direito](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:-----|:--------|
|<ul><li>Inclua o nome do seu suplemento no título.</li></ul>|<ul><li>Não adicione o nome da sua empresa ao título.</li></ul>|
|<ul><li>Use nomes descritivos curtos no título.</li></ul>|<ul><li>Não acrescente cadeias de caracteres, como "suplemento", "para Word" ou "para Office", ao título do seu suplemento.</li></ul>|
|<ul><li>Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</li></ul>||
|<ul><li>Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.</li></ul>||

## <a name="variants"></a>Variantes

As imagens a seguir mostram os vários tamanhos de painel de tarefas com a faixa de opções do aplicativo do Office em uma resolução 1366x768. Para o Excel, é necessário espaço vertical adicional para acomodar a barra de fórmulas.  

*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*

![Diagrama exibindo os tamanhos de painel de tarefas da área de trabalho na Resolução 1366x768](../images/office-2016-taskpane-sizes.png)

- Excel-320 x 455 pixels
- PowerPoint-320 x 531 pixels
- Pixels do Word-320 x 531
- Pixels do Outlook-348 x 535

<br/>

*Figura 3. Tamanhos de painel de tarefas do Office*

![Diagrama exibindo os tamanhos de painel de tarefas na Resolução 1366x768](../images/office-365-taskpane-sizes.png)

- Excel-350 x 378 pixels
- PowerPoint-348 x 391 pixels
- Pixels do Word-329 x 445
- Outlook (na Web)-320 x 570 pixels

## <a name="personality-menu"></a>Menu de personalidade

Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 4. Menu de personalidade no Windows*

![Diagrama mostrando o menu de personalidade na área de trabalho do Windows](../images/personality-menu-win.png)

No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.

*Figura 5. Menu de personalidade no Mac*

![Diagrama mostrando o menu de personalidade na área de trabalho Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub.

## <a name="see-also"></a>Confira também

- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
