---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: bea2bf43d5d10a39e36cee679bf03790ca683126d915609e6a746176ff18da6c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081728"
---
# <a name="task-panes-in-office-add-ins"></a>Painéis de tarefas nos Suplementos do Office

Painéis de tarefas são superfícies de interface que normalmente são exibidas no lado direito da janela no Word, PowerPoint, Excel e Outlook. As painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefa quando não precisar inserir a funcionalidade diretamente no documento.

*Figura 1. Layout típico do painel de tarefa*

![Ilustração exibindo um layout típico do painel de tarefas com guias de seção na parte superior, logotipo da empresa e nome da empresa na parte inferior esquerda e um ícone de configurações na parte inferior direita.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:-----|:--------|
|<ul><li>Inclua o nome do seu suplemento no título.</li></ul>|<ul><li>Não adicione o nome da sua empresa ao título.</li></ul>|
|<ul><li>Use nomes descritivos curtos no título.</li></ul>|<ul><li>Não adicione cadeias de caracteres como "add-in", "for Word" ou "for Office" ao título do seu complemento.</li></ul>|
|<ul><li>Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</li></ul>||
|<ul><li>Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.</li></ul>||

## <a name="variants"></a>Variantes

As imagens a seguir mostram os vários tamanhos do painel de tarefas com Aplicativo do Office faixa de opções em uma resolução de 1366x768. No Excel, é necessário um espaço vertical adicional para acomodar a barra de fórmulas.  

*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*

![Diagrama que exibe os tamanhos do painel de tarefas da área de trabalho na resolução 1366x768.](../images/office-2016-taskpane-sizes.png)

- Excel - 320 x 455 pixels
- PowerPoint - 320 x 531 pixels
- Word - 320x531 pixels
- Outlook - 348 x 535 pixels

<br/>

*Figura 3. Office tamanhos do painel de tarefas*

![Diagrama exibindo os tamanhos do painel de tarefas na resolução 1366x768.](../images/office-365-taskpane-sizes.png)

- Excel - 350 x 378 pixels
- PowerPoint - 348 x 391 pixels
- Word - 329x445 pixels
- Outlook (na Web) - 320x570 pixels

## <a name="personality-menu"></a>Menu de personalidade

Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 4. Menu de personalidade no Windows*

![Diagrama mostrando o menu de personalidade na Windows desktop.](../images/personality-menu-win.png)

No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.

*Figura 5. Menu de personalidade no Mac*

![Diagrama mostrando o menu de personalidade na área de trabalho do Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub.

## <a name="see-also"></a>Confira também

- [Núcleo da Malha em Suplementos do Office](fabric-core.md)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
