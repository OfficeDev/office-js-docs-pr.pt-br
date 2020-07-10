---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093753"
---
# <a name="task-panes-in-office-add-ins"></a>Painéis de tarefas nos Suplementos do Office
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figura 1. Layout típico do painel de tarefa*

![Imagem exibindo um layout típico do painel de tarefas](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:-----|:--------|
|<ul><li>Inclua o nome do seu suplemento no título.</li></ul>|<ul><li>Não adicione o nome da sua empresa ao título.</li></ul>|
|<ul><li>Use nomes descritivos curtos no título.</li></ul>|<ul><li>Não acrescente cadeias de caracteres, como "suplemento", "para Word" ou "para Office", ao título do seu suplemento.</li></ul>|
|<ul><li>Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.</li></ul>||
|<ul><li>Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.</li></ul>||


## <a name="variants"></a>Variantes

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*

![Imagem exibindo os tamanhos de painel de tarefas da área de trabalho em 1366 x 768](../images/office-2016-taskpane-sizes.png)

- Excel – 320 x 455
- PowerPoint – 320 x 531
- Word – 320 x 531
- Outlook – 348 x 535

<br/>

*Figura 3. Tamanhos de painel de tarefas do Office*

![Imagem exibindo os tamanhos de painel de tarefas da área de trabalho em 1366 x 768](../images/office-365-taskpane-sizes.png)

- Excel – 350 x 378
- PowerPoint – 348 x 391
- Word – 329 x 445
- Outlook (na Web) - 320x570

## <a name="personality-menu"></a>Menu de personalidade

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 4. Menu de personalidade no Windows*

![Imagem mostrando o menu do personalidade na área de trabalho do Windows](../images/personality-menu-win.png)

No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.

*Figura 5. Menu de personalidade no Mac*

![Imagem mostrando o menu de personalidade na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub. 


## <a name="see-also"></a>Confira também

- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md) 
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)

