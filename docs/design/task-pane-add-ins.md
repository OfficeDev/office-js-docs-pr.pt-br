---
title: Painéis de tarefas nos Suplementos do Office
description: Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: d911101a7df1f1ad8aa01b8e0006bd93d994a193
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092914"
---
# <a name="task-panes-in-office-add-ins"></a>Painéis de tarefas nos Suplementos do Office

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*Figura 1. Layout típico do painel de tarefa*

![Ilustração exibindo um layout típico do painel de tarefas com guias de seção na parte superior, logotipo da empresa e nome da empresa na parte inferior esquerda e um ícone de configurações na parte inferior direita.](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:-----|:--------|
|Inclua o nome do seu suplemento no título.|Não adicione o nome da sua empresa ao título.|
|Use nomes descritivos curtos no título.|Não acrescente cadeias de caracteres como "suplemento", "para Word" ou "para o Office" ao título do suplemento.|
|Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.|*Nenhum.*|
|Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento, a menos que seu suplemento seja voltado para uso no Outlook.|*Nenhum.*|

## <a name="variants"></a>Variantes

As imagens a seguir mostram os vários tamanhos do painel de tarefas com a faixa de opções do aplicativo do Office em uma resolução de 1366 x 768. No Excel, é necessário um espaço vertical adicional para acomodar a barra de fórmulas.  

*Figura 2. Tamanhos de painel de tarefas da área de trabalho do Office 2016*

![Diagrama exibindo os tamanhos do painel de tarefas da área de trabalho na resolução 1366x768.](../images/office-2016-taskpane-sizes.png)

- Excel – 320 x 455 pixels
- PowerPoint – 320 x 531 pixels
- Word – 320 x 531 pixels
- Outlook – 348 x 535 pixels

<br/>

*Figura 3. Tamanhos do painel de tarefas do Office*

![Diagrama exibindo os tamanhos do painel de tarefas na resolução 1366x768.](../images/office-365-taskpane-sizes.png)

- Excel – 350 x 378 pixels
- PowerPoint – 348 x 391 pixels
- Word – 329 x 445 pixels
- Outlook (na Web) – 320 x 570 pixels

## <a name="personality-menu"></a>Menu de personalidade

Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac. (Não há suporte para o menu de personalidade no Outlook.)

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 4. Menu de personalidade no Windows*

![Diagrama mostrando o menu de personalidade na área de trabalho do Windows.](../images/personality-menu-win.png)

No Mac, no menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço para 34 x 32 pixels, como mostrado.

*Figura 5. Menu de personalidade no Mac*

![Diagrama mostrando o menu de personalidade na área de trabalho do Mac.](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver uma amostra que implementa um painel de tarefas, confira [Suplemento do Excel JS Tendências de Despesas do WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) no GitHub.

## <a name="see-also"></a>Confira também

- [Núcleo da Malha em Suplementos do Office](fabric-core.md)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
