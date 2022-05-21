---
title: Diretrizes de layout para Suplementos do Office
description: Obtenha diretrizes sobre como layout de um painel de tarefas ou caixa de diálogo em um Office Suplemento.
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 134e8c01a5a6057f84ef2f4f62c290a161e94cfa
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628072"
---
# <a name="layout"></a>Layout

Cada contêiner HTML inserido no Office terá um layout. Esses layouts são das telas principais do suplemento. Nelas, você criará experiências que permitem que os clientes iniciem ações, modifiquem configurações, exibam, rolem ou naveguem pelo conteúdo. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experiência. Se você tiver um site existente com o qual ps clientes estão familiarizados, considere a reutilização de layouts de páginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em contêineres HTML do Office.

Para obter diretrizes sobre layout, consulte [Painel de tarefas](task-pane-add-ins.md), [Conteúdo](content-add-ins.md). Para obter mais informações sobre como montar Fluent interface do usuário [React](using-office-ui-fabric-react.md) ou [Office UI Fabric JS](fabric-core.md), componentes em layouts comuns e [fluxos](ux-design-pattern-templates.md) de experiência do usuário, consulte modelos de padrões de design da experiência do usuário.

Aplique as diretrizes gerais a seguir para layouts.

- Evite margens estreitas ou amplas em contêineres HTML. 20 pixels é um ótimo padrão.
- Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.
- As interfaces do Office estão em uma grade de 4px. Procure manter o preenchimento entre os elementos como múltiplos de 4.
- Sobrecarregar a interface pode causar confusão e prejudicar a facilidade de uso com interações de toque.
- Mantenha layouts consistentes entre as telas. Alterações de layout inesperadas parecem bugs visuais que contribuem para a falta de confiança na solução.
- Siga os padrões de layout comuns. As convenções ajudam os usuários a compreender como usar uma interface.
- Evite elementos redundantes como identidade visual ou comandos.
- Consolide os controles e modos de exibição para evitar exigir muitos movimentos do mouse.
- Crie experiências ágeis que se adaptem a alturas e larguras de contêineres HTML.
