---
title: Diretrizes de layout para Suplementos do Office
description: Obter diretrizes sobre como layout de um painel de tarefas ou caixa de diálogo em um Office Add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c004171333ee3ded54a36efe227fcd1714935a59
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349361"
---
# <a name="layout"></a>Layout

Cada contêiner HTML inserido no Office terá um layout. Esses layouts são das telas principais do suplemento. Nelas, você criará experiências que permitem que os clientes iniciem ações, modifiquem configurações, exibam, rolem ou naveguem pelo conteúdo. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experiência. Se você tiver um site existente com o qual ps clientes estão familiarizados, considere a reutilização de layouts de páginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em contêineres HTML do Office.

Para obter diretrizes de layout, confira [Painel de tarefas](task-pane-add-ins.md), [Conteúdo](content-add-ins.md) e [Caixa de diálogo](dialog-boxes.md). Para obter mais informações sobre como montar Fluent interface do usuário React [,](using-office-ui-fabric-react.md)ou [Office UI Fabric JS](fabric-core.md), [componentes](ux-design-pattern-templates.md)em layouts comuns e fluxos de experiência do usuário, consulte modelos de padrões de design deux .

Aplique as seguintes diretrizes gerais para layouts.

- Evite margens estreitas ou amplas em contêineres HTML. 20 pixels é um ótimo padrão.
- Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.
- As interfaces do Office estão em uma grade de 4px. Procure manter o preenchimento entre os elementos como múltiplos de 4.
- Sobrecarregar a interface pode causar confusão e prejudicar a facilidade de uso com interações de toque.
- Mantenha layouts consistentes entre as telas. Alterações de layout inesperadas parecem bugs visuais que contribuem para a falta de confiança na solução.
- Siga os padrões de layout comuns. As convenções ajudam os usuários a compreender como usar uma interface.
- Evite elementos redundantes como identidade visual ou comandos.
- Consolide os controles e modos de exibição para evitar exigir muitos movimentos do mouse.
- Crie experiências ágeis que se adaptem a alturas e larguras de contêineres HTML.
