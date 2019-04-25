---
title: Diretrizes de layout para Suplementos do Office
description: ''
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: 9570bf041cf1df70ab95af656decb3c458c0d480
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32447165"
---
# <a name="layout"></a>Layout
Cada contêiner HTML inserido no Office terá um layout. Esses layouts são das telas principais do suplemento. Nelas, você criará experiências que permitem que os clientes iniciem ações, modifiquem configurações, exibam, rolem ou naveguem pelo conteúdo. Projeta o suplemento com layouts consistentes nas telas para garantir a continuidade da experiência. Se você tiver um site existente com o qual ps clientes estão familiarizados, considere a reutilização de layouts de páginas da Web existentes. Adapte-as para se ajustar de forma harmoniosa em contêineres HTML do Office.

Para obter diretrizes de layout, confira [Painel de tarefas](task-pane-add-ins.md), [Conteúdo](content-add-ins.md) e [Caixa de diálogo](dialog-boxes.md). Para obter mais informações sobre como montar componentes do Office UI Fabric em layouts comuns e fluxos de experiência do usuário, confira [Modelos de padrões de design da experiência do usuário](ux-design-pattern-templates.md).

Aplique as seguintes diretrizes gerais aos layouts:

*   Evite margens estreitas ou amplas em contêineres HTML. 20 pixels é um ótimo padrão.
*   Alinhe os elementos intencionalmente. Recuos extras e novos pontos de alinhamento devem auxiliar na hierarquia visual.
*   As interfaces do Office estão em uma grade de 4px. Procure manter o preenchimento entre os elementos como múltiplos de 4.
*   Sobrecarregar a interface pode causar confusão e prejudicar a facilidade de uso com interações de toque.
*   Mantenha layouts consistentes entre as telas. Alterações de layout inesperadas parecem bugs visuais que contribuem para a falta de confiança na solução.
*   Siga os padrões de layout comuns. As convenções ajudam os usuários a compreender como usar uma interface.
*   Evite elementos redundantes como identidade visual ou comandos.
*   Consolide os controles e modos de exibição para evitar exigir muitos movimentos do mouse.
*   Crie experiências ágeis que se adaptem a alturas e larguras de contêineres HTML.
