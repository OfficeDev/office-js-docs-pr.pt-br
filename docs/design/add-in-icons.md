---
title: Diretrizes de ícone para suplementos do Office
description: Obter uma visão geral de como projetar ícones e os estilos de design Fresh e Monoline para comandos de complemento.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 3073472332a31688676fba796dccd9920a49581d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937247"
---
# <a name="icons"></a>Ícones

Ícones são a representação visual de um comportamento ou conceito. Eles são usados frequentemente para adicionar significado a controles e comandos. Os elementos visuais, realistas ou simbólicos, permitem ao usuário navegar pela interface do usuário da mesma maneira que os avisos os ajudam a navegar pelo ambiente. Eles devem ser simples, claros e conter apenas os detalhes necessários para permitir que os clientes analisem rapidamente a ação que ocorrerá ao escolherem um controle.

Aplicativo do Office de faixa de opções têm um estilo visual padrão. Isso garante a consistência e a familiaridade em aplicativos do Office. As diretrizes ajudarão você a criar um conjunto de ativos PNG para sua solução que se ajuste como parte natural do Office.

Muitos contêineres HTML contêm controles com iconografia. Use a fonte personalizada do Fabric Core para renderizar Office ícones estilados no seu complemento. A fonte de ícone fornecida pelo [Fabric Core](fabric-core.md) contém muitos glifos para metáforas comuns Office que você pode dimensionar, cor e estilo para atender às suas necessidades. Se você tiver uma linguagem visual existente com seu próprio conjunto de ícones, fique à vontade para usá-la em telas HTML. Criar continuidade com sua própria marca com um conjunto de ícones padrão é uma parte importante de qualquer linguagem de design. Tenha cuidado para não criar confusão para os clientes entrando em conflito com as metáforas do Office.

## <a name="design-icons-for-add-in-commands"></a>Desenvolver ícones para comandos de suplemento

Os [Comandos de suplementos](add-in-commands.md) adicionam botões, texto e ícones à interface do usuário do Office. Os botões de comando de suplemento devem fornecer ícones significativos e rótulos que identifiquem claramente a ação que o usuário está realizando ao usar um comando. Os artigos a seguir fornecem diretrizes estilísticas e de produção para ajudá-lo a projetar ícones que se integram perfeitamente Office.

- Para o estilo monoline de Microsoft 365, consulte Diretrizes de ícone de estilo monoline para Office [Desempaco.](add-in-icons-monoline.md)
- Para o estilo Fresh de não assinatura Office 2013+, consulte Diretrizes de ícones de estilo novo para Office [Desem.](add-in-icons-fresh.md)

> [!NOTE]
> Você deve escolher um estilo ou o outro e o seu complemento usará os mesmos ícones se ele está sendo executado no Microsoft 365 ou não Office.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
