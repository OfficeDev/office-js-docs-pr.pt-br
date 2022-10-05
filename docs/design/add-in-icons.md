---
title: Diretrizes de ícone para suplementos do Office
description: Obtenha uma visão geral de como projetar ícones e os estilos de design Fresh e Monoline para comandos de suplemento.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 523c1341f84b09a428cfb7d6d7a3a933d4632604
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68466947"
---
# <a name="icons"></a>Ícones

Ícones são a representação visual de um comportamento ou conceito. Eles são usados frequentemente para adicionar significado a controles e comandos. Os elementos visuais, realistas ou simbólicos, permitem ao usuário navegar pela interface do usuário da mesma maneira que os avisos os ajudam a navegar pelo ambiente. Eles devem ser simples, claros e conter apenas os detalhes necessários para permitir que os clientes analisem rapidamente a ação que ocorrerá ao escolherem um controle.

As interfaces da faixa de opções do aplicativo do Office têm um estilo visual padrão. Isso garante a consistência e a familiaridade em aplicativos do Office. As diretrizes ajudarão você a criar um conjunto de ativos PNG para a solução que se ajustem como parte natural do Office.

Muitos contêineres HTML contêm controles com iconografia. Use a fonte personalizada do Fabric Core para renderizar ícones com estilo do Office em seu suplemento. A fonte do ícone fornecida pelo [Fabric Core](fabric-core.md) contém muitos glifos para metáforas comuns do Office que você pode dimensionar, cor e estilo para atender às suas necessidades. Se você tiver uma linguagem visual existente com seu próprio conjunto de ícones, fique à vontade para usá-la em telas HTML. Criar continuidade com sua própria marca com um conjunto de ícones padrão é uma parte importante de qualquer linguagem de design. Tenha cuidado para não criar confusão para os clientes entrando em conflito com as metáforas do Office.

## <a name="design-icons-for-add-in-commands"></a>Desenvolver ícones para comandos de suplemento

Os [Comandos de suplementos](add-in-commands.md) adicionam botões, texto e ícones à interface do usuário do Office. Os botões de comando de suplemento devem fornecer ícones significativos e rótulos que identifiquem claramente a ação que o usuário está realizando ao usar um comando. Os artigos a seguir fornecem diretrizes estilísticos e de produção para ajudá-lo a projetar ícones que se integram perfeitamente ao Office.

- Para o estilo Monoline do Microsoft 365, confira as diretrizes do ícone de estilo [Monoline para Suplementos do Office](add-in-icons-monoline.md).
- Para o estilo Fresh do Office 2013 ou superior perpétuo, confira as diretrizes do ícone estilo Fresh [para Suplementos do Office](add-in-icons-fresh.md).

> [!NOTE]
> Você deve escolher um estilo ou outro, e seu suplemento usará os mesmos ícones se ele estiver em execução no Microsoft 365 ou no Office perpétuo.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
