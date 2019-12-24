---
title: 'Desenvolver Suplementos do Office '
description: Uma introdução ao desenvolvimento de Suplementos do Office.
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 731226883e2bdea4b68d0720042010a0f0117098
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851677"
---
# <a name="develop-office-add-ins"></a>Desenvolver Suplementos do Office 

> [!TIP]
> Examine [Criação de Suplementos do Office](../overview/office-add-ins-fundamentals.md) antes de ler este artigo.

Todos os Suplementos do Office são criados com base na plataforma de Suplementos do Office. Eles compartilham uma estrutura comum por meio da qual certas funcionalidades podem ser implementadas. Para qualquer suplemento que você crie, você precisará entender conceitos importantes como a disponibilidade do host e da plataforma, os padrões de programação da API do Office JavaScript, como especificar as configurações e os recursos do suplemento no arquivo de manifesto e muito mais. Os principais conceitos de desenvolvimento, como estes mencionados acima, são abordados aqui na seção **Conceitos básicos** > **Desenvolver** dessa documentação. Releia tal documentação antes de explorar a documentação específica do host que corresponde ao suplemento que você está criando (por exemplo, [Excel](../excel/index.md)).

> [!NOTE]
> A seção **Conceitos básicos** > **Desenvolver** > **Como** desta documentação contém artigos voltados para tarefas ou conceitos específicos de desenvolvimento. Por exemplo, você encontrará informações sobre tarefas como [desenvolvendo suplementos com o código do Visual Studio](develop-add-ins-vscode.md), [abrir automaticamente um painel de tarefas com um documento](automatically-open-a-task-pane-with-a-document.md), [criar comandos de suplemento](create-addin-commands.md)e [abrir uma caixa de diálogo](dialog-api-in-office-add-ins.md).

## <a name="next-steps"></a>Próximas etapas

Depois de se familiarizar com os conceitos básicos abordados aqui, explore a documentação específica do host que corresponde ao suplemento que você está criando (por exemplo, [Excel](../excel/index.md)). Cada seção específica do host da documentação contém informações específicas sobre a criação de suplementos para um determinado host do Office.

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Criando Suplementos do Office ](../overview/office-add-ins-fundamentals.md)
- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publicar Suplementos do Office](../publish/publish.md)