---
title: Idioma de design de suplemento do Office
description: Saiba como tornar o suplemento do Office visualmente compatível com o Office.
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: 0a1d175401ebaabe9c17cae18d196bc6461ba57c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718591"
---
# <a name="office-add-in-design-language"></a>Idioma de design de suplemento do Office

A linguagem de design do Office é um sistema visual claro e simples que garante a consistência nas experiências. Ela contém um conjunto de elementos visuais que definem as interfaces do Office, incluindo:

- Um tipo de fonte padrão
- Uma paleta de cores comuns
- Um conjunto de pesos e tamanhos tipográficos
- Diretrizes de ícones
- Ativos de ícones compartilhados
- Definições de animação
- Componentes comuns

O [Office UI Fabric](https://developer.microsoft.com/fabric) é a estrutura de front-end oficial para criação com a linguagem de design do Office. O uso do Fabric é opcional, mas é a maneira mais rápida de garantir que os suplementos sejam como uma extensão natural do Office. Tire proveito do Fabric para projetar e criar suplementos que complementam o Office.

Vários suplementos do Office estão associados a uma marca pré-existente. Você pode manter uma marca forte e sua linguagem visual ou de componente no suplemento. Procure oportunidades para manter sua própria linguagem visual durante a integração ao Office. Considere maneiras de substituir cores, tipografia, ícones ou outros elementos estilísticos pelos elementos de sua própria marca do Office. Considere maneiras de seguir layouts comuns de suplemento ou padrões de design da experiência do usuário durante a inserção de controles e componentes que são familiares para seus clientes.

Inserir uma interface do usuário baseada em HTML com uma forte identidade visual no Office pode criar dissonâncias para os clientes. Encontre um equilíbrio que se ajuste perfeitamente ao Office, mas também se alinhe claramente à sua marca pai ou serviço. Quando um suplemento não se ajusta ao Office, normalmente é porque elementos estilísticos estão em conflito. Por exemplo, a tipografia é muito grande e está fora da grade, as cores são contrastantes ou particularmente fortes ou as animações são supérfluas e se comportam de maneira diferente do Office. A aparência e o comportamento de controles ou componentes se desviam demasiadamente dos padrões do Office.
