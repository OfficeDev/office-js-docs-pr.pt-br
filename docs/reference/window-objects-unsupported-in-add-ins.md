---
title: Objetos Window que não são compatíveis com suplementos do Office
description: Este artigo especifica alguns dos objetos de tempo de execução da janela que não funcionam em suplementos do Office.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160499"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Objetos Window que não são compatíveis com suplementos do Office

Para algumas versões do Windows e do Office, os suplementos são executados em um tempo de execução do Internet Explorer 11. (Para obter detalhes, consulte [navegadores usados por suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).) Algumas propriedades ou subpropriedades do `window` objeto global não são suportadas no Internet Explorer 11. Essas propriedades estão desabilitadas em suplementos para garantir que o suplemento forneça uma experiência consistente para todos os usuários, independentemente do navegador que o suplemento estiver usando. Isso também ajuda o AngularJS a carregar corretamente.

Veja a seguir uma lista das propriedades desabilitadas. A lista é um trabalho em andamento. Se você descobrir `window` Propriedades adicionais que não funcionam em suplementos, use a ferramenta de comentários abaixo para nos dizer.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>Confira também

- [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md)