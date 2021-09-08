---
title: Objetos window que não têm suporte em Office Desindados
description: Este artigo especifica alguns dos objetos do tempo de execução da janela que não funcionam em Office de complementos.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937695"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Objetos window que não têm suporte em Office Desindados

Para algumas versões Windows e Office, os complementos são executados em um tempo de execução do Internet Explorer 11. (Para obter detalhes, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Algumas propriedades ou subpropropriedades do objeto global `window` não são suportadas no Internet Explorer 11. Essas propriedades são desabilitadas em complementos para garantir que o seu complemento fornece uma experiência consistente para todos os usuários, independentemente do navegador que o add-in está usando. Isso também ajuda o AngularJS a carregar corretamente.

Veja a seguir uma lista das propriedades desabilitadas. A lista é um trabalho em andamento. Se você descobrir propriedades adicionais que não funcionam em `window` complementos, use a ferramenta de comentários abaixo para nos dizer.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>Confira também

- [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md)