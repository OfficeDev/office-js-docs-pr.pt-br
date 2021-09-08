---
title: Diretrizes de acessibilidade para suplementos do Office
description: Saiba como tornar seu Office de usuário acessível a todos os usuários.
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: 889563af8ab5f7bbcd4037eedb42933369a92cf2
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937671"
---
# <a name="accessibility-guidelines"></a>Diretrizes de acessibilidade

À medida que você projeta e desenvolve seus suplementos do Office, convém verificar se todos os usuários e clientes potenciais são capazes de usar seu suplemento com êxito. Aplique as seguintes diretrizes para garantir que sua solução seja acessível a todos os públicos.

## <a name="design-for-multiple-input-methods"></a>Projetar para vários métodos de entrada

- Certifique-se de que os usuários possam realizar operações usando apenas o teclado. Os usuários devem conseguir se mover para todos os elementos acionáveis da página usando uma combinação das teclas Tab e de setas.
- Em um dispositivo móvel, quando os usuários operam um controle por toque, o dispositivo deve fornecer um feedback sonoro útil.
- Forneça rótulos úteis para todos os controles interativos. 

## <a name="make-your-add-in-easy-to-use"></a>Tornar seu suplemento fácil de usar

- Não dependa de um único atributo, como cor, tamanho, forma, local, orientação ou som, para atribuir significados na sua interface do usuário.
- Evite alterações inesperadas de contexto, como mover o foco para outro elemento da interface do usuário sem uma ação do usuário.
- Ofereça uma maneira de verificar, confirmar ou reverter todas as ações de associação.
- Forneça uma maneira de pausar ou parar mídias, como áudio e vídeo.
- Não estabeleça um limite de tempo para uma ação do usuário.

## <a name="make-your-add-in-easy-to-see"></a>Deixar seu suplemento fácil de ver

- Evite mudanças de cor inesperadas.
- Forneça informações significativas e em tempo hábil para descrever elementos de interface do usuário, títulos e cabeçalhos, entradas e erros. Verifique se os nomes dos controles descrevem adequadamente o objetivo do controle.
- Siga as [diretrizes padrão](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) de contraste de cor.

## <a name="account-for-assistive-technologies"></a>Incluir tecnologias adaptativas

- Evite usar recursos que interfiram em tecnologias adaptativas, incluindo em interações visuais, auditivas ou outras.
- Não forneça o texto em um formato de imagem. Os leitores de tela não podem ler o texto em imagens.
- Forneça uma maneira para os usuários ajustarem ou desativarem todas as fontes de áudio.
- Forneça uma maneira para os usuários ativarem legendas ou descrições de áudio com fontes de áudio.
- Forneça alternativas para o som como um meio para alertar os usuários, como indicações visuais ou vibrações.

## <a name="see-also"></a>Confira também

- [Diretrizes de Acessibilidade para Conteúdo da Web (WCAG) 2.0](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Orientações sobre a Aplicação das WCAG 2.0 para Tecnologias de Comunicação e Informação que não Sejam da Web (WCAG2ICT)](https://www.w3.org/TR/wcag2ict/)
- [Padrão Europeu para requisitos de acessibilidade para Tecnologias de Comunicação e Informação (ICT)](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
