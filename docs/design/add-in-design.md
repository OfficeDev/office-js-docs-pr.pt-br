---
title: Projete a IU dos suplementos do Office
description: Conhecer as práticas recomendadas para o design visual de Suplementos do Office.
ms.date: 05/12/2021
localization_priority: Priority
ms.openlocfilehash: 7b5314a07e15c5d57b4e5c27e781ebba5c1a3492
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330049"
---
# <a name="design-the-ui-of-office-add-ins"></a>Projete a IU dos suplementos do Office

Os suplementos ampliam a experiência do Office fornecendo funcionalidades contextuais que os usuários podem acessar nos clientes Office. Os suplementos capacitam os usuários a produzir mais, permitindo o acesso a funcionalidades de terceiros dentro do Office, sem mudanças de contexto dispendiosas.

Seu design de IU adicional deve se integrar perfeitamente ao Office para fornecer uma interação eficiente e natural para seus usuários. Aproveite as vantagens dos [comandos de suplemento](add-in-commands.md) para fornecer acesso ao seu suplemento e aplique as práticas recomendadas que recomendamos ao criar uma IU personalizada baseada em HTML.

## <a name="office-design-principles"></a>Princípios de design do Office

Os aplicativos do Office seguem um conjunto geral de diretrizes de interação. Os aplicativos compartilham conteúdo e possuem elementos com aparência e comportamento semelhantes. Essa convergência baseia-se em um conjunto de princípios de design. Eles ajudam a equipe do Office a criar interfaces que oferecem suporte às tarefas dos clientes. Ao compreender e segui esses princípios você poderá oferecer suporte às metas de seus clientes dentro do Office.

Siga os princípios de design do Office para criar experiências positivas com os suplementos:

- **Crie designs explicitamente para o Office.** A funcionalidade, bem como a aparência de um suplemento, devem complementar harmoniosamente a experiência do Office. Os suplementos devem parecer nativos. Eles devem se encaixar perfeitamente no Word em um iPad ou PowerPoint na web. Um suplemento bem projetado será uma combinação adequada da sua experiência, a plataforma e o aplicativo do Office. Aplique temas ao documento e à interface do usuário quando apropriado. Considere o uso da [IU do Fluent para a web](https://developer.microsoft.com/fluentui#/get-started/web) como sua linguagem de design e conjunto de ferramentas. A IU do Fluent para a web tem dois tipos:

  - **Para IUs não React:** Use **Fabric Core**, uma coleção de código aberto de classes CSS e mixins de SASS que fornecem acesso a cores, animações, fontes, ícones e grades. (É chamado de "Fabric Core" em vez de "Fluent Core" por motivos históricos.) Para começar, consulte [Fabric Core em Suplementos do Office](fabric-core.md).
  - **Para IUs React:** use o **Fluent UI React**, uma estrutura de front-end do React projetada para criar experiências que se encaixam perfeitamente em uma ampla gama de produtos da Microsoft. Ele fornece componentes robustos, atualizados e acessíveis baseados no React que são altamente personalizáveis usando o CSS-in-JS. Para começar, consulte [Suplementos da Fluent UI React no Office](using-office-ui-fabric-react.md).

- **Favoreça mais o conteúdo do que a aparência**. Permita que a página, o slide ou a planilha do cliente permaneçam no foco da experiência. Um suplemento é uma interface auxiliar. Nenhum elemento supérfluo de interface do usuário deve interferir no conteúdo e nas funcionalidades do suplemento. Crie uma identidade visual para sua experiência de maneira sensata. Sabemos que é importante fornecer aos usuários uma experiência exclusiva e reconhecível, mas evite distrações. Tente manter o foco no conteúdo e na conclusão de tarefas, não na marca.

- **Torne-o interessante e mantenha os usuários no controle.** As pessoas gostam de usar produtos funcionais e visualmente atraentes. Crie sua experiência com atenção. Leve em conta cada interação e detalhe visual para acertar em todos os elementos. Permita que os usuários controlem a experiência. As etapas necessárias para concluir uma tarefa devem ser claras e relevantes. Decisões importantes devem ser fáceis de entender. As ações devem ser revertidas com facilidade. Um suplemento não é um destino, mas sim uma melhoria à funcionalidade do Office.

- **Design para todas as plataformas e métodos de entrada**. Os suplementos são projetados para funcionar em todas as plataformas com suporte do Office, portanto, a Experiência de Usuário do suplemento deve ser otimizada para funcionar em plataformas e fatores forma. Dê suporte a mouse/teclado e dispositivos de entrada por toque e verifique se a interface do usuário HTML personalizada responde na adaptação aos diversos fatores forma. Para saber mais, confira o tópico [Otimizar para toque](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Confira também

- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
