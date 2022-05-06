---
title: Projete a IU dos suplementos do Office
description: Conhecer as práticas recomendadas para o design visual de Suplementos do Office.
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: efbb0ee5f0ba75170b8bd4343392c07d9eda8501
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244748"
---
# <a name="design-the-ui-of-office-add-ins"></a>Projete a IU dos suplementos do Office

Os Suplementos do Office estendem a experiência do Office fornecendo funcionalidade contextual que os usuários podem acessar nos clientes do Office. Os suplementos capacitam os usuários a fazer mais, permitindo que eles acessem funcionalidades externas no Office, sem trocas de contexto dispendiosas.

O design da experiência de usuário do suplemento deve se integrar perfeitamente ao Office para proporcionar uma interação eficiente e natural aos usuários. Aproveite os [comandos de suplemento](add-in-commands.md) para fornecer acesso ao seu suplemento e aplique nossas práticas recomendadas ao criar sua interface do usuário personalizada baseada em HTML.

## <a name="office-design-principles"></a>Princípios de design do Office

Os aplicativos do Office seguem um conjunto geral de diretrizes de interação. Os aplicativos compartilham conteúdo e possuem elementos com aparência e comportamento semelhantes. Essa convergência baseia-se em um conjunto de princípios de design. Eles ajudam a equipe do Office a criar interfaces que oferecem suporte às tarefas dos clientes. Ao compreender e segui esses princípios você poderá oferecer suporte às metas de seus clientes dentro do Office.

Siga os princípios de design do Office para criar experiências positivas com os suplementos.

- **Design explicitamente para o Office.** A funcionalidade, bem como a aparência de um suplemento, deve complementar a experiência do Office. Os suplementos devem parecer nativos. Eles devem se ajustar perfeitamente ao Word em um iPad ou PowerPoint na Web. Um suplemento bem projetado será uma combinação apropriada de sua experiência, da plataforma e do aplicativo do Office. Aplique temas de documento e interface do usuário quando apropriado. Considere usar [IU do Fluent para a web](https://developer.microsoft.com/fluentui#/get-started/web) como seu idioma de design e conjunto de ferramentas. A interface do usuário do Fluent para a Web tem dois tipos.

  - **Para IUs não React:** Use **Fabric Core**, uma coleção de código aberto de classes CSS e mixins de SASS que fornecem acesso a cores, animações, fontes, ícones e grades. (É chamado de "Fabric Core" em vez de "Fluent Core" por motivos históricos.) Para começar, consulte [Fabric Core em Suplementos do Office](fabric-core.md).
  - **Para IUs React:** use o **Fluent UI React**, uma estrutura de front-end do React projetada para criar experiências que se encaixam perfeitamente em uma ampla gama de produtos da Microsoft. Ele fornece componentes robustos, atualizados e acessíveis baseados no React que são altamente personalizáveis usando o CSS-in-JS. Para começar, consulte [Suplementos da Fluent UI React no Office](using-office-ui-fabric-react.md).

- **Favoreça mais o conteúdo do que a aparência.** Permita que a página, o slide ou a planilha dos clientes permaneça como o foco da experiência. Um suplemento é uma interface auxiliar. Nenhum elemento supérfluo de interface do usuário deve interferir no conteúdo e nas funcionalidades do suplemento. Crie uma identidade visual para sua experiência de maneira sensata. Sabemos que é importante fornecer aos usuários uma experiência reconhecível e exclusiva, mas evite distrações. Tente manter o foco no conteúdo e na conclusão de tarefas, não na marca.

- **Torne-o interessante e mantenha os usuários no controle.** As pessoas gostam de usar produtos funcionais e visualmente atraentes. Crie sua experiência com atenção. Leve em conta cada interação e detalhe visual para acertar em todos os elementos. Permita que os usuários controlem a experiência. As etapas necessárias para concluir uma tarefa devem ser claras e relevantes. Decisões importantes devem ser fáceis de entender. As ações devem ser revertidas com facilidade. Um suplemento não é um destino, mas sim uma melhoria à funcionalidade do Office.

- **Design para todas as plataformas e métodos de entrada**. Os suplementos são projetados para funcionar em todas as plataformas com suporte do Office, portanto, a Experiência de Usuário do suplemento deve ser otimizada para funcionar em plataformas e fatores forma. Dê suporte a mouse/teclado e dispositivos de entrada por toque e verifique se a interface do usuário HTML personalizada responde na adaptação aos diversos fatores forma. Para saber mais, confira o tópico [Otimizar para toque](../concepts/add-in-development-best-practices.md#optimize-for-touch).

## <a name="see-also"></a>Confira também

- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
