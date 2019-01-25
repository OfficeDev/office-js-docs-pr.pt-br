---
title: Crie o design dos seus suplementos do Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 70faca768f5af70baf389c16fe8259427a85e8d9
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388714"
---
# <a name="design-your-office-add-ins"></a>Crie o design dos seus suplementos do Office

Os suplementos ampliam a experiência do Office fornecendo funcionalidades contextuais que os usuários podem acessar nos clientes Office. Os suplementos capacitam os usuários a produzir mais, permitindo o acesso a funcionalidades de terceiros dentro do Office, sem mudanças de contexto dispendiosas. 

O design da experiência de usuário do suplemento deve se integrar perfeitamente ao Office para proporcionar uma interação eficiente e natural aos usuários. Aproveite os [comandos de suplemento](add-in-commands.md) para fornecer acesso ao seu suplemento e aplique nossas práticas recomendadas ao criar sua interface do usuário personalizada baseada em HTML.

## <a name="office-design-principles"></a>Princípios de design do Office

Os aplicativos do Office seguem um conjunto geral de diretrizes de interação. Esses apps compartilham conteúdo e possuem elementos com aparência e comportamento semelhantes. Essa convergência baseia-se em um conjunto de princípios de design. Eles ajudam a equipe do Office a criar interfaces que dão suporte a tarefas dos clientes. Ao compreender e seguir esses princípios, você poderá oferecer suporte às metas de seus clientes dentro do Office.

Siga os princípios de design do Office para criar experiências positivas com os suplementos:

- **Crie designs explicitamente para o Office.** A funcionalidade e aparência de um suplemento deve complementar harmoniosamente a experiência no Office. Os suplementos devem parecer nativos. Eles têm que se encaixar diretamente no Word em um iPad ou no PowerPoint Online. Um suplemento bem projetado será uma combinação adequada da sua experiência, da plataforma e do aplicativo do Office. Considere a possibilidade de usar o Office UI Fabric como sua linguagem de design. Aplique temas ao documento e à interface do usuário quando apropriado.

- **Concentre-se em algumas tarefas importantes e realize-as de forma adequada** Ajude os clientes a concluir um trabalho sem interferir em outros trabalhos. Forneça um valor real aos clientes. Concentre-se em casos de uso comuns e escolha com cuidado aqueles que mais beneficiem os usuários durante a interação com documentos do Office.

- **Favoreça mais o conteúdo do que a aparência**. Permita que a página, o slide ou a planilha dos clientes permaneça como o foco da experiência. Um suplemento é uma interface auxiliar. Nenhum elemento supérfluo de interface do usuário deve interferir no conteúdo e nas funcionalidades do suplemento. Crie uma identidade visual para sua experiência de maneira sensata. Sabemos que é importante fornecer aos usuários uma experiência reconhecível e exclusiva, mas evite distrações. Tente manter o foco no conteúdo e na conclusão de tarefas, não na marca.

- **Torne-o interessante e mantenha os usuários no controle.** As pessoas gostam de usar produtos funcionais e visualmente atraentes. Crie sua experiência com atenção. Leve em conta cada interação e detalhe visual para acertar em todos os elementos. Permita que os usuários controlem a experiência. As etapas necessárias para concluir uma tarefa devem ser claras e relevantes. Decisões importantes devem ser fáceis de entender. As ações devem ser revertidas com facilidade. Um suplemento não é um destino, mas sim uma melhoria à funcionalidade do Office.

- **Design para todas as plataformas e métodos de entrada**. Os suplementos são projetados para funcionar em todas as plataformas com suporte do Office, portanto, a Experiência de Usuário do suplemento deve ser otimizada para funcionar em plataformas e fatores forma. Dê suporte a mouse/teclado e dispositivos de entrada por toque e verifique se a interface do usuário HTML personalizada responde na adaptação aos diversos fatores forma. Para saber mais, confira o tópico [Otimizar para toque](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Confira também
- [Office UI Fabric](https://developer.microsoft.com/pt-BR/fabric) 
- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)

