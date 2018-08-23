---
title: Crie o design dos seus suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 80e349c677a3727f2867a0780a202277f3a6a0d9
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437399"
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

- **Torne isso agradável e mantenha os usuários no controle.** As pessoas gostam de usar produtos funcionais e visualmente atraentes. Crie sua experiência com cuidado. Acerte bem os pormenores, considerando cada interação e os detalhes visuais. Permita que os usuários controlem suas experiências. As etapas necessárias para concluir uma tarefa devem ser claras e relevantes. Decisões importantes devem ser fáceis de entender. As ações devem ser facilmente reversíveis. Um suplemento não é um destino, é um aprimoramento da funcionalidade do Office.

- **Design para todas as plataformas e métodos de entrada**. Os suplementos são projetados para funcionar em todas as plataformas com suporte do Office, portanto, a Experiência de Usuário do suplemento deve ser otimizada para funcionar em plataformas e fatores forma. Dê suporte a mouse/teclado e dispositivos de entrada por toque e verifique se a interface do usuário HTML personalizada responde na adaptação aos diversos fatores forma. Para saber mais, confira o tópico [Otimizar para toque](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Veja também
- [Office UI Fabric](https://dev.office.com/fabric) 
- [Práticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)

