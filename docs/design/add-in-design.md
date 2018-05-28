---
title: Crie o design dos seus suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 80e349c677a3727f2867a0780a202277f3a6a0d9
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="design-your-office-add-ins"></a>Crie o design dos seus suplementos do Office

Os suplementos ampliam a experi?ncia do Office fornecendo funcionalidades contextuais que os usu?rios podem acessar nos clientes Office. Os suplementos capacitam os usu?rios a produzir mais, permitindo o acesso a funcionalidades de terceiros dentro do Office, sem mudan?as de contexto dispendiosas. 

O design da experi?ncia de usu?rio do suplemento deve se integrar perfeitamente ao Office para proporcionar uma intera??o eficiente e natural aos usu?rios. Aproveite os [comandos de suplemento](add-in-commands.md) para fornecer acesso ao seu suplemento e aplique nossas pr?ticas recomendadas ao criar sua interface do usu?rio personalizada baseada em HTML.

## <a name="office-design-principles"></a>Princ?pios de design do Office

Os aplicativos do Office seguem um conjunto geral de diretrizes de intera??o. Esses apps compartilham conte?do e possuem elementos com apar?ncia e comportamento semelhantes. Essa converg?ncia baseia-se em um conjunto de princ?pios de design. Eles ajudam a equipe do Office a criar interfaces que d?o suporte a tarefas dos clientes. Ao compreender e seguir esses princ?pios, voc? poder? oferecer suporte ?s metas de seus clientes dentro do Office.

Siga os princ?pios de design do Office para criar experi?ncias positivas com os suplementos:

- **Crie designs explicitamente para o Office.** A funcionalidade e apar?ncia de um suplemento deve complementar harmoniosamente a experi?ncia no Office. Os suplementos devem parecer nativos. Eles t?m que se encaixar diretamente no Word em um iPad ou no PowerPoint Online. Um suplemento bem projetado ser? uma combina??o adequada da sua experi?ncia, da plataforma e do aplicativo do Office. Considere a possibilidade de usar o Office UI Fabric como sua linguagem de design. Aplique temas ao documento e ? interface do usu?rio quando apropriado.

- **Concentre-se em algumas tarefas importantes e realize-as de forma adequada** Ajude os clientes a concluir um trabalho sem interferir em outros trabalhos. Forne?a um valor real aos clientes. Concentre-se em casos de uso comuns e escolha com cuidado aqueles que mais beneficiem os usu?rios durante a intera??o com documentos do Office.

- **Favore?a mais o conte?do do que a apar?ncia**. Permita que a p?gina, o slide ou a planilha dos clientes permane?a como o foco da experi?ncia. Um suplemento ? uma interface auxiliar. Nenhum elemento sup?rfluo de interface do usu?rio deve interferir no conte?do e nas funcionalidades do suplemento. Crie uma identidade visual para sua experi?ncia de maneira sensata. Sabemos que ? importante fornecer aos usu?rios uma experi?ncia reconhec?vel e exclusiva, mas evite distra??es. Tente manter o foco no conte?do e na conclus?o de tarefas, n?o na marca.

- **Torne isso agrad?vel e mantenha os usu?rios no controle.** As pessoas gostam de usar produtos funcionais e visualmente atraentes. Crie sua experi?ncia com cuidado. Acerte bem os pormenores, considerando cada intera??o e os detalhes visuais. Permita que os usu?rios controlem suas experi?ncias. As etapas necess?rias para concluir uma tarefa devem ser claras e relevantes. Decis?es importantes devem ser f?ceis de entender. As a??es devem ser facilmente revers?veis. Um suplemento n?o ? um destino, ? um aprimoramento da funcionalidade do Office.

- **Design para todas as plataformas e m?todos de entrada**. Os suplementos s?o projetados para funcionar em todas as plataformas com suporte do Office, portanto, a Experi?ncia de Usu?rio do suplemento deve ser otimizada para funcionar em plataformas e fatores forma. D? suporte a mouse/teclado e dispositivos de entrada por toque e verifique se a interface do usu?rio HTML personalizada responde na adapta??o aos diversos fatores forma. Para saber mais, confira o t?pico [Otimizar para toque](../concepts/add-in-development-best-practices.md#optimize-for-touch). 

## <a name="see-also"></a>Veja tamb?m
- [Office UI Fabric](https://dev.office.com/fabric) 
- [Pr?ticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)

