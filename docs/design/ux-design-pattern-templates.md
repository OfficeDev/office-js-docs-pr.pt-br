---
title: Padrões de design da experiência do usuário para suplementos do Office
description: ''
ms.date: 06/27/2018
ms.openlocfilehash: 635fc27d18a2c671dd1ac5a521c9d0a920c154ed
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432470"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Padrões de design da experiência do usuário para suplementos do Office

O design da experiência do usuário para os suplementos do Office deve fornecer uma experiência atraente para os usuários do Office e estender a experiência geral do Office, ajustando-se perfeitamente à interface do usuário padrão do Office.  

Nossos padrões de experiência do usuário são compostos de componentes. Os componentes são controles que ajudam os clientes a interagir com os elementos do software ou serviço. Botões, navegação e menus são exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.

O Office UI Fabric renderiza componentes que têm aparência e comportamento como os de uma parte do Office. Aproveite o Fabric para se integrar facilmente ao Office. Se o suplemento tiver sua própria linguagem de componente pré-existente, não será necessário descartá-lo para usar o Fabric. Procure oportunidades para mantê-lo durante a integração ao Office. Considere maneiras de trocar elementos estilísticos, remover conflitos ou adotar estilos e comportamentos que removam a confusão para o usuário.

Os padrões fornecidos são soluções de práticas recomendadas com base em cenários comuns de clientes e pesquisa de experiência do usuário. Eles servem para fornecer um ponto de entrada rápido para projetar e desenvolver suplementos, bem como orientação para alcançar o equilíbrio entre os elementos da Microsoft e da marca. Proporcionar uma experiência de usuário limpa e moderna que equilibre elementos de design da linguagem de design do Microsoft Fabric e a identidade de marca exclusiva do parceiro pode ajudar a aumentar a retenção de usuários e a adoção do seu suplemento.

Use os modelos padrão de experiência do usuário para:

* Aplicar soluções a cenários comuns de clientes.
* Aplicar as práticas recomendadas de design.
* Incorporar componentes e estilos do [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started).
* Criar suplementos que se integram visualmente à interface do usuário padrão do Office.
* Idealizar e visualizar a experiência do usuário.


## <a name="getting-started"></a>Introdução

Os padrões são organizados por ações principais ou experiências comuns em um suplemento. Os principais grupos são:

* [Tela de apresentação (FRE)](../design/first-run-experience-patterns.md)
* [Autenticação](../design/authentication-patterns.md)
* [Navegação](../design/navigation-patterns.md)
* [Design de identidade Visual](../design/branding-patterns.md)

Navegar por cada agrupamento para ter uma ideia de como você pode projetar o suplemento usando as práticas recomendadas.



>Observação: As telas de exemplo mostradas durante esta documentação estão projetadas e exibidas na resolução de **1366 x 768**




## <a name="see-also"></a>Confira também
* [Kits de ferramentas de design](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [Práticas recomendadas para o desenvolvimento de Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Introdução ao uso do Fabric React](https://docs.microsoft.com/office/dev/add-ins/design/using-office-ui-fabric-react)
