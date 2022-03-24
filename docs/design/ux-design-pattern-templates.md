---
title: Padrões de design da experiência do usuário para suplementos do Office
description: Obter uma visão geral dos padrões de design da interface do usuário para Office de Office, incluindo padrões de navegação, autenticação, primeira executar e identidade visual.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9349096ecb03e4fa7980b9430f3940b5d09bdab2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743116"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Padrões de design da experiência do usuário para suplementos do Office

O design da experiência do usuário para os suplementos do Office deve fornecer uma experiência atraente para os usuários do Office e estender a experiência geral do Office, ajustando-se perfeitamente à interface do usuário padrão do Office.  

Nossos padrões de experiência do usuário são compostos de componentes. Os componentes são controles que ajudam os clientes a interagir com os elementos do software ou serviço. Botões, navegação e menus são exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.

[Fluent componentes da interface](using-office-ui-fabric-react.md) do usuário React parecem e se comportam como parte do Office, assim como os componentes neutros da [estrutura do Office UI Fabric JS](fabric-core.md). Aproveite qualquer conjunto de componentes para se integrar com Office. Como alternativa, se o seu complemento tiver seu próprio idioma de componente preexistência, você não precisará descartar. Procure oportunidades para mantê-lo durante a integração ao Office. Considere maneiras de trocar elementos estilísticos, remover conflitos ou adotar estilos e comportamentos que removam a confusão para o usuário.

Os padrões fornecidos são soluções de práticas recomendadas com base em cenários comuns de clientes e pesquisa de experiência do usuário. Eles devem fornecer um ponto de entrada rápido para projetar e desenvolver os complementos, bem como orientações para alcançar o equilíbrio entre os elementos de marca da Microsoft e seus próprios. Fornecer uma experiência de usuário limpa e moderna que equilibra elementos de design da linguagem de design da interface do usuário Fluent da Microsoft e a identidade de marca exclusiva do parceiro pode ajudar a aumentar a retenção e a adoção do usuário do seu complemento.

Use os modelos padrão de experiência do usuário para:

* Aplicar soluções a cenários comuns de clientes.
* Aplicar as práticas recomendadas de design.
* [Incorpore Fluent e estilos da interface do](https://developer.microsoft.com/fluentui#/get-started) usuário.
* Criar suplementos que se integram visualmente à interface do usuário padrão do Office.
* Idealizar e visualizar a experiência do usuário.

## <a name="getting-started"></a>Introdução

Os padrões são organizados por ações principais ou experiências comuns em um suplemento. Os principais grupos são:

* [Tela de apresentação (FRE)](../design/first-run-experience-patterns.md)
* [Autenticação](../design/authentication-patterns.md)
* [Navegação](../design/navigation-patterns.md)
* [Design de identidade Visual](../design/branding-patterns.md)

Navegar por cada agrupamento para ter uma ideia de como você pode projetar o suplemento usando as práticas recomendadas.

> [!NOTE]
> As telas de exemplo mostradas ao longo desta documentação, estão projetadas e exibidas na resolução de **1366x768**.

## <a name="see-also"></a>Confira também

* [Kits de ferramentas de design](design-toolkits.md)
* [Interface do usuário do Fluent](https://developer.microsoft.com/fluentui#)
* [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
* [Interface do usuário Fluent React em Suplementos do Office](using-office-ui-fabric-react.md)
