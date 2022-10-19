---
title: Diretrizes para implantar suplementos do Office em nuvens governamentais
description: Saiba como implantar seu Suplemento do Office para ambientes de nuvem seguros e governamentais
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: f3995c62a1b7fb482df6a15da870f747f55e9508
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607587"
---
# <a name="guidance-for-deploying-office-add-ins-on-government-clouds"></a>Diretrizes para implantar suplementos do Office em nuvens governamentais

A Microsoft fornece as opções de nuvem governamental para nossos clientes sensíveis à privacidade em organizações governamentais locais, governamentais e nacionais. Isso oferece aos parceiros oportunidades para direcionar clientes críticos com seus Suplementos do Office. Devido à natureza mais restrita desses ambientes, que é importante para as necessidades de privacidade e segurança dos clientes, nem todos os recursos que normalmente estão disponíveis em um ambiente de produção padrão estão disponíveis nessas nuvens.

Para parceiros que fornecem seus Suplementos do Office aos clientes nesses ambientes de nuvem restritos, há diferenças importantes do ambiente de nuvem pública padrão que devem ser consideradas. As informações a seguir fornecem os detalhes que exigem tratamento especial por desenvolvedores que escrevem Suplementos do Office destinados aos clientes nesses ambientes.

## <a name="all-sovereign-environments"></a>Todos os ambientes soberanos

Para todos os ambientes de nuvem governamental (ou seja, Nuvem Soberana), a Office Store pública não está disponível. Isso significa que os usuários finais não podem adquirir Suplementos do Office diretamente da loja pública. Os administradores também não podem implantar suplementos do Office diretamente do repositório público em seus Administração Portal. Em vez disso, você deve trabalhar com administradores para garantir o seguinte:

- Os recursos e serviços necessários para sua solução estão disponíveis dentro do limite de nuvem. Você trabalha com os administradores de locatários para provisionar seu serviço e recursos dentro do limite de nuvem ou trabalha com o administrador de rede para habilitar o acesso aos recursos que residem fora do limite de nuvem.

- Os recursos que o Suplemento do Office acessa estão em conformidade com os requisitos da nuvem governamental da qual eles estão sendo acessados. Eles também devem estar em conformidade com quaisquer requisitos adicionais do locatário do cliente para o qual a solução está sendo provisionada. Esses requisitos incluem a transferência, o gerenciamento e o armazenamento de dados potencialmente confidenciais, bem como procedimentos de verificação de acesso e segurança de código e recursos.

- O manifesto do Suplemento do Office que descreve a solução e seu local de origem conforme aplicável à implantação de nuvem governamental específica é obtido do parceiro e ingerido para implantação para os usuários apropriados por meio do portal Administração.

## <a name="us-government-community-cloud-gcc"></a>Nuvem da Comunidade governamental dos EUA (GCC)

Além dos requisitos aplicáveis a todas as Nuvens Soberanas, cada ambiente de Nuvem Soberana tem suas próprias considerações que podem afetar os Suplementos do Office direcionados ao ambiente. O GCC é o menos restritivo dos ambientes de nuvem do governo e alguns dos recursos exigidos pelo suplemento estão disponíveis em seus pontos de extremidade de produção usuais nesse ambiente. Um desses recursos é a biblioteca de API JavaScript do Office. Os parceiros de solução podem continuar referenciando o recurso Office.js público como fazem com sua solução de produção pública.

## <a name="gcc-high-gcch-us-department-of-defense-dod-or-other-sovereign-managed-clouds"></a>GCC High (GCCH), DOD (Departamento de Defesa dos EUA) ou outras nuvens gerenciadas soberanas

Essas nuvens governamentais permanecem conectadas à Internet, mas o conjunto de pontos de extremidade públicos disponibilizados é severamente restrito. Um desses pontos de extremidade restritos é o ponto de extremidade público para carregar a biblioteca de API JavaScript do Office. O local de CDN público para Office.js não estará acessível de dentro desses ambientes. No entanto, haverá uma CDN interna por nuvem do Microsoft Office provisionada com o mesmo recurso. Isso significa que a URL do ponto de extremidade Office.js será diferente e seu Suplemento do Office pode precisar de algum nível de personalização para ser executado. Considerando as restrições adicionais, é provável que qualquer solução fornecida aos clientes também exija serviços de provedor de hospedagem dentro do ambiente, exigindo personalizações adicionais. Você precisará determinar a melhor maneira de fornecer sua solução aos clientes, de modo que ela esteja em conformidade com as restrições adicionais impostas aos serviços em execução dentro dos limites desses ambientes.

## <a name="airgapped-sovereign-clouds"></a>Nuvens soberanas airgapped

Essas nuvens governamentais estão essencialmente desconectadas da Internet pública inteiramente. Qualquer recurso que normalmente seria acessado de recursos públicos deve ser provisionado de forma personalizada dentro desses ambientes de nuvem. Nas nuvens GCCH e DOD mencionadas anteriormente, a maioria dos provedores de soluções (se não todos) precisará provisionar seus serviços e recursos dentro da nuvem. Há uma opção para fazer exceções de firewall que permitem o acesso a serviços e recursos públicos. No entanto, esse bypass não é possível em nuvens airgapped. Assim como nas nuvens GCCH e DOD, haverá uma CDN do Microsoft Office provisionada dentro de cada ambiente de nuvem que hospeda os recursos necessários, como a Office.js biblioteca. Você precisará trabalhar em conjunto com os administradores de locatários do cliente para determinar a melhor maneira de fornecer seus serviços e recursos de maneira que esteja em conformidade com os requisitos estritos de acesso para nuvens soberanas aéreas.
