---
title: Manter seu Suplemento do Office
description: Entenda nossos compromissos com a compatibilidade e como manter seu suplemento atualizado.
ms.date: 05/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7f70eab252af516ab8dda591668d48392ce9f04
ms.sourcegitcommit: e63d8e32b25a9987f4a39b92a342a82b37a3404c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/17/2022
ms.locfileid: "65432187"
---
# <a name="maintain-your-office-add-in"></a>Manter seu Suplemento do Office

Depois de publicar o suplemento, você deve mantê-lo atualizado com quaisquer alterações importantes das bibliotecas upstream. A aplicação de patch de problemas de segurança é essencial para criar a confiança do cliente. Como essas alterações não têm efeito no manifesto publicado, seus clientes não precisarão executar nenhuma ação para obter as versões mais recentes do suplemento.

## <a name="breaking-changes-in-officejs"></a>Alterações interruptivas no Office.js

A Microsoft 365 developer platform está comprometida em garantir a compatibilidade do seu suplemento. Nos esforçamos para evitar fazer alterações interruptivas na superfície e no comportamento da API. No entanto, há casos em que precisamos fazer atualizações interruptivas para fins de segurança ou confiabilidade. Nesses casos raros, as etapas a seguir são executadas para garantir que os usuários do seu suplemento não sejam afetados.

- Comunicados que descrevem os recursos afetados e as alterações recomendadas são feitos [no Microsoft 365 Blog do Desenvolvedor](https://devblogs.microsoft.com/microsoft365dev/).
- Se o suplemento for publicado no [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), você será contatado por meio das informações fornecidas.
- Sempre que possível, os administradores de locatários Microsoft 365 afetados (incluindo locatários de desenvolvedor) serão [contatados](https://developer.microsoft.com/microsoft-365/dev-program) por meio do [Centro de Mensagens](/microsoft-365/admin/manage/message-center). É responsabilidade do administrador contatar provedores de soluções de suplemento publicadas fora do AppSource.

### <a name="deprecation-policy"></a>Política de substituição

APIs ou ferramentas com alternativas melhores podem ser preteridas. A Microsoft se esforça para declarar algo como preterido pelo menos 24 meses antes de desativá-lo. Da mesma forma, para APIs individuais que geralmente estão disponíveis (GA), a Microsoft declara uma API como preterida com antecedência de pelo menos 24 meses antes de removê-la da versão de GA.

A substituição não significa necessariamente que o recurso ou a API será removido e inutilizável pelos desenvolvedores. Ele mostra que, após o período de 24 meses, a Microsoft não dá mais suporte à API ou ao recurso.

Quando uma API é marcada como preterida, é altamente recomendável que você migre para a versão mais recente assim que possível. Em alguns casos, anunciaremos que novos aplicativos devem começar a usar as novas APIs um pouco depois que as APIs originais forem preteridas. Nesses casos, apenas os aplicativos ativos que usam atualmente as APIs preteridas podem continuar a usá-las.

> [!IMPORTANT]
> O período de substituição de 24 meses será acelerado se a espera por tanto tempo representar um risco de segurança para o suplemento ou a Microsoft.

### <a name="app-assure"></a>Garantia de Aplicativo

O serviço [de Garantia](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) de Aplicativo da Microsoft cumpre a promessa da Microsoft de compatibilidade de aplicativos: seus aplicativos funcionarão Windows e Microsoft 365 Apps. Os engenheiros de Garantia de Aplicativo estão disponíveis para ajudar a resolver quaisquer problemas que você possa ter sem custo adicional.

Se você encontrar um problema de compatibilidade de aplicativo, os engenheiros da Garantia de Aplicativo trabalharão com você para ajudá-lo a resolver o problema. Nossos especialistas vão:

- Ajudar você a solucionar problemas e identificar uma causa raiz.
- Forneça diretrizes para ajudá-lo a corrigir o problema de compatibilidade do aplicativo.
- Envolva-se com ISVs (fornecedores independentes de software) em seu nome para corrigir parte do aplicativo para que ele seja funcional na versão mais moderna de nossos produtos.
- Trabalhe com as equipes de engenharia de produtos da Microsoft para corrigir bugs de produto.

Para saber mais sobre a Garantia de Aplicativo, assista [a Traga seus aplicativos para Microsoft Edge App Assure: dicas e truques](https://techcommunity.microsoft.com/t5/video-hub/bring-your-apps-to-microsoft-edge-with-app-assure-tips-and/ba-p/2167619). Para enviar sua solicitação de compatibilidade de aplicativos com a Garantia de [](https://aka.ms/AppAssureRequest) Aplicativo, preencha o formulário de registro Microsoft FastTrack ou envie um email para [achelp@microsoft.com.](mailto:achelp@microsoft.com)

## <a name="changes-to-yeoman-templates-and-web-dependencies"></a>Alterações em modelos Yeoman e dependências da Web

O [Gerador Yeoman para Office suplementos](../develop/yeoman-generator-overview.md) depende de várias bibliotecas da Microsoft e de outras pessoas. Essas bibliotecas são atualizadas independentemente de qualquer Microsoft 365 atividade. Todos os projetos criados com o gerador devem ser mantidos atualizados à medida que você desenvolve, publica e mantém seu suplemento. As ferramentas a seguir podem ajudar a garantir que seu projeto esteja usando versões seguras de qualquer biblioteca dependente.

- [npm auditoria](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot e outros recursos GitHub segurança](https://github.com/features/security)

Essas diretrizes também se aplicam a cópias de exemplos obtidos Office [exemplos](https://github.com/OfficeDev/Office-Add-in-samples) de código do suplemento e outras fontes.

### <a name="officejs-npm-package"></a>office.js pacote NPM

O [pacote NPM do office-js](https://www.npmjs.com/package/@microsoft/office-js) é uma cópia do que está hospedado naOffice.js de distribuição de conteúdo [ (CDN)](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library). Ele se destina a cenários em que o acesso direto ao CDN não é possível. O pacote NPM não se destina a fornecer referências com controle de versão para office.js. É altamente recomendável sempre usar o CDN para garantir que você esteja usando a versão mais recente das APIs Office JavaScript.

## <a name="current-best-practices"></a>Práticas recomendadas atuais

Embora nos esforçamos para manter a compatibilidade com versões anteriores, os padrões e as práticas que recomendamos evoluem continuamente. Nossa documentação se esforça para apresentar as práticas recomendadas atuais. Para se manter informado sobre novos recursos que podem melhorar sua funcionalidade existente, [junte-se aos nossos suplementos Office suplementos Community Chamada](../overview/office-add-ins-community-call.md).

## <a name="community-engagement"></a>Community compromisso

À medida que as atualizações são propostas para a Microsoft 365 developer platform, escutaremos comentários. Relate preocupações, possíveis consequências ou outras perguntas aos canais listados [Office recursos adicionais de suplementos](../resources/resources-links-help.md).
