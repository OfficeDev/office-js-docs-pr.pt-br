---
title: Diretrizes de design de autenticação para suplementos do Office
ms.date: 02/09/2021
description: Saiba como projetar visualmente uma página de logon ou inscrição em um Suplemento do Office.
ms.localizationpriority: medium
ms.openlocfilehash: 45d11d509585a199135889273e6f9a96ce98e691
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810257"
---
# <a name="authentication-patterns"></a>Padrões de autenticação

Os suplementos podem exigir que os usuários entrem ou se inscrevam para acessar os recursos e funcionalidades. As caixas de entrada de nome de usuário e senha ou botões que iniciam fluxos de credenciais de terceiros são controles comuns da interface em experiências de autenticação. Uma experiência de autenticação simples e eficiente é uma primeira etapa importante para iniciar os usuários no uso de seu suplemento.

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:----|:----|
|Antes de entrar, descreva o valor do suplemento ou demonstre a funcionalidade sem exigir uma conta. |Espere que os usuários entrem sem compreender o valor e os benefícios do suplemento.|
|Oriente os usuários pelos fluxos de autenticação com um botão principal bem visível em cada tela. |Chame atenção para as tarefas secundárias e terciárias com outros botões e chamadas para ação.|
|Use rótulos de botão claros que descrevam tarefas específicas, como “Entrar” ou “Criar conta”. |Use rótulos de botão vagos como “Enviar” ou “Começar” para orientar os usuários por meio de fluxos de autenticação.|
|Use uma caixa de diálogo para concentrar a atenção do usuário em formulários de autenticação. |Encha seu painel de tarefas com uma primeira experiência de execução e formulários de autenticação.|
|Inclua pequenos recursos eficientes no fluxo como foco automático em caixas de entrada. |Adicione etapas desnecessárias à interação como exigir que os usuários cliquem nos campos de formulário.|
|Forneça uma maneira de os usuários sairem e reauthenticar. |Force os usuários a fazer a desinstalação para alternar identidades.|

## <a name="authentication-flow"></a>Fluxo de autenticação

1. Marcador de primeira execução: coloque o botão de entrada como uma chamada para ação clara na primeira experiência de execução do seu suplemento.

    ![Captura de tela mostrando um painel de tarefas de suplemento em um aplicativo do Office.](../images/add-in-fre-value-placemat.png)

1. Caixa de diálogo de opções do provedor de identidade: exiba uma lista clara de provedores de identidade, incluindo um formulário de nome de usuário e senha, se aplicável. A interface de usuário do seu suplemento poderá ser bloqueada enquanto a caixa de diálogo de autenticação estiver aberta.

    ![Captura de tela mostrando a caixa de diálogo Escolhas do Provedor de Identidade em um aplicativo do Office.](../images/add-in-auth-choices-dialog.png)

1. Entrada de um provedor de identidade: os provedores de identidade têm as próprias interfaces de usuário. Microsoft Azure Active Directory permite a personalização das páginas do painel de entrada e acesso para uma aparência consistente com seu serviço. [Saiba mais](/azure/active-directory/fundamentals/customize-branding).

    ![Captura de tela mostrando a caixa de diálogo Entrada do Provedor de Identidade em um aplicativo do Office.](../images/add-in-auth-identity-sign-in.png)

1. Progresso: indique o progresso enquanto as configurações e a interface do usuário são carregadas.

    ![Captura de tela mostrando uma caixa de diálogo com um indicador de progresso em um aplicativo do Office.](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> Ao usar o serviço de identidade da Microsoft, você terá a oportunidade de usar um botão de entrada com marca que poderá ser personalizado com temas claros e escuros. Saiba mais.

## <a name="single-sign-on-authentication-flow"></a>Fluxo de autenticação de Sign-On único

> [!NOTE]
> Atualmente, há suporte para a API de logon único para Word, Excel, Outlook e PowerPoint. Para obter mais informações sobre o suporte ao logon único, consulte [Conjuntos de requisitos identityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Microsoft 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Use logon único para uma experiência mais suave do usuário final. A identidade do usuário no Office (uma conta da Microsoft ou uma identidade do Microsoft 365) é usada para entrar no suplemento. Como resultado, os usuários só entrarão uma vez. Isso remove conflitos na experiência e faz com que os clientes comecem a usar o suplemento sem dificuldades.

1. Como um suplemento está sendo instalado, um usuário verá uma janela de consentimento semelhante à seguinte:

    ![Captura de tela mostrando a janela de consentimento em um aplicativo do Office quando um suplemento está sendo instalado.](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > O publicador do suplemento terá controle sobre o logotipo, sobre as cadeias de caracteres e escopos de permissão incluídos na janela de consentimento. A interface do usuário é pré-configurada pela Microsoft.

1. O suplemento será carregado após o consentimento do usuário. Ele pode extrair e exibir todas as informações personalizadas necessárias do usuário.

    ![Captura de tela mostrando um aplicativo do Office com botões de suplemento exibidos na faixa de opções.](../images/add-in-ribbon.png)

## <a name="see-also"></a>Confira também

- Saiba mais sobre [o desenvolvimento de suplementos de SSO](../develop/sso-in-office-add-ins.md)
