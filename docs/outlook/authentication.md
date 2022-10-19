---
title: Opções de autenticação em suplementos do Outlook
description: Os suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.
ms.date: 10/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d8ae8971c4095e5314885514226cd8f52728fb07
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607524"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Opções de autenticação em suplementos do Outlook

O suplemento do Outlook pode acessar informações de qualquer lugar na Internet, seja do servidor que hospeda o suplemento, da sua rede interna ou de outro lugar na nuvem. Se essas informações estiverem protegidas, o suplemento precisará de uma forma de autenticar o usuário. Suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.

## <a name="single-sign-on-access-token"></a>Token de acesso de logon único

Os tokens de acesso de logon único oferecem uma maneira simples de o suplemento autenticar e obter tokens de acesso para fazer uma chamada para a [API do Microsoft Graph](/graph/overview). Esse recurso reduz conflitos porque o usuário não precisa inserir credenciais.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Considere usar tokens de acesso SSO se o suplemento:

- For usado principalmente por usuários do Microsoft 365
- Precisa de acesso para:
  - Os serviços Microsoft que são expostos como parte do Microsoft Graph
  - Um serviço que não seja da Microsoft que você controle

O método de autenticação SSO usa o [Fluxo Em Nome De do OAuth2 fornecido pelo Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). Ele exige o registro do suplemento no [Portal de Registro do Aplicativo](https://apps.dev.microsoft.com/) e a especificação dos escopos necessários do Microsoft Graph no manifesto.

> [!NOTE]
> Se o suplemento estiver usando o manifesto do [Teams para Suplementos do Office (](../develop/json-manifest-overview.md)versão prévia), haverá alguma configuração de manifesto, mas os escopos do Microsoft Graph não são especificados. Os suplementos habilitados para SSO que usam o manifesto do Teams podem ser sideload, mas não podem ser implantados de outra maneira no momento.

Usando este método, o suplemento pode obter um token de acesso com escopo para a API de back-end do servidor. O suplemento usa isso como um token de portador no cabeçalho `Authorization` para autenticar um retorno de chamada para sua API. Nesse ponto, o servidor pode:

- concluir o fluxo Em Nome De para obter um token de acesso com escopo para a API do Microsoft Graph
- Usar as informações de identidade no token para estabelecer a identidade do usuário e autenticar seus serviços de back-end

Para obter uma visão geral mais detalhada, confira a [visão geral completa do método de autenticação SSO](../develop/sso-in-office-add-ins.md).

Para obter detalhes sobre como usar o token SSO em um suplemento do Outlook, confira [Autenticar o usuário com um token de logon único em um suplemento do Outlook](authenticate-a-user-with-an-sso-token.md).

Para obter um exemplo de suplemento que usa o token de SSO, confira [Suplemento de SSO do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).

## <a name="exchange-user-identity-token"></a>Token de identidade do usuário do Exchange

Os tokens de identidade do usuário do Exchange fornecem uma maneira de o suplemento estabelecer a identidade do usuário. Ao verificar a identidade do usuário, em seguida, você pode executar uma única autenticação no seu sistema de back-end e aceitar o token de identidade de usuário como uma autorização solicitações futuras. Use o token de identidade do usuário do Exchange:

- Quando o suplemento for usado principalmente por usuários locais do Exchange.
- Quando o suplemento precisar acessar um serviço que não seja da Microsoft que você controle.
- Como uma autenticação de recurso quando o suplemento está sendo executado em uma versão do Office que não suporta SSO.

Seu suplemento pode chamar [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)) para obter tokens de identidade do usuário do Exchange. Para obter detalhes sobre o uso desses tokens, confira [Autenticar um usuário com um token de identidade para o Exchange](authenticate-a-user-with-an-identity-token.md).

## <a name="access-tokens-obtained-via-oauth2-flows"></a>Tokens de acesso obtidos por meio de fluxos do OAuth2

Os suplementos também podem acessar serviços da Microsoft e de outros que oferecem suporte ao OAuth2 para autorização. Considere usar tokens OAuth2 se o suplemento:

- Precisa de acesso a um serviço fora do seu controle.

Usando esse método, seu suplemento solicita que o usuário entre no serviço usando o método [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) para inicializar o fluxo OAuth2.

## <a name="callback-tokens"></a>Tokens de retorno de chamada

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- Precisar acessar a caixa de correio do usuário a partir do back-end do servidor.

Os suplementos obtêm tokens de retorno de chamada usando um dos métodos [getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods). O nível de acesso é controlado pelas permissões especificadas no manifesto do suplemento.
