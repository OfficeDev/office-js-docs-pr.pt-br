---
title: Autenticação de usuário com um token de logon único
description: Saiba como usar o token de logon único fornecido por um suplemento do Outlook para implementar o SSO com o serviço.
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 41eddbcc1db05ca618506ce4810bf2bb795e59f7
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151888"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Autenticar um usuário com um token de login único em um Outlook de usuário

O SSO (logon único) oferece uma maneira simples para que o suplemento autentique usuários (e, opcionalmente, obtenha tokens de acesso para fazer uma chamada à [API do Microsoft Graph](/graph/overview)).

Usando este método, o suplemento pode obter um token de acesso com escopo para a API de back-end do servidor. O suplemento usa isso como um token de portador no cabeçalho `Authorization` para autenticar um retorno de chamada para sua API. Opcionalmente, você também pode ter seu código do lado do servidor.

- concluir o fluxo Em Nome De para obter um token de acesso com escopo para a API do Microsoft Graph
- Usar as informações de identidade no token para estabelecer a identidade do usuário e autenticar seus serviços de back-end

Para obter uma visão geral do SSO em suplementos do Office, confira [Habilitar o logon único para suplementos do Office](../develop/sso-in-office-add-ins.md) e [Autorizar acesso ao Microsoft Graph em suplementos do Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Habilitar a autenticação moderna em sua Microsoft 365 de Microsoft 365 de usuário

Para usar o SSO com um Outlook, você deve habilitar a Autenticação Moderna para o Microsoft 365 de autenticação. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Registrar seu suplemento

Para usar o SSO, o suplemento do Outlook precisará ter uma API Web no lado do servidor registrada com o AAD (Azure Active Directory) v2.0. Para mais informações, confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Fornecer consentimento quando estiver realizando o sideload de um suplemento

Quando você estiver desenvolvendo um complemento, você terá que fornecer consentimento antecipadamente. Para obter mais informações, consulte [Conceder consentimento ao administrador para o complemento](../develop/grant-admin-consent-to-an-add-in.md).

## <a name="update-the-add-in-manifest"></a>Atualizar o manifesto do suplemento

A próxima etapa para habilitar o SSO no suplemento é adicionar um elemento `WebApplicationInfo` ao fim do elemento `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md). Para mais informações, confira [Configurar o suplemento](../develop/sso-in-office-add-ins.md#configure-the-add-in).

## <a name="get-the-sso-token"></a>Obter o token SSO

O suplemento é um token SSO com script no lado do cliente. Para saber mais, confira [Adicionar o código no lado do cliente](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Usar o token SSO no back-end

Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá. Para obter detalhes sobre o que pode e deve ser feito no lado do servidor, confira [Adicionar código no lado do servidor](../develop/sso-in-office-add-ins.md#add-server-side-code).

> [!IMPORTANT]
> Ao usar o token SSO como uma identidade em um suplemento do *Outlook*, é recomendável [usar também o token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) como uma identidade alternativa. Os usuários do suplemento podem usar vários clientes, mas alguns podem não ser compatíveis com o fornecimento de tokens SSO. Usando o token de identidade do Exchange como uma alternativa, é possível evitar solicitações múltiplas de credenciais a esses usuários. Para mais informações, confira [Cenário: implementar o logon único no serviço em um Suplemento do Outlook](implement-sso-in-outlook-add-in.md).

## <a name="see-also"></a>Confira também

- Para ver um Outlook de exemplo que usa o token SSO para acessar a API do Microsoft Graph, consulte [Outlook SSO de complemento.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Referência da API do SSO](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [Conjunto de requisitos IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
