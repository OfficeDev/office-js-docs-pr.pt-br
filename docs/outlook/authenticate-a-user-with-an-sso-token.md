---
title: Autenticação de usuário com um token de logon único
description: Saiba como usar o token de logon único fornecido por um suplemento do Outlook para implementar o SSO com o serviço.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23b7936cc0ba4453a2a10cbfe0731941a913c118
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607440"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Autenticar um usuário com um token de logon único em um suplemento do Outlook

O SSO (logon único) oferece uma maneira simples para que o suplemento autentique usuários (e, opcionalmente, obtenha tokens de acesso para fazer uma chamada à [API do Microsoft Graph](/graph/overview)).

Usando este método, o suplemento pode obter um token de acesso com escopo para a API de back-end do servidor. O suplemento usa isso como um token de portador no cabeçalho `Authorization` para autenticar um retorno de chamada para sua API. Opcionalmente, você também pode ter o código do lado do servidor.

- concluir o fluxo Em Nome De para obter um token de acesso com escopo para a API do Microsoft Graph
- Usar as informações de identidade no token para estabelecer a identidade do usuário e autenticar seus serviços de back-end

Para obter uma visão geral do SSO em suplementos do Office, confira [Habilitar o logon único para suplementos do Office](../develop/sso-in-office-add-ins.md) e [Autorizar acesso ao Microsoft Graph em suplementos do Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Habilitar a autenticação moderna em seu locatário do Microsoft 365

Para usar o SSO com um suplemento do Outlook, você deve habilitar a Autenticação Moderna para o locatário do Microsoft 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Registrar seu suplemento

Para usar o SSO, o suplemento do Outlook precisará ter uma API Web no lado do servidor registrada com o AAD (Azure Active Directory) v2.0. Para mais informações, confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Fornecer consentimento quando estiver realizando o sideload de um suplemento

Ao desenvolver um suplemento, você precisará fornecer consentimento com antecedência. Para obter mais informações, consulte [Conceder consentimento do administrador ao suplemento](../develop/grant-admin-consent-to-an-add-in.md).

## <a name="update-the-add-in-manifest"></a>Atualizar o manifesto do suplemento

A próxima etapa para habilitar o SSO no suplemento é adicionar algumas informações ao manifesto do registro de plataforma de identidade da Microsoft suplemento. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: adicione um `WebApplicationInfo` elemento no final do `VersionOverridesV1_1` [elemento VersionOverrides](/javascript/api/manifest/versionoverrides) . Em seguida, adicione os elementos filho necessários. Para obter informações detalhadas sobre a marcação, consulte [Configurar o suplemento](../develop/sso-in-office-add-ins.md#configure-the-add-in).
- **Manifesto do Teams (versão prévia)**: adicione uma propriedade "webApplicationInfo" ao objeto `{ ... }` raiz no manifesto. Dê a este objeto uma propriedade "id" filho definida como a ID do aplicativo Web do suplemento como ele foi gerado no portal do Azure quando você registrou o suplemento. (Consulte a seção [Registrar seu suplemento anteriormente](#register-your-add-in) neste artigo.) Além disso, dê a ele uma propriedade filho de "recurso" definida como o mesmo **URI da ID** do Aplicativo que você definiu quando registrou o suplemento. Esse URI deve ter o formulário `api://<fully-qualified-domain-name>/<application-id>`. Apresentamos um exemplo a seguir.

   ```json
   "webApplicationInfo": {
        "id": "a661fed9-f33d-4e95-b6cf-624a34a2f51d",
        "resource": "api://addin.contoso.com/a661fed9-f33d-4e95-b6cf-624a34a2f51d"
    },
   ```

  > [!NOTE]
  > Os suplementos habilitados para SSO que usam o manifesto do Teams podem ser sideload, mas não podem ser implantados de outra maneira no momento.

## <a name="get-the-sso-token"></a>Obter o token SSO

O suplemento é um token SSO com script no lado do cliente. Para saber mais, confira [Adicionar o código no lado do cliente](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Usar o token SSO no back-end

Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá. Para obter detalhes sobre o que pode e deve ser feito no lado do servidor, confira [Adicionar código no lado do servidor](../develop/sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).

> [!IMPORTANT]
> Ao usar o token SSO como uma identidade em um suplemento do *Outlook*, é recomendável [usar também o token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) como uma identidade alternativa. Os usuários do suplemento podem usar vários clientes, mas alguns podem não ser compatíveis com o fornecimento de tokens SSO. Usando o token de identidade do Exchange como uma alternativa, é possível evitar solicitações múltiplas de credenciais a esses usuários. Para mais informações, confira [Cenário: implementar o logon único no serviço em um Suplemento do Outlook](implement-sso-in-outlook-add-in.md).

## <a name="sso-for-event-based-activation"></a>SSO para ativação baseada em evento

Há etapas adicionais a serem executadas se o suplemento usar a ativação baseada em evento. Para obter mais informações, consulte [Habilitar SSO (logon único) em suplementos do Outlook que usam a ativação baseada em evento](use-sso-in-event-based-activation.md).

## <a name="see-also"></a>Confira também

- [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))
- Para obter um exemplo de suplemento do Outlook que usa o token de SSO para acessar o Microsoft API do Graph, consulte [o SSO do Suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).
- [Referência da API do SSO](/javascript/api/office/office.auth#office-office-auth-getaccesstoken-member(1))
- [Conjunto de requisitos IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Habilitar o SSO (logon único) em suplementos do Outlook que usam a ativação baseada em evento](use-sso-in-event-based-activation.md)
