---
title: Opções de autenticação em suplementos do Outlook
description: 'Os suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.'
ms.date: 09/03/2021
ms.localizationpriority: high
---

# <a name="authentication-options-in-outlook-add-ins"></a>Opções de autenticação em suplementos do Outlook

O suplemento do Outlook pode acessar informações de qualquer lugar na Internet, seja do servidor que hospeda o suplemento, da sua rede interna ou de outro lugar na nuvem. Se essas informações estiverem protegidas, o suplemento precisará de uma forma de autenticar o usuário. Suplementos do Outlook oferecem diversos métodos de autenticação, dependendo do cenário específico.

## <a name="single-sign-on-access-token"></a>Token de acesso de logon único

Os tokens de acesso de logon único oferecem uma maneira simples de o suplemento autenticar e obter tokens de acesso para fazer uma chamada para a [API do Microsoft Graph](/graph/overview). Esse recurso reduz conflitos porque o usuário não precisa inserir credenciais.

> [!NOTE]
> A API de Logon Único é compatível com Word, Excel, Outlook e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md).
> Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Microsoft 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Considere usar tokens de acesso SSO se o suplemento:

- For usado principalmente por usuários do Microsoft 365
- Precisa de acesso para:
  - Os serviços Microsoft que são expostos como parte do Microsoft Graph
  - Um serviço que não seja da Microsoft que você controle

O método de autenticação SSO usa o [Fluxo Em Nome De do OAuth2 fornecido pelo Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). Ele exige o registro do suplemento no [Portal de Registro do Aplicativo](https://apps.dev.microsoft.com/) e a especificação dos escopos necessários do Microsoft Graph no manifesto.

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

Os suplementos também podem acessar serviços de terceiros que oferecem suporte ao OAuth2 para autorização. Considere usar tokens OAuth2 se o suplemento:

- Precisar acessar um serviço de terceiros fora do seu controle

Com esse método, o suplemento solicita que o usuário entre no serviço usando o método [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) para inicializar o fluxo do OAuth2 ou usando a [biblioteca office-js-helpers](https://github.com/OfficeDev/office-js-helpers) para o fluxo do OAuth2 Implícito.

## <a name="callback-tokens"></a>Tokens de retorno de chamada

Os tokens de retorno de chamada fornecem acesso à caixa de correio do usuário a partir do back-end do servidor usando o [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) ou a [API REST do Outlook](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Considere usar tokens de retorno de chamada se o suplemento:

- Precisar acessar a caixa de correio do usuário a partir do back-end do servidor.

Os suplementos obtêm tokens de retorno de chamada usando um dos métodos [getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods). O nível de acesso é controlado pelas permissões especificadas no manifesto do suplemento.
