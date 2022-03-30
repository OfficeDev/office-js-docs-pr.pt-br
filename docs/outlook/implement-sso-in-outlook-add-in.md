---
title: 'Cenário: implementar o logon único no seu serviço'
description: Saiba como usar o token de logon único e o token de identidade do Exchange fornecidos por um suplemento do Outlook para implementar o SSO com o serviço.
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2b9c4031a0011d2333582b4a10abe42f6844f763
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496919"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a>Cenário: implementar o logon único no serviço em um suplemento do Outlook

Neste artigo exploraremos um método recomendado de usar o [token de acesso de logon único](authenticate-a-user-with-an-sso-token.md) e o [token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) juntos para fornecer um logon único na implementação do seu próprio serviço de back-end. Usando dois tokens em conjunto, será possível aproveitar os benefícios do token de acesso SSO quando ele estiver disponível, garantindo que o suplemento funcionará quando ele não estiver disponível, como quando o usuário alterna para um cliente não compatível ou quando a caixa de correio do usuário está em um servidor do Exchange local.

Para um exemplo de complemento que implementa as ideias neste artigo, [consulte Outlook SSO de complemento](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).


> [!NOTE]
> Atualmente, a API de logon único é compatível com Word, Excel, Outlook e PowerPoint. Para mais informações sobre onde a API Logon Único tem suporte no momento, veja [Conjuntos de requisitos IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a autenticação moderna para a locação do Microsoft 365. Para informações sobre como fazer isso, consulte [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## <a name="why-use-the-sso-access-token"></a>Por que usar o token de acesso SSO?

O token de identidade do Exchange está disponível em todos os conjuntos de requisitos de APIs do suplemento, portanto, pode parecer tentador depender simplesmente desse token e ignorar o token SSO completamente. No entanto, o token SSO oferece algumas vantagens em relação ao token de identidade do Exchange, portanto, quando disponível, torna-se o método recomendado.

- O token SSO usa um formato padrão OpenID e é emitido pelo Azure. Isso simplifica bastante o processo de validação desses tokens. Em comparação, os tokens de identidade do Exchange usam um formato personalizado com base no Token Web JSON padrão, exigindo trabalho personalizado para validar o token.
- O token SSO pode ser usado pelo back-end para recuperar um token de acesso do Microsoft Graph sem que o usuário tenha que fazer qualquer outra ação de entrada.
- O token SSO fornece informações avançadas de identidade, como o nome para exibição do usuário.

## <a name="add-in-scenario"></a>Cenário de suplemento

Para este exemplo, considere um suplemento formado pela interface do usuário e scripts (HTML + JavaScript) do suplemento e uma API Web de back-end chamada pelo suplemento. A API Web de back-end faz chamadas para a [API do Microsoft Graph](/graph/overview) e a API de Dados da Contoso, uma API fictícia criada por terceiros. Como a API do Microsoft Graph, a API de Dados da Contoso requer autenticação OAuth. O requisito é que a API Web de back-end seja capaz de chamar as duas APIs sem ter que solicitar ao usuário que forneça credenciais sempre que um token de acesso expirar.

Para fazer isso, a API de backend cria um banco de dados de usuários seguro. Cada usuário receberá uma entrada no banco de dados onde o back-end armazenará tokens de atualização de vida longa da API do Microsoft Graph API e da API de Dados da Contoso. A marcação JSON a seguir representa uma entrada do usuário no banco de dados.

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

O suplemento inclui o token de acesso SSO (se estiver disponível) ou o token de identidade do Exchange (se o token SSO não estiver disponível) com todas as chamadas feitas para a API Web de back-end.

### <a name="add-in-startup"></a>Inicialização do suplemento

1. Quando o suplemento iniciar, ele enviará uma solicitação à API Web de back-end para determinar se o usuário está registrado (por exemplo, se tem um registro associado no banco de dados do usuário) e se a API tem tokens de atualização para o Graph e para a Contoso. Nessa chamada, o suplemento inclui o token SSO (se disponível) e o token de identidade.

1. A API Web utiliza os métodos em [Autenticar um usuário com um token de logon único em um suplemento do Outlook](authenticate-a-user-with-an-sso-token.md) e [Autenticar um usuário com um token de identidade do Exchange](authenticate-a-user-with-an-identity-token.md) para validar e gerar um identificador exclusivo a partir dos dois tokens.

1. Se um token SSO tiver sido fornecido, a API Web consultará o banco de dados do usuário em busca de uma entrada que tenha um valor `ssoId` que corresponda ao identificador exclusivo gerado pelo token SSO.
   - Se não houver uma entrada, vá para a próxima etapa.
   - Se houver uma entrada, vá para a etapa 5.

1. A API Web consultará o banco de dados em busca de uma entrada que tenha um valor `exchangeId` que corresponda ao identificador exclusivo gerado pelo token de identidade do Exchange.
   - Se houver uma entrada e um token SSO tiver sido fornecido, atualize o registro do usuário no banco de dados para definir o valor `ssoId` para o identificador exclusivo gerado a partir do token SSO e prossiga para a etapa 5.
   - Se houver uma entrada e nenhum token SSO tiver sido fornecido, prossiga para a etapa 5.
   - Se não houver entradas, crie uma nova entrada. Defina `ssoId` como o identificador exclusivo gerado por meio do token SSO (se disponível) e defina `exchangeId` como o identificador exclusivo gerado por meio do token de identidade do Exchange.

1. Verifique se há um token de atualização válido no valor `graphRefreshToken` do usuário.
   - Se o valor for inválido ou estiver ausente no token SSO fornecido, use o [Fluxo Em Nome De do OAuth2](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) para obter um token de acesso e atualizar o token do Microsoft Graph. Salve o token de atualização válido no valor `graphRefreshToken` para o usuário.

1. Procure tokens de atualização válidos em `graphRefreshToken` e `contosoRefreshToken`.
   - Se ambos valores forem válidos, responda o suplemento para indicar que o usuário já está registrado e configurado.
   - Se o valor for inválido, responda o suplemento para indicar que a configuração do usuário é obrigatória, além de quais serviços (Contoso ou Graph) precisam ser configurados.

1. O suplemento verifica a resposta.
   - Se o usuário já estiver registrado e configurado, o suplemento prosseguirá com a operação normal.
   - Se a configuração do usuário for exigida, o suplemento entrará em modo "configuração" e solicitará que o usuário autorize o suplemento.

### <a name="authorize-the-backend-web-api"></a>Autorizar a API Web de back-end

Para minimizar a necessidade de ter que informar o usuário de fazer login, o ideal é que o procedimento para autorizar a API Web de back-end a chamar a API do Microsoft Graph e a API de Dados da Contoso ocorra apenas uma vez.

Com base na resposta da API Web de back-end, talvez o suplemento precise da autorização do usuário da API do Microsoft Graph, da API de Dados da Contoso ou de ambas APIs. Como as duas APIs usam a autenticação OAuth2, o método é semelhante para ambas.

1. O suplemento informa o usuário que precisa que ele autorize o uso da API e pede a ele para clicar em um link ou em um botão para iniciar o processo.

    > [!NOTE]
    > O exemplo de add-in no SSO do Outlook de complemento mostra como usar a [API de Diálogo](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) e [a](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) biblioteca [office-js-helpers](https://github.com/OfficeDev/office-js-helpers) como opções para iniciar o fluxo de Código de Autorização [OAuth2](/azure/active-directory/develop/active-directory-protocols-oauth-code) para a API.

1. Após a conclusão do fluxo, o suplemento envia o token de atualização à API Web de back-end e inclui o token SSO (se disponível) ou o token de identidade do Exchange.

1. A API Web de back-end localiza o usuário no banco de dados e atualiza o token de atualização apropriado.

1. O suplemento prossegue com a operação normal.

### <a name="normal-operation"></a>Operação normal

Sempre que o suplemento chamar a API Web de back-end, incluirá o token SSO ou o token de identidade do Exchange. A API Web de back-end localiza o usuário pelo token e usa os tokens de atualização armazenados para obter tokens de acesso da API do Microsoft Graph e da API de Dados da Contoso. Enquanto os tokens de atualização forem válidos, o usuário não terá que entrar novamente.
