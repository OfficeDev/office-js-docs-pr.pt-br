---
title: Autenticar um usuário com um token de identidade em um suplemento.
description: Saiba como usar o token de identidade fornecido por um suplemento do Outlook para implementar o SSO com o seu serviço.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4134aa8ff21262f2f384d141db002b56a4a32f0a
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165765"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a>Autenticar um usuário com um token de identidade para o Exchange

Os tokens de identidade do usuário do Exchange fornecem uma maneira de o suplemento identificar exclusivamente um usuário do suplemento. Ao estabelecer a identidade do usuário, você pode implementar um esquema de autenticação de logon único (SSO) para seu serviço de back-end que permitirá que os clientes que usam os suplementos do Outlook se conectem ao serviço sem fazer logon. Confira [Token de identidade do usuário do Exchange](authentication.md#exchange-user-identity-token) para saber mais sobre quando usar esse tipo de token. Neste artigo, vamos dar uma olhada em uma forma simples de usar o token de identidade do Exchange para autenticar um usuário para seu back-end.

> [!IMPORTANT]
> Isso é apenas um exemplo simples de uma implementação de SSO. Como sempre, quando você está lidando com identidade e autenticação, deve garantir que seu código atenda aos requisitos de segurança de sua organização.

## <a name="send-the-id-token-with-each-request"></a>Enviar o token de ID com cada solicitação

A primeira etapa é que o seu suplemento obtenha o token de identidade do usuário do Exchange do servidor chamando [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods). Em seguida, o suplemento envia esse token com cada solicitação realizada para o back-end. Isso pode ocorrer em um cabeçalho ou como parte do corpo da solicitação.

## <a name="validate-the-token"></a>Validar o token

O back-end DEVE validar o token antes de aceitá-lo. Esta é uma etapa importante para garantir que o token foi emitido pelo servidor do Exchange do usuário.
 Para obter informações sobre a validação de tokens de identidade do usuário do Exchange, confira [Validar um token de identidade do Exchange](validate-an-identity-token.md).

Após validada e decodificada, a carga do token terá uma aparência semelhante à seguinte.

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a>Mapear o token para um usuário em seu back-end

O serviço de back-end pode calcular uma ID de usuário exclusiva a partir do token e mapeá-la para um usuário em seu sistema de usuário interno. Por exemplo, se usar um banco de dados para armazenar os usuários, você poderá adicionar essa ID exclusiva ao registro do usuário no banco de dados.

### <a name="generate-a-unique-id"></a>Gerar uma ID exclusiva

Recomendamos usar uma combinação das propriedades `msexchuid` e `amurl`. Você pode, por exemplo, concatenar os dois valores em conjunto e gerar uma cadeia de caracteres codificada em Base64. Esse valor poderá sempre ser confiavelmente gerado a partir do token para que você possa mapear um token de identidade do usuário do Exchange para o usuário em seu sistema.

### <a name="check-the-user"></a>Verificar o usuário

Com a ID exclusiva gerada, a próxima etapa é verificar se há um usuário em seu sistema com essa ID associada.

- Se o usuário for encontrado, o back-end tratará a solicitação como autenticada e permitirá o progresso da solicitação.

- Se o usuário não for encontrado, o back-end retornará um erro indicando que o usuário precisa se conectar. Em seguida, o suplemento solicita que o usuário acesse o back-end usando seu método de autenticação existente. Quando o usuário é autenticado, o token de identidade do usuário do Exchange é enviado com os detalhes da autenticação do usuário. Em seguida, o back-end pode atualizar o registro do usuário no sistema com a identificação exclusiva.
