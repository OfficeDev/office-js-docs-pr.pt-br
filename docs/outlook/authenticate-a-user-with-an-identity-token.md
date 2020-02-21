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
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a><span data-ttu-id="099e2-103">Autenticar um usuário com um token de identidade para o Exchange</span><span class="sxs-lookup"><span data-stu-id="099e2-103">Authenticate a user with an identity token for Exchange</span></span>

<span data-ttu-id="099e2-104">Os tokens de identidade do usuário do Exchange fornecem uma maneira de o suplemento identificar exclusivamente um usuário do suplemento.</span><span class="sxs-lookup"><span data-stu-id="099e2-104">Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user.</span></span> <span data-ttu-id="099e2-105">Ao estabelecer a identidade do usuário, você pode implementar um esquema de autenticação de logon único (SSO) para seu serviço de back-end que permitirá que os clientes que usam os suplementos do Outlook se conectem ao serviço sem fazer logon.</span><span class="sxs-lookup"><span data-stu-id="099e2-105">By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without logging in.</span></span> <span data-ttu-id="099e2-106">Confira [Token de identidade do usuário do Exchange](authentication.md#exchange-user-identity-token) para saber mais sobre quando usar esse tipo de token.</span><span class="sxs-lookup"><span data-stu-id="099e2-106">See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type.</span></span> <span data-ttu-id="099e2-107">Neste artigo, vamos dar uma olhada em uma forma simples de usar o token de identidade do Exchange para autenticar um usuário para seu back-end.</span><span class="sxs-lookup"><span data-stu-id="099e2-107">In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="099e2-108">Isso é apenas um exemplo simples de uma implementação de SSO.</span><span class="sxs-lookup"><span data-stu-id="099e2-108">This is just a simple example of an SSO implementation.</span></span> <span data-ttu-id="099e2-109">Como sempre, quando você está lidando com identidade e autenticação, deve garantir que seu código atenda aos requisitos de segurança de sua organização.</span><span class="sxs-lookup"><span data-stu-id="099e2-109">As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.</span></span>

## <a name="send-the-id-token-with-each-request"></a><span data-ttu-id="099e2-110">Enviar o token de ID com cada solicitação</span><span class="sxs-lookup"><span data-stu-id="099e2-110">Send the ID token with each request</span></span>

<span data-ttu-id="099e2-111">A primeira etapa é que o seu suplemento obtenha o token de identidade do usuário do Exchange do servidor chamando [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="099e2-111">The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span> <span data-ttu-id="099e2-112">Em seguida, o suplemento envia esse token com cada solicitação realizada para o back-end.</span><span class="sxs-lookup"><span data-stu-id="099e2-112">Then the add-in sends this token with every request it makes to your back-end.</span></span> <span data-ttu-id="099e2-113">Isso pode ocorrer em um cabeçalho ou como parte do corpo da solicitação.</span><span class="sxs-lookup"><span data-stu-id="099e2-113">This could be in a header, or as part of the request body.</span></span>

## <a name="validate-the-token"></a><span data-ttu-id="099e2-114">Validar o token</span><span class="sxs-lookup"><span data-stu-id="099e2-114">Validate the token</span></span>

<span data-ttu-id="099e2-115">O back-end DEVE validar o token antes de aceitá-lo.</span><span class="sxs-lookup"><span data-stu-id="099e2-115">The back-end MUST validate the token before accepting it.</span></span> <span data-ttu-id="099e2-116">Esta é uma etapa importante para garantir que o token foi emitido pelo servidor do Exchange do usuário.
</span><span class="sxs-lookup"><span data-stu-id="099e2-116">This is an important step to ensure that the token was issued by the user's Exchange server.</span></span> <span data-ttu-id="099e2-117">Para obter informações sobre a validação de tokens de identidade do usuário do Exchange, confira [Validar um token de identidade do Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="099e2-117">For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span>

<span data-ttu-id="099e2-118">Após validada e decodificada, a carga do token terá uma aparência semelhante à seguinte.</span><span class="sxs-lookup"><span data-stu-id="099e2-118">Once validated and decoded, the payload of the token looks something like the following.</span></span>

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

## <a name="map-the-token-to-a-user-in-your-backend"></a><span data-ttu-id="099e2-119">Mapear o token para um usuário em seu back-end</span><span class="sxs-lookup"><span data-stu-id="099e2-119">Map the token to a user in your backend</span></span>

<span data-ttu-id="099e2-120">O serviço de back-end pode calcular uma ID de usuário exclusiva a partir do token e mapeá-la para um usuário em seu sistema de usuário interno.</span><span class="sxs-lookup"><span data-stu-id="099e2-120">Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system.</span></span> <span data-ttu-id="099e2-121">Por exemplo, se usar um banco de dados para armazenar os usuários, você poderá adicionar essa ID exclusiva ao registro do usuário no banco de dados.</span><span class="sxs-lookup"><span data-stu-id="099e2-121">For example, if you use a database to store users, you could add this unique ID to the user's record in your database.</span></span>

### <a name="generate-a-unique-id"></a><span data-ttu-id="099e2-122">Gerar uma ID exclusiva</span><span class="sxs-lookup"><span data-stu-id="099e2-122">Generate a unique ID</span></span>

<span data-ttu-id="099e2-123">Recomendamos usar uma combinação das propriedades `msexchuid` e `amurl`.</span><span class="sxs-lookup"><span data-stu-id="099e2-123">We recommend that you use a combination of the `msexchuid` and `amurl` properties.</span></span> <span data-ttu-id="099e2-124">Você pode, por exemplo, concatenar os dois valores em conjunto e gerar uma cadeia de caracteres codificada em Base64.</span><span class="sxs-lookup"><span data-stu-id="099e2-124">For example, you could concatenate the two values together and generate a base 64-encoded string.</span></span> <span data-ttu-id="099e2-125">Esse valor poderá sempre ser confiavelmente gerado a partir do token para que você possa mapear um token de identidade do usuário do Exchange para o usuário em seu sistema.</span><span class="sxs-lookup"><span data-stu-id="099e2-125">This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.</span></span>

### <a name="check-the-user"></a><span data-ttu-id="099e2-126">Verificar o usuário</span><span class="sxs-lookup"><span data-stu-id="099e2-126">Check the user</span></span>

<span data-ttu-id="099e2-127">Com a ID exclusiva gerada, a próxima etapa é verificar se há um usuário em seu sistema com essa ID associada.</span><span class="sxs-lookup"><span data-stu-id="099e2-127">With the unique ID generated, the next step is to check for a user in your system with that associated ID.</span></span>

- <span data-ttu-id="099e2-128">Se o usuário for encontrado, o back-end tratará a solicitação como autenticada e permitirá o progresso da solicitação.</span><span class="sxs-lookup"><span data-stu-id="099e2-128">If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.</span></span>

- <span data-ttu-id="099e2-129">Se o usuário não for encontrado, o back-end retornará um erro indicando que o usuário precisa se conectar.</span><span class="sxs-lookup"><span data-stu-id="099e2-129">If the user is not found, then the back-end returns an error indicating that the user needs to sign in.</span></span> <span data-ttu-id="099e2-130">Em seguida, o suplemento solicita que o usuário acesse o back-end usando seu método de autenticação existente.</span><span class="sxs-lookup"><span data-stu-id="099e2-130">The add-in then prompts the user to sign in to the back-end using your existing authentication method.</span></span> <span data-ttu-id="099e2-131">Quando o usuário é autenticado, o token de identidade do usuário do Exchange é enviado com os detalhes da autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="099e2-131">Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details.</span></span> <span data-ttu-id="099e2-132">Em seguida, o back-end pode atualizar o registro do usuário no sistema com a identificação exclusiva.</span><span class="sxs-lookup"><span data-stu-id="099e2-132">The back-end can then update the user's record in your system with the unique ID.</span></span>
