---
title: Habilitar o SSO (logon único) em suplementos do Outlook que usam a ativação baseada em evento
description: Saiba como habilitar o SSO ao trabalhar em um suplemento de ativação baseado em evento.
ms.date: 06/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9a162e0ebe43c400f1b526d321cf049675047a6b
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092928"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Habilitar o SSO (logon único) em suplementos do Outlook que usam a ativação baseada em evento

Quando um suplemento do Outlook usa a ativação baseada em evento, os eventos são executados em um [runtime separado](../testing/runtimes.md). Depois de concluir as etapas em Autenticar um usuário com um token de logon único em um suplemento do [Outlook](authenticate-a-user-with-an-sso-token.md), siga as etapas adicionais descritas neste artigo para habilitar o SSO para o código de manipulação de eventos. Depois de habilitar o SSO, você poderá chamar a [API getAccessToken()](/javascript/api/office-runtime/officeruntime.auth) para obter um token de acesso com a identidade do usuário.

> [!IMPORTANT]
> Embora `OfficeRuntime.auth.getAccessToken` e `Office.auth.getAccessToken` execute a mesma funcionalidade de recuperar um token de acesso, `OfficeRuntime.auth.getAccessToken` recomendamos chamar seu suplemento baseado em evento. Essa API tem suporte em todas as versões de cliente do Outlook que dão suporte à ativação baseada em eventos e SSO. Por outro lado, só `Office.auth.getAccessToken` há suporte no Outlook no Windows a partir da versão 2111 (Build 14701.20000).

Para o Outlook no Windows, no manifesto do suplemento do Outlook, você identifica um único arquivo JavaScript a ser carregado para ativação baseada em evento. Você também precisa especificar ao Office que esse arquivo tem permissão para dar suporte ao SSO. Faça isso criando uma lista de todos os suplementos e seus arquivos JavaScript para fornecer ao Office por meio de um URI conhecido.

> [!NOTE]
> As etapas neste artigo se aplicam somente ao executar seu suplemento do Outlook no Windows. Isso ocorre porque o Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>Listar suplementos permitidos com um URI conhecido

Para listar quais suplementos têm permissão para trabalhar com SSO, crie um arquivo JSON que identifique cada arquivo JavaScript para cada suplemento. Em seguida, hospede esse arquivo JSON em um URI conhecido. Um URI conhecido permite a especificação de todos os arquivos JS hospedados que estão autorizados a obter tokens para a origem da Web atual. Isso garante que o proprietário da origem tenha controle total sobre quais arquivos JS hospedados devem ser usados em um suplemento e quais não estão, impedindo quaisquer vulnerabilidades de segurança em relação à representação, por exemplo.

O exemplo a seguir mostra como habilitar o SSO para dois suplementos (uma versão principal e uma versão beta). Você pode listar quantos suplementos for necessário, dependendo de quantos você fornecer do servidor Web.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Hospede o arquivo JSON em um local nomeado `.well-known` no URI na raiz da origem. Por exemplo, se a origem for `https://addin.contoso.com:8000/`, o URI conhecido será `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`.

A origem refere-se a um padrão de esquema + subdomínio + domínio + porta. O nome do local **deve** ser `.well-known`, e o nome do arquivo de **recurso deve** ser `microsoft-officeaddins-allowed.json`. Esse arquivo deve conter um objeto JSON `allowed` com um atributo chamado cujo valor é uma matriz de todos os arquivos JavaScript autorizados para SSO para seus respectivos suplementos.

## <a name="see-also"></a>Confira também

- [Autenticar um usuário com um token de logon único em um suplemento do Outlook](authenticate-a-user-with-an-sso-token.md)
- [Configurar seu suplemento do Outlook para ativação baseada em evento](autolaunch.md)
