---
title: Habilitar o SSO (logon único) Outlook suplementos que usam a ativação baseada em evento
description: Saiba como habilitar o SSO ao trabalhar em um suplemento de ativação baseado em evento.
ms.date: 06/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 477ecb8c0ab84ab472763f83e342258998749861
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229726"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Habilitar o SSO (logon único) Outlook suplementos que usam a ativação baseada em evento

Quando um Outlook suplemento usa a ativação baseada em evento, os eventos são executados em um runtime separado do JavaScript. Depois de concluir as etapas em Autenticar um usuário com um token de logon único em um suplemento do Outlook, siga [as](authenticate-a-user-with-an-sso-token.md) etapas adicionais descritas neste artigo para habilitar o SSO para o código de manipulação de eventos. Depois de habilitar o SSO, você poderá chamar a [API getAccessToken()](/javascript/api/office-runtime/officeruntime.auth) para obter um token de acesso com a identidade do usuário.

> [!IMPORTANT]
> Embora `OfficeRuntime.auth.getAccessToken` e `Office.auth.getAccessToken` execute a mesma funcionalidade de recuperar um token de acesso, `OfficeRuntime.auth.getAccessToken` recomendamos chamar seu suplemento baseado em evento. Essa API tem suporte em todas as Outlook cliente que dão suporte à ativação e ao SSO baseados em eventos. Por outro lado, `Office.auth.getAccessToken` só há suporte no Outlook no Windows a partir da versão 2111 (build 14701.20000).

Para Outlook no Windows, no manifesto do suplemento Outlook, você identifica um único arquivo JavaScript a ser carregado para ativação baseada em evento. Você também precisa especificar para Office que esse arquivo tem permissão para dar suporte ao SSO. Faça isso criando uma lista de todos os suplementos e seus arquivos JavaScript para fornecer Office por meio de um URI conhecido.

> [!NOTE]
> As etapas neste artigo se aplicam somente ao executar seu Outlook suplemento no Windows. Isso ocorre porque Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode referenciar o mesmo arquivo JavaScript.

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

- [Autenticar um usuário com um token de logon único em um Outlook suplemento](authenticate-a-user-with-an-sso-token.md)
- [Configurar seu Outlook para ativação baseada em evento](autolaunch.md)
