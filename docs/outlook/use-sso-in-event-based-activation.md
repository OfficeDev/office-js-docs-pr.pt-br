---
title: Habilitar o SSO (login único) em Outlook de complementos que usam a ativação baseada em evento
description: Saiba como habilitar o SSO ao trabalhar em um complemento de ativação baseado em eventos.
ms.date: 03/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: bb52678356fe0cf456cbbf023febee738cccdb31
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63710927"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Habilitar o SSO (login único) em Outlook de complementos que usam a ativação baseada em evento

Quando um Outlook de usuário usa a ativação baseada em evento, os eventos são executados em um tempo de execução JavaScript separado. Depois de concluir as etapas em [Autenticar](authenticate-a-user-with-an-sso-token.md) um usuário com um token de logom único em um Outlook de Outlook, siga as etapas adicionais descritas neste artigo para habilitar o SSO para o código de tratamento de eventos. Depois de habilitar o SSO, você pode `getAccessToken()` chamar a API para obter um token de acesso com a identidade do usuário.

> [!NOTE]
> As etapas deste artigo só se aplicam ao executar seu Outlook de Windows. Isso porque Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.

Para Outlook no Windows, no manifesto do seu Outlook de Outlook, você identifica um único arquivo JavaScript a ser carregado para ativação baseada em evento. Você também precisa especificar Office que esse arquivo tem permissão para dar suporte ao SSO. Você faz isso criando uma lista de todos os complementos e seus arquivos JavaScript para fornecer Office por meio de um URI conhecido.

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>Listar os complementos permitidos com um URI conhecido

Para listar quais complementos têm permissão para trabalhar com o SSO, crie um arquivo JSON que identifique cada arquivo JavaScript para cada complemento. Em seguida, hospede esse arquivo JSON em um URI conhecido. Um URI conhecido permite a especificação de todos os arquivos JS hospedados que estão autorizados a obter tokens para a origem da Web atual. Isso garante que o proprietário da origem tenha controle total sobre quais arquivos JS hospedados devem ser usados em um complemento e quais não estão, impedindo qualquer vulnerabilidade de segurança em torno da representação, por exemplo.

O exemplo a seguir mostra como habilitar o SSO para dois complementos (uma versão principal e uma versão beta). Você pode listar quantos complementos for necessário, dependendo de quantos você fornecer do seu servidor Web.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Hospede o arquivo JSON em um local chamado `.well-known` no URI na raiz da origem. Por exemplo, se a origem for `https://addin.contoso.com:8000/`, o URI conhecido será `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`.

A origem refere-se a um padrão de esquema + subdomínio + domínio + porta. O nome do local **deve** ser `.well-known`, e o nome do arquivo de **recurso deve** ser `microsoft-officeaddins-allowed.json`. Esse arquivo deve conter um objeto JSON `allowed` com um atributo chamado cujo valor é uma matriz de todos os arquivos JavaScript autorizados para SSO para seus respectivos complementos.

## <a name="see-also"></a>Confira também

- [Autenticar um usuário com um token de login único em um Outlook de usuário](authenticate-a-user-with-an-sso-token.md)
- [Configurar seu Outlook para ativação baseada em eventos](autolaunch.md)
