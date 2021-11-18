---
title: Habilitar o SSO (login único) em Outlook que usam a ativação baseada em evento
description: Saiba como habilitar o SSO ao trabalhar em um complemento de ativação baseado em eventos.
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 66d1edb8b7b0092ee107b73af24d5420caee8677
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066650"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Habilitar o SSO (login único) em Outlook que usam a ativação baseada em evento

Quando um Outlook de usuário usa a ativação baseada em evento, os eventos são executados em um tempo de execução JavaScript separado. Depois de concluir as etapas em [Autenticar](authenticate-a-user-with-an-sso-token.md)um usuário com um token de logom único em um Outlook de Outlook, siga as etapas adicionais descritas neste artigo para habilitar o SSO para o código de manipulação de eventos. Depois de habilitar o SSO, você pode chamar a API para obter um `getAccessToken()` token de acesso com a identidade do usuário.

> [!NOTE]
> As etapas deste artigo só se aplicam ao executar seu Outlook de Windows. Isso porque Outlook no Windows usa um arquivo JavaScript, enquanto Outlook na Web usa um arquivo HTML que pode fazer referência ao mesmo arquivo JavaScript.

Para Outlook no Windows, no manifesto do seu Outlook, você identifica um único arquivo JavaScript a ser carregado para ativação baseada em evento. Você também precisa especificar para Office que esse arquivo tem permissão para dar suporte ao SSO. Há duas abordagens para fazer isso. Você pode criar uma lista de todos os complementos e seus arquivos JavaScript para fornecer Office por meio de um URI conhecido. Ou você pode adicionar um header de resposta personalizado para habilitar o SSO.

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

Hospede o arquivo JSON em um local `.well-known` chamado no URI na raiz da origem. Por exemplo, se a origem for , o URI conhecido `https://addin.contoso.com:8000/` será `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json` .

A origem refere-se a um padrão de esquema + subdomínio + domínio + porta. O nome do local **deve** `.well-known` ser , e o nome do arquivo de recurso **deve** ser `microsoft-officeaddins-allowed.json` . Esse arquivo deve conter um objeto JSON com um atributo chamado cujo valor é uma matriz de todos os arquivos JavaScript autorizados para SSO para seus `allowed` respectivos complementos.

## <a name="add-a-custom-response-header"></a>Adicionar um header de resposta personalizado

Uma segunda abordagem é adicionar um header de resposta personalizado chamado `MS-OfficeAddins-Allowed-Origin` . O valor do header deve ser a origem do arquivo JavaScript.

Por exemplo, se o arquivo JavaScript estiver localizado em `https://addin.contoso.com:8000/main/js/autorun.js` , adicione o seguinte header de resposta.

`MS-OfficeAddins-Allowed-Origin : https://addin.contoso.com:8000`

Você precisará consultar sua documentação específica do servidor Web para saber como adicionar o header de resposta personalizado.

## <a name="see-also"></a>Confira também

- [Autenticar um usuário com um token de login único em um Outlook de usuário](authenticate-a-user-with-an-sso-token.md)
- [Configurar seu Outlook para ativação baseada em eventos](autolaunch.md)
