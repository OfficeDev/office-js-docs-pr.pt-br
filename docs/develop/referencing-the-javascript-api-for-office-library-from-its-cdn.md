---
title: Fazendo referência à biblioteca da API JavaScript do Office
description: Saiba como fazer referência à biblioteca Office da API JavaScript e as definições de tipo no seu complemento.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 5348e1acab35e01dc6c467d20a65721fb98722d47c729edeb65a2efe4a8c45f8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080252"
---
# <a name="referencing-the-office-javascript-api-library"></a>Fazendo referência à biblioteca da API JavaScript do Office

A [Office api JavaScript](../reference/javascript-api-for-office.md) fornece as APIs que o seu complemento pode usar para interagir com o Office aplicativo. A maneira mais simples de fazer referência à biblioteca é usar a rede de distribuição de conteúdo (CDN) adicionando a marca a seguir na seção `<script>` `<head>` de sua página HTML.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Isso baixará e armazenará em cache os arquivos da API JavaScript Office primeira vez que o seu complemento for carregado para garantir que ele está usando a implementação mais atualizada do Office.js e seus arquivos associados para a versão especificada.

> [!IMPORTANT]
> Você deve fazer referência Office API JavaScript de dentro da seção da página para garantir que a API seja totalmente inicializada antes `<head>` de qualquer elemento do corpo.

## <a name="api-versioning-and-backward-compatibility"></a>Versão da API e compatibilidade com versões versões

No trecho HTML anterior, o na frente da URL CDN especifica a versão incremental mais recente na versão 1 do `/1/` `office.js` Office.js. Como a api Office JavaScript mantém a compatibilidade com versões anteriores, a versão mais recente continuará a dar suporte a membros da API introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, consulte [Update the version of your Office JAVAScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!NOTE]
> Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Habil IntelliSense para um projeto TypeScript

Além de fazer referência à API JavaScript Office como descrito anteriormente, você também pode habilitar o IntelliSense para o projeto de add-in TypeScript usando as definições de tipo de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Para fazer isso, execute o seguinte comando em um prompt de sistema habilitado para nó (ou janela git bash) na raiz da pasta do projeto. Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>APIs de visualização

As novas APIs JavaScript são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são necessários.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
