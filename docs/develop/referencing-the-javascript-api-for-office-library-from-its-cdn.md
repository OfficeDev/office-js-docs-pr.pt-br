---
title: Fazendo referência à biblioteca da API JavaScript do Office
description: Saiba como fazer referência à biblioteca de API JavaScript do Office e às definições de tipo no seu complemento.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 346a34c0cbc31b5e569a5106dcd2bc01593b114a
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505189"
---
# <a name="referencing-the-office-javascript-api-library"></a>Fazendo referência à biblioteca da API JavaScript do Office

A [biblioteca de API JavaScript](../reference/javascript-api-for-office.md) do Office fornece as APIs que seu complemento pode usar para interagir com o aplicativo do Office. A maneira mais simples de fazer referência à biblioteca é usar a rede de distribuição de conteúdo (CDN) adicionando a seguinte marca na seção de `<script>` `<head>` sua página HTML:  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Isso baixará e armazenará em cache os arquivos da API JavaScript do Office na primeira vez que o seu complemento for carregado para garantir que ele está usando a implementação mais atualizada do Office.js e seus arquivos associados para a versão especificada.

> [!IMPORTANT]
> Você deve fazer referência à API JavaScript do Office de dentro da seção da página para garantir que a API seja totalmente `<head>` inicializada antes de qualquer elemento do corpo.

## <a name="api-versioning-and-backward-compatibility"></a>Versão da API e compatibilidade com versões versões

No trecho HTML anterior, o na frente da URL da CDN especifica a versão incremental mais recente na versão `/1/` `office.js` 1 do Office.js. Como a API JavaScript do Office mantém a compatibilidade com versões anteriores, a versão mais recente continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, consulte [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!NOTE]
> Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Habil IntelliSense para um projeto TypeScript

Além de fazer referência à API JavaScript do Office conforme descrito anteriormente, você também pode habilitar o IntelliSense para o projeto de add-in TypeScript usando as definições de tipo de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Para fazer isso, execute o seguinte comando em um prompt de sistema habilitado para nó (ou janela git bash) na raiz da pasta do projeto. Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>APIs de visualização

As novas APIs JavaScript são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são necessários.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
