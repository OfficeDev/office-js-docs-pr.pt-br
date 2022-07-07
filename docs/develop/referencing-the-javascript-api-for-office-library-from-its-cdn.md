---
title: Fazendo referência à biblioteca da API JavaScript do Office
description: Saiba como referenciar a biblioteca de API JavaScript do Office e as definições de tipo em seu suplemento.
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38121fe3d3df0a86fef3e2c8e3a58399640f1e2a
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660113"
---
# <a name="referencing-the-office-javascript-api-library"></a>Fazendo referência à biblioteca da API JavaScript do Office

A [biblioteca de API JavaScript do Office](../reference/javascript-api-for-office.md) fornece as APIs que seu suplemento pode usar para interagir com o aplicativo do Office. A maneira mais simples de fazer referência à biblioteca é usar a CDN (rede de distribuição de conteúdo) `<script>` `<head>` adicionando a seguinte marca na seção da página HTML.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Isso baixará e armazenará em cache os arquivos da API JavaScript do Office na primeira vez que o suplemento for carregado para garantir que ele esteja usando a implementação mais atualizada do Office.js e seus arquivos associados para a versão especificada.

> [!IMPORTANT]
> Você deve fazer referência à API `<head>` JavaScript do Office de dentro da seção da página para garantir que a API seja totalmente inicializada antes de qualquer elemento do corpo.

## <a name="api-versioning-and-backward-compatibility"></a>Controle de versão da API e compatibilidade com versões anteriores

No snippet HTML anterior, `/1/` o na `office.js` frente da URL da CDN especifica a versão incremental mais recente na versão 1 do Office.js. Como a API JavaScript do Office mantém a compatibilidade com versões anteriores, a versão mais recente continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, consulte Atualizar a versão da [API JavaScript do Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!NOTE]
> Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Habilitando o IntelliSense para um projeto TypeScript

Além de fazer referência à API JavaScript do Office, conforme descrito anteriormente, você também pode habilitar o IntelliSense para projeto de suplemento TypeScript usando as definições de tipo do [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Para fazer isso, execute o comando a seguir em um prompt do sistema habilitado para Node (ou janela git bash) na raiz da pasta do projeto. Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>APIs de visualização

Novas APIs JavaScript são introduzidas primeiro em "versão prévia" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são adquiridos.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Confira também

- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
