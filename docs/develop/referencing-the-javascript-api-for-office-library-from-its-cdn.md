---
title: Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6b9512d5d0969e185902d7ab9d3227e820c4d0dc
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353815"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)

> [!NOTE]
> Além das etapas descritas neste artigo, se você quiser usar o TypeScript e obter o IntelliSense, precisará executar o seguinte comando em um prompt de sistema com nó habilitado (ou na janela git bash) na raiz da pasta do seu projeto. Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).
> 
> ```command&nbsp;line
> npm install --save-dev @types/office-js
> ```

A biblioteca da [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. 


A maneira mais simples de fazer referência à API é usar nossa CDN adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

O `/1/` antes de `office.js` da URL da CDN especifica a versão incremental mais recente na versão 1 do Office .js. Como a API JavaScript para Office mantém a compatibilidade com versões anteriores, a última versão continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, confira [Atualizar a versão da API JavaScript para Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!IMPORTANT]
> Ao desenvolver um suplemento para qualquer aplicativo host do Office, faça referência à API JavaScript para Office de dentro da seção `<head>` da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo. Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação. Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.

## <a name="see-also"></a>Confira também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office)
