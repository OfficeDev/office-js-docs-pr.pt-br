---
title: Referenciando a API JavaScript para a biblioteca do Office a partir de sua rede de distribuição de conteúdo (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 422cbd947dde09a8cd19559db9a86ddacd5e2dba
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348090"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Referenciando a API JavaScript para a biblioteca do Office a partir de sua rede de distribuição de conteúdo (CDN)

> [!NOTE]
> Além das etapas descritas neste artigo, se você quiser usar TypeScript e depois obter o Intellisense, você precisará executar o comando a seguir em um prompt de sistema habilitado para Node (ou janela git bash) partindo da raiz da pasta do seu projeto. Você deve ter [Node.js](https://nodejs.org) instalado (que inclui npm).
> 
> ```
> npm install --save-dev @types/office-js
> ```

A biblioteca da [API JavaScript para Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. 


A maneira mais simples de fazer referência à API é usar nossa CDN adicionando o seguinte `<script>` à tag `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

O `/1/` antes de `office.js` na URL da CDN especifica a versão incremental mais recente na versão 1 do Office .js. Como a API JavaScript para Office mantém a compatibilidade com versões anteriores, a última versão continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, confira [Atualizar a versão da API JavaScript para Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!IMPORTANT]
>  Ao desenvolver um suplemento para qualquer aplicativo host do Office, faça referência à API JavaScript para Office de dentro da seção `<head>` da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo. Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação. Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.       

## <a name="see-also"></a>Confira também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript para Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
