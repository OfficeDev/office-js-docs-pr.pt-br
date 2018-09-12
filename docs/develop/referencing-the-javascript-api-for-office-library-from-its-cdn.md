---
title: Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 9d3328ba09e2f69e76bd55f21064d52a8537cfa9
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943898"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)


A biblioteca da [API JavaScript para Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. 


A maneira mais simples de fazer referência à API é usar nossa CDN adicionando o seguinte `<script>` à marca `<head>` da sua página:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

O `/1/` antes de `office.js` da URL da CDN especifica a versão incremental mais recente na versão 1 do Office .js. Como a API JavaScript para Office mantém a compatibilidade com versões anteriores, a última versão continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, confira [Atualizar a versão da API JavaScript para Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.

> [!IMPORTANT]
>  Ao desenvolver um suplemento para qualquer aplicativo host do Office, faça referência à API JavaScript para Office de dentro da seção `<head>` da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo. Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação. Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.       

## <a name="see-also"></a>Veja também

- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript para Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
