---
title: Suporte ao Internet Explorer 11
description: Saiba como dar suporte ao Javascript do Internet Explorer 11 e ES5 no suplemento.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: aff6004af4ce28aea865cb34cd34e13e23fb549f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810271"
---
# <a name="support-internet-explorer-11"></a>Suporte ao Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer ainda usado em suplementos do Office**
>
> Algumas combinações de plataformas e versões do Office, incluindo versões perpétuas por meio do Office 2019, ainda usam o controle webview que vem com o Internet Explorer 11 para hospedar suplementos, conforme explicado em [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Recomendamos (mas não requer) que você continue a dar suporte a essas combinações, pelo menos de forma mínima, fornecendo aos usuários do seu suplemento uma mensagem de falha graciosa quando seu suplemento é iniciado na webview do Internet Explorer. Tenha esses pontos adicionais em mente:
>
> - Office na Web não é mais aberto no Internet Explorer. Consequentemente, o [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) não testa mais os suplementos em Office na Web usando o Internet Explorer como navegador.
> - O AppSource ainda testa combinações de versões da plataforma e da *área de trabalho* do Office que usam o Internet Explorer, no entanto, ele só emite um aviso quando o suplemento não dá suporte ao Internet Explorer; o suplemento não é rejeitado pelo AppSource.
> - A [ferramenta Script Lab](../overview/explore-with-script-lab.md) não dá mais suporte ao Internet Explorer.

Os suplementos do Office são aplicativos Web exibidos dentro de IFrames ao executar em Office na Web. Os suplementos do Office são exibidos usando controles de navegador inseridos ao executar no Office no Windows ou Office no Mac. Os controles de navegador inseridos são fornecidos pelo sistema operacional ou por um navegador instalado no computador do usuário.

Se você planeja dar suporte a versões mais antigas do Windows e do Office, seu suplemento deve funcionar no controle do navegador inserível baseado no Internet Explorer 11 (IE11). Para obter informações sobre quais combinações do Windows e do Office usam o controle de navegador baseado em IE11, consulte [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> O Internet Explorer 11 não dá suporte a alguns recursos HTML5, como mídia, gravação e localização. Se o suplemento precisar dar suporte ao Internet Explorer 11, você deverá projetar o suplemento para evitar esses recursos sem suporte ou o suplemento deve detectar quando o Internet Explorer está sendo usado e fornecer uma experiência alternativa que não use os recursos sem suporte. Para obter mais informações, consulte [Determinar no runtime se o suplemento está em execução no Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Suporte para versões recentes do JavaScript

O Internet Explorer 11 não dá suporte a versões JavaScript posteriores ao ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, ou TypeScript, você terá duas opções, conforme descrito neste artigo. Você também pode combinar essas duas técnicas.

### <a name="use-a-transpiler"></a>Usar um transpiler

Você pode escrever seu código no TypeScript ou javaScript moderno e transpilá-lo em tempo de compilação no JavaScript ES5. Os arquivos ES5 resultantes são o que você carrega no aplicativo Web do seu suplemento.

Há dois transpiladores populares. Ambos podem trabalhar com arquivos de origem que são TypeScript ou JavaScript pós-ES5. Eles também funcionam com React arquivos (.jsx e .tsx).

- [Babel](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

Consulte a documentação de qualquer um deles para obter informações sobre como instalar e configurar o transpiler em seu projeto de suplemento. Recomendamos que você use um gerenciador de tarefas, como [Grunt](https://gruntjs.com/) ou [WebPack](https://webpack.js.org/) , para automatizar a transpilação. Para obter um suplemento de exemplo que usa tsc, consulte [Microsoft Graph do Suplemento do Office React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Para obter um exemplo que usa babel, consulte [Suplemento de armazenamento offline](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Se você estiver usando o Visual Studio (não Visual Studio Code), o tsc provavelmente será mais fácil de usar. Você pode instalar o suporte para ele com um pacote nuget. Para obter mais informações, confira [JavaScript e TypeScript no Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Para usar babel com o Visual Studio, crie um script de build ou use o Gerenciador de Tarefas no Visual Studio com ferramentas como o [Gerenciador de Tarefas do WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou o [Gerenciador de Tarefas do NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Usar um polifill

Um [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) é JavaScript de versão anterior que duplica a funcionalidade de versões mais recentes do JavaScript. O polyfill funciona com em navegadores que não dão suporte às versões posteriores do JavaScript. Por exemplo, o método `startsWith` de cadeia de caracteres não fazia parte da versão ES5 do JavaScript e, portanto, não será executado no Internet Explorer 11. Há bibliotecas de polifill, escritas no ES5, que definem e implementam um `startsWith` método. Recomendamos a biblioteca de polifill [core-js](https://github.com/zloirock/core-js) .

Para usar uma biblioteca de polifill, carregue-a como qualquer outro arquivo ou módulo JavaScript. Por exemplo, você pode usar uma `<script>` marca no arquivo HTML da página inicial do suplemento (por exemplo `<script src="/js/core-js.js"></script>`), ou pode usar uma `import` instrução em um arquivo JavaScript (por exemplo, `import 'core-js';`). Quando o mecanismo JavaScript vir um método como `startsWith`, ele primeiro procurará para ver se há um método desse nome integrado ao idioma. Se houver, ele chamará o método nativo. Se, e somente se, o método não for interno, o mecanismo examinará todos os arquivos carregados para ele. Portanto, a versão polyfilled não é usada em navegadores que dão suporte à versão nativa.

A importação de toda a biblioteca core-js importará todos os recursos do core-js. Você também pode importar apenas os polífilis necessários para o suplemento do Office. Para obter instruções sobre como fazer isso, consulte [APIs CommonJS](https://github.com/zloirock/core-js#commonjs-api). A biblioteca core-js tem a maioria dos polífilis de que você precisa. Há algumas exceções detalhadas na seção [Polyfills Ausentes](https://github.com/zloirock/core-js#missing-polyfills) da documentação core-js. Por exemplo, ele não dá suporte `fetch`a , mas você pode usar o polyfill [fetch](https://github.com/github/fetch) .

Para obter um suplemento de exemplo que usa core.js, consulte [Suplemento do Word Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Determinar no runtime se o suplemento está em execução no Internet Explorer

Seu suplemento pode descobrir se ele está em execução no Internet Explorer lendo a propriedade [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Isso permite que o suplemento forneça uma experiência alternativa ou falhe graciosamente. Apresentamos um exemplo a seguir. Observe que o Internet Explorer envia uma cadeia de caracteres começando com "Trident" como o valor de userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade 
    //      either to perpetual Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> Normalmente, não é uma boa prática ler a `userAgent` propriedade. Verifique se você está familiarizado com o artigo, [detecção de navegador usando o agente de usuário](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent), incluindo as recomendações e alternativas à leitura `userAgent`. Em particular, se você estiver tomando a opção 1 na cláusula acima, considere usar a `else` detecção de recursos em vez de testar para o agente de usuário.
>
> A partir de 30 de setembro de 2021, o texto na seção [Qual parte do agente de usuário contém as informações que você está procurando?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) datas de antes do Internet Explorer 11 ser lançado. Ele ainda é geralmente preciso e as *tabelas* na seção da versão em inglês do artigo estão atualizadas. Da mesma forma, o texto e, na maioria dos casos, as tabelas, nas versões não em inglês do artigo estão desatualizadas.

## <a name="test-an-add-in-on-internet-explorer"></a>Testar um suplemento no Internet Explorer

Consulte [Teste do Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Recursos adicionais

- [Tabela de compatibilidade ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Posso usar... Tabelas de suporte para HTML5, CSS3 etc.](https://caniuse.com/)
