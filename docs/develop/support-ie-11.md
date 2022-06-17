---
title: Suporte ao Internet Explorer 11
description: Saiba como dar suporte ao Internet Explorer 11 e javascript ES5 em seu suplemento.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1cb641f1ed1a75fcff23291d1fa566bbf6dc008b
ms.sourcegitcommit: fb3b1c6055e664d015703623661d624251ceb6b7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2022
ms.locfileid: "66136422"
---
# <a name="support-internet-explorer-11"></a>Suporte ao Internet Explorer 11

> [!IMPORTANT]
> **O Internet Explorer ainda é Office suplementos**
>
> Algumas combinações de plataformas e versões do Office, incluindo versões de compra única por meio do Office 2019, ainda usam o controle de modo de exibição da Web que vem com o Internet Explorer 11 para hospedar suplementos, conforme explicado em Navegadores usados por [suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Recomendamos (mas não exige) que você continue a dar suporte a essas combinações, pelo menos de maneira mínima, fornecendo aos usuários do seu suplemento uma mensagem de falha normal quando o suplemento é iniciado no modo de exibição da Web do Internet Explorer. Lembre-se destes pontos adicionais:
>
> - Office na Web abre mais no Internet Explorer. Consequentemente, [o AppSource](/office/dev/store/submit-to-appsource-via-partner-center) não testa mais suplementos no Office na Web usando o Internet Explorer como navegador.
> - O AppSource ainda testa combinações de versões de plataforma e área de  trabalho do Office que usam o Internet Explorer, no entanto, ele só emite um aviso quando o suplemento não dá suporte ao Internet Explorer; o suplemento não é rejeitado pelo AppSource.
> - A [Script Lab não dá](../overview/explore-with-script-lab.md) mais suporte ao Internet Explorer.

Office suplementos são aplicativos Web exibidos dentro de IFrames durante a execução em Office na Web. Office suplementos são exibidos usando controles de navegador inseridos durante a execução em Office no Windows ou Office no Mac. Os controles de navegador inseridos são fornecidos pelo sistema operacional ou por um navegador instalado no computador do usuário.

Se você planeja dar suporte a versões mais antigas do Windows e Office, seu suplemento deve funcionar no controle de navegador inserível baseado no Internet Explorer 11 (IE11). Para obter informações sobre quais combinações de Windows e Office usam o controle de navegador baseado em IE11, consulte Navegadores usados por Office [Suplementos](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> O Internet Explorer 11 não dá suporte a alguns recursos html5, como mídia, gravação e localização. Se o suplemento precisar dar suporte ao Internet Explorer 11, você deverá projetar o suplemento para evitar esses recursos sem suporte ou o suplemento deverá detectar quando o Internet Explorer está sendo usado e fornecer uma experiência alternativa que não usa os recursos sem suporte. Para obter mais informações, [consulte Determinar em runtime se o suplemento está em execução no Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Suporte para versões recentes do JavaScript

O Internet Explorer 11 não dá suporte a versões javaScript posteriores ao ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, ou TypeScript, terá duas opções, conforme descrito neste artigo. Você também pode combinar essas duas técnicas.

### <a name="use-a-transpiler"></a>Usar um transcompilador

Você pode escrever seu código em TypeScript ou JavaScript moderno e transpilá-lo em tempo de compilação para JavaScript ES5. Os arquivos ES5 resultantes são o que você carrega no aplicativo Web do suplemento.

Há dois transcompiladores populares. Ambos podem trabalhar com arquivos de origem que são TypeScript ou JavaScript pós-ES5. Eles também funcionam com React arquivos (.jsx e .tsx).

- [Babel](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

Consulte a documentação de qualquer um deles para obter informações sobre como instalar e configurar o transcompilador em seu projeto de suplemento. Recomendamos que você use um executor de tarefas, como [o Grunt](https://gruntjs.com/) ou [o WebPack](https://webpack.js.org/) , para automatizar a transpilação. Para obter um suplemento de exemplo que usa tsc, [consulte Office Suplemento Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Para obter um exemplo que usa o babel, consulte [o Suplemento Armazenamento Offline](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Se você estiver usando Visual Studio (não Visual Studio Code), o tsc provavelmente será mais fácil de usar. Você pode instalar o suporte para ele com um pacote nuget. Para obter mais informações, [consulte JavaScript e TypeScript no Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Para usar o babel com Visual Studio, crie um script de build ou use o Gerenciador do Executor de Tarefas no Visual Studio com ferramentas como o Executor de Tarefas do [WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou o Executor de Tarefas [NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Usar um polyfill

Um [polyfill é](https://en.wikipedia.org/wiki/Polyfill_(programming)) o JavaScript de versão anterior que duplica a funcionalidade de versões mais recentes do JavaScript. O polyfill funciona com navegadores que não dão suporte a versões posteriores do JavaScript. Por exemplo, o método de `startsWith` cadeia de caracteres não fazia parte da versão ES5 do JavaScript e, portanto, não será executado no Internet Explorer 11. Há bibliotecas de polyfill, escritas em ES5, que definem e implementam um `startsWith` método. Recomendamos a [biblioteca de polyfill core-js](https://github.com/zloirock/core-js) .

Para usar uma biblioteca de polyfill, carregue-a como qualquer outro arquivo ou módulo JavaScript. Por exemplo, `<script>` você pode usar uma marca no arquivo HTML da home page do suplemento ( `<script src="/js/core-js.js"></script>`por exemplo), `import` ou pode usar uma instrução em um arquivo JavaScript (por exemplo, `import 'core-js';`). Quando o mecanismo JavaScript `startsWith`vir um método como , primeiro ele procurará ver se há um método desse nome incorporado à linguagem. Se houver, ele chamará o método nativo. Se, e somente se, o método não for interno, o mecanismo procurará por ele em todos os arquivos carregados. Portanto, a versão polido não é usada em navegadores que dão suporte à versão nativa.

Importar toda a biblioteca core-js importará todos os recursos do core-js. Você também pode importar apenas os polyfills que seu Office suplemento requer. Para obter instruções sobre como fazer isso, consulte [APIs do CommonJS](https://github.com/zloirock/core-js#commonjs-api). A biblioteca core-js tem a maioria dos polyfills de que você precisa. Há algumas exceções detalhadas na seção [Polyfills Ausentes](https://github.com/zloirock/core-js#missing-polyfills) da documentação do core-js. Por exemplo, ele não dá suporte `fetch`, mas você pode usar o [polyfill fetch](https://github.com/github/fetch) .

Para obter um suplemento de exemplo que usa core.js, consulte [o Suplemento do Word Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Determinar em runtime se o suplemento está em execução no Internet Explorer

O suplemento pode descobrir se ele está em execução no Internet Explorer lendo a [propriedade window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Isso permite que o suplemento forneça uma experiência alternativa ou falhe normalmente. Apresentamos um exemplo a seguir. Observe que o Internet Explorer envia uma cadeia de caracteres que começa com "Trident" como o valor de userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> Geralmente, não é uma boa prática ler a `userAgent` propriedade. Verifique se você está familiarizado com o [artigo, detecção](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent) de navegador usando o agente do usuário, incluindo as recomendações e alternativas à leitura `userAgent`. Em particular, se você estiver usando a opção 1 `else` na cláusula acima, considere usar a detecção de recursos em vez de testar para o agente do usuário.
>
> A partir de 30 de setembro de 2021, o texto na seção Qual parte do agente do usuário contém as informações que você está procurando [?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) Datas de antes do lançamento do Internet Explorer 11. Ele ainda é geralmente preciso e *as tabelas* na seção da versão em inglês do artigo estão atualizadas. Da mesma forma, o texto e, na maioria dos casos, as tabelas nas versões não em inglês do artigo estão desatualizadas.

## <a name="test-an-add-in-on-internet-explorer"></a>Testar um suplemento no Internet Explorer

Consulte [o teste do Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Recursos adicionais

- [Tabela de compatibilidade do ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Posso usar... Tabelas de suporte para HTML5, CSS3 etc.](https://caniuse.com/)
