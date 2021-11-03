---
title: Suporte ao Internet Explorer 11
description: Saiba como dar suporte ao Javascript do Internet Explorer 11 e do ES5 no seu complemento.
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: a6f762231face1b69a3354b584ca0bbea1742050
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681129"
---
# <a name="support-internet-explorer-11"></a>Suporte ao Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer ainda usado em Office de complementos**
>
> A Microsoft está encerrando o suporte para o Internet Explorer, mas isso não afeta significativamente Office Desempios. Algumas combinações de plataformas e versões Office, incluindo versões de compra única por meio do Office 2019, continuarão a usar o controle webview que vem com o Internet Explorer 11 para hospedar os complementos, conforme explicado em Navegadores usados por Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, o suporte a essas combinações e, portanto, para o Internet Explorer, ainda é necessário para os complementos enviados ao [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Duas coisas *estão mudando:*
>
> - Office na Web abre mais no Internet Explorer. Consequentemente, o AppSource não testa mais os Office na Web usando o Internet Explorer como navegador. Mas o AppSource ainda testa combinações de plataforma e Office *desktop* que usam o Internet Explorer.
> - A [Script Lab não](../overview/explore-with-script-lab.md) dá mais suporte ao Internet Explorer.

Office Os complementos são aplicativos Web que são exibidos dentro de IFrames ao executar em Office na Web. Office Os complementos são exibidos usando controles de navegador incorporados ao executar Office em Windows ou Office no Mac. Os controles de navegador incorporados são fornecidos pelo sistema operacional ou por um navegador instalado no computador do usuário.

Se você planeja comercializar seu complemento por meio do AppSource ou planeja dar suporte a versões mais antigas do Windows e Office, o seu complemento deve funcionar no controle de navegador in-loca que se baseia no Internet Explorer 11 (IE11). Para obter informações sobre quais combinações de Windows e Office usam o controle de navegador baseado no IE11, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> O Internet Explorer 11 não dá suporte a alguns recursos HTML5, como mídia, gravação e local. Se o seu complemento deve oferecer suporte ao Internet Explorer 11, você deve projetar o complemento para evitar esses recursos sem suporte ou o complemento deve detectar quando o Internet Explorer está sendo usado e fornecer uma experiência alternativa que não usa os recursos sem suporte. Para obter mais informações, [consulte Determine at runtime if the add-in is running in Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Suporte para versões recentes do JavaScript

O Internet Explorer 11 não dá suporte a versões JavaScript posteriores ao ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, ou TypeScript, você tem duas opções, conforme descrito neste artigo. Você também pode combinar essas duas técnicas.

### <a name="use-a-transpiler"></a>Usar um transpiler

Você pode escrever seu código em TypeScript ou JavaScript moderno e transpile-lo no tempo de composição para JavaScript ES5. Os arquivos ES5 resultantes são o que você carrega no aplicativo Web do seu complemento.

Há dois transpiladores populares. Ambos podem trabalhar com arquivos de origem que são TypeScript ou JavaScript pós-ES5. Eles também funcionam com React arquivos (.jsx e .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Consulte a documentação para obter informações sobre como instalar e configurar o transpiler em seu projeto de complemento. Recomendamos que você use um participante de tarefas, como [o Grunhido](https://gruntjs.com/) ou [WebPack,](https://webpack.js.org/) para automatizar a transpilação. Para um exemplo de complemento que usa tsc, consulte [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React). Para um exemplo que usa o babel, consulte [Offline Armazenamento Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Se você estiver usando Visual Studio (não Visual Studio Code), o tsc provavelmente será mais fácil de usar. Você pode instalar o suporte para ele com um pacote nuget. Para obter mais informações, consulte [JavaScript e TypeScript no Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Para usar o babel com Visual Studio, crie um script de complicação ou use o Explorador de Tarefas no Visual Studio com ferramentas como o [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou [o NpM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Usar um polyfill

Um [polyfill é](https://en.wikipedia.org/wiki/Polyfill_(programming)) JavaScript de versão anterior que duplica a funcionalidade de versões mais recentes do JavaScript. O polyfill funciona com navegadores que não suportam as versões javaScript posteriores. Por exemplo, o método string não fazia parte da versão `startsWith` do ES5 do JavaScript e, portanto, não será executado no Internet Explorer 11. Há bibliotecas de polifilamento, escritas no ES5, que definem e implementam um `startsWith` método. Recomendamos a [biblioteca de polifilamento core-js.](https://github.com/zloirock/core-js)

Para usar uma biblioteca de polyfill, carregue-a como qualquer outro arquivo ou módulo JavaScript. Por exemplo, você pode usar uma marca no arquivo HTML da home page do complemento (por exemplo), ou pode usar uma instrução em um arquivo `<script>` `<script src="/js/core-js.js"></script>` `import` JavaScript (por exemplo, `import 'core-js';` ). Quando o mecanismo JavaScript vir um método como , ele procurará primeiro ver se há um método desse nome integrado `startsWith` ao idioma. Se houver, ele chamará o método nativo. Se, e somente se, o método não for integrado, o mecanismo procurará em todos os arquivos carregados para ele. Portanto, a versão poli preenchida não é usada em navegadores que suportam a versão nativa.

Importar toda a biblioteca core-js importará todos os recursos core-js. Você também pode importar apenas os polyfills que seu Office Add-in requer. Para obter instruções sobre como fazer isso, consulte [COMMONJS APIs](https://github.com/zloirock/core-js#commonjs-api). A biblioteca core-js tem a maioria dos polyfills necessários. Há algumas exceções detalhadas na seção [Polyfills Ausentes](https://github.com/zloirock/core-js#missing-polyfills) da documentação core-js. Por exemplo, ele não dá suporte `fetch` , mas você pode usar o [polyfill de](https://github.com/github/fetch) busca.

Para um exemplo de complemento que usa core.js, consulte [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Determinar em tempo de execução se o complemento está sendo executado no Internet Explorer

Seu complemento pode descobrir se ele está sendo executado no Internet Explorer lendo a [propriedade window.navigator.userAgent.](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) Isso permite que o complemento forneça uma experiência alternativa ou falhe normalmente. Apresentamos um exemplo a seguir. Observe que o Internet Explorer envia uma cadeia de caracteres começando com "Trident" como o valor de userAgent.

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
> Normalmente, não é uma boa prática ler a `userAgent` propriedade. Certifique-se de que você está familiarizado com o [artigo,](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent)Detecção de navegador usando o agente do usuário , incluindo as recomendações e alternativas à leitura `userAgent` . Em particular, se você estiver tomando a opção 1 na cláusula acima, considere usar a detecção de recursos em vez de `else` testar para o agente do usuário.
>
> A partir de 30 de setembro de 2021, o texto na seção Qual parte do agente do usuário contém as informações que você está [procurando?](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) datas de antes do Internet Explorer 11 ser lançado. Ele ainda é geralmente preciso e as *tabelas* na seção da versão em inglês do artigo estão atualizadas. Da mesma forma, o texto e, na maioria dos casos, as tabelas, nas versões que não são em inglês do artigo, estão desaconsupridas.

## <a name="test-an-add-in-on-internet-explorer"></a>Testar um complemento no Internet Explorer

Consulte [Teste do Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Recursos adicionais

- [Tabela de compatibilidade do ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Posso usar... Tabelas de suporte para HTML5, CSS3 etc.](https://caniuse.com/)
