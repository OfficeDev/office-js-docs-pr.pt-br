---
title: Suporte ao Internet Explorer 11
description: Saiba como dar suporte ao Javascript do Internet Explorer 11 e do ES5 no seu complemento.
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: 3677b12d265cb70d2c048e91fc32ff5f9619908b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075905"
---
# <a name="support-internet-explorer-11"></a>Suporte ao Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer ainda usado em Office de complementos**
>
> A Microsoft está encerrando o suporte para o Internet Explorer, mas isso não afeta significativamente Office Desempios. Algumas combinações de plataformas e versões Office, incluindo todas as versões de compra única por meio do Office 2019, continuarão a usar o controle webview que vem com o Internet Explorer 11 para hospedar os complementos, conforme explicado em Navegadores usados por Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, o suporte a essas combinações e, portanto, para o Internet Explorer, ainda é necessário para os complementos enviados ao [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Duas coisas *estão mudando:*
>
> - O AppSource não testa mais os Office na Web usando o Internet Explorer como navegador. Mas o AppSource ainda testa combinações de plataforma e Office *desktop* que usam o Internet Explorer.
> - A [Script Lab de usuário](../overview/explore-with-script-lab.md) para de funcionar no Internet Explorer em algum momento de 2021.

Office Os complementos são aplicativos Web que são exibidos dentro de IFrames ao executar em Office na Web. Office Os complementos são exibidos usando controles de navegador incorporados ao executar Office em Windows ou Office no Mac. Os controles de navegador incorporados são fornecidos pelo sistema operacional ou por um navegador instalado no computador do usuário.

Se você planeja comercializar seu complemento por meio do AppSource ou planeja dar suporte a versões mais antigas do Windows e Office, o seu complemento deve funcionar no controle de navegador in-loca que se baseia no Internet Explorer 11 (IE11). Para obter informações sobre quais combinações de Windows e Office usam o controle de navegador baseado no IE11, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> O Internet Explorer 11 não dá suporte a alguns recursos HTML5, como mídia, gravação e local. Se o seu complemento deve dar suporte ao Internet Explorer 11, não é possível usar esses recursos.

O Internet Explorer 11 não dá suporte a versões JavaScript posteriores ao ES5. Se você quiser usar a sintaxe e os recursos do ECMAScript 2015 ou posterior, ou TypeScript, você tem duas opções, conforme descrito neste artigo. Você também pode combinar essas duas técnicas.

## <a name="use-a-transpiler"></a>Usar um transpiler

Você pode escrever seu código em TypeScript ou JavaScript moderno e transpile-lo no tempo de composição para JavaScript ES5. Os arquivos ES5 resultantes são o que você carrega no aplicativo Web do seu complemento.

Há dois transpiladores populares. Ambos podem trabalhar com arquivos de origem que são TypeScript ou JavaScript pós-ES5. Eles também funcionam com React arquivos (.jsx e .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Consulte a documentação para obter informações sobre como instalar e configurar o transpiler em seu projeto de complemento. Recomendamos que você use um participante de tarefas, como [o Grunhido](https://gruntjs.com/) ou [WebPack,](https://webpack.js.org/) para automatizar a transpilação. Para um exemplo de complemento que usa tsc, consulte [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React). Para um exemplo que usa o babel, consulte [Offline Armazenamento Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Se você estiver usando Visual Studio (não Visual Studio Code), o tsc provavelmente será mais fácil de usar. Você pode instalar o suporte para ele com um pacote nuget. Para obter mais informações, consulte [JavaScript e TypeScript no Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Para usar o babel com Visual Studio, crie um script de complicação ou use o Explorador de Tarefas no Visual Studio com ferramentas como o [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou [o NpM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

## <a name="use-a-polyfill"></a>Usar um polyfill

Um [polyfill é](https://en.wikipedia.org/wiki/Polyfill_(programming)) JavaScript de versão anterior que duplica a funcionalidade de versões mais recentes do JavaScript. O polyfill funciona com navegadores que não suportam as versões javaScript posteriores. Por exemplo, o método string não fazia parte da versão `startsWith` do ES5 do JavaScript e, portanto, não será executado no Internet Explorer 11. Há bibliotecas de polifilamento, escritas no ES5, que definem e implementam um `startsWith` método. Recomendamos a [biblioteca de polifilamento core-js.](https://github.com/zloirock/core-js)

Para usar uma biblioteca de polyfill, carregue-a como qualquer outro arquivo ou módulo JavaScript. Por exemplo, você pode usar uma marca no arquivo HTML da home page do complemento (por exemplo), ou pode usar uma instrução em um arquivo `<script>` `<script src="/js/core-js.js"></script>` `import` JavaScript (por exemplo, `import 'core-js';` ). Quando o mecanismo JavaScript vir um método como , ele procurará primeiro ver se há um método desse nome integrado `startsWith` ao idioma. Se houver, ele chamará o método nativo. Se, e somente se, o método não for integrado, o mecanismo procurará em todos os arquivos carregados para ele. Portanto, a versão poli preenchida não é usada em navegadores que suportam a versão nativa.

Importar toda a biblioteca core-js importará todos os recursos core-js. Você também pode importar apenas os polyfills que seu Office Add-in requer. Para obter instruções sobre como fazer isso, consulte [COMMONJS APIs](https://github.com/zloirock/core-js#commonjs-api). A biblioteca core-js tem a maioria dos polyfills necessários. Há algumas exceções detalhadas na seção [Polyfills Ausentes](https://github.com/zloirock/core-js#missing-polyfills) da documentação core-js. Por exemplo, ele não dá suporte `fetch` , mas você pode usar o [polyfill de](https://github.com/github/fetch) busca.

Para um exemplo de complemento que usa core.js, consulte [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="testing-an-add-in-on-internet-explorer"></a>Testando um complemento no Internet Explorer

Consulte [Teste do Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Recursos adicionais

- [Tabela de compatibilidade do ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Posso usar... Tabelas de suporte para HTML5, CSS3 etc.](https://caniuse.com/)
