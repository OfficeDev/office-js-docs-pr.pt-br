---
title: Depurar suplementos do Office
description: Localize a diretrizes de depuração do Suplemento do Office para seu ambiente de desenvolvimento.
ms.date: 07/11/2022
ms.localizationpriority: high
ms.openlocfilehash: f23e55b2d3ceb84e32365ffbbcb9efafedfebcfc
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423269"
---
# <a name="overview-of-debugging-office-add-ins"></a>Visão geral da depuração de Suplementos do Office

A depuração de Suplementos do Office é essencialmente a mesma que a depuração qualquer aplicativo Web. No entanto, um único conjunto de ferramentas não funcionará para todos os desenvolvedores de suplementos. Isso ocorre porque os suplementos podem ser desenvolvidos em diferentes sistemas operacionais e executados em várias plataformas. Este artigo ajuda você a encontrar as diretrizes de depuração detalhadas para seu ambiente de desenvolvimento.

> [!TIP]
> Este artigo está preocupado com a depuração no sentido estrito de definir pontos de interrupção e percorrer o código. Para obter as diretrizes sobre testes e solução de problemas, comece com [Testar Suplementos do Office](test-debug-office-add-ins.md) e [Solução de problemas de erros de desenvolvimento com Suplementos do Office](troubleshoot-development-errors.md).

> [!NOTE]
> Embora você deva *testar* seu suplemento em todas as plataformas às quais deseja oferecer suporte, você raramente precisará *depurar* em um ambiente diferente do seu computador de desenvolvimento. Por esse motivo, este artigo utiliza “seu computador de desenvolvimento” e “seu ambiente de desenvolvimento” para se referir ao ambiente no qual você está depurando. Se um problema no código ocorrer apenas em uma plataforma diferente daquela em seu computador de desenvolvimento e você precisar definir pontos de interrupção ou percorrer o código para resolvê-lo, o ambiente no qual você está depurando não é literalmente seu ambiente de desenvolvimento.

## <a name="server-side-or-client-side"></a>Do lado do servidor ou do lado do cliente?

Depurar o código do lado do servidor de um suplemento do Office é o mesmo que depurar o lado do servidor de qualquer aplicativo Web. Veja as instruções de depuração do seu IDE ou de outras ferramentas. A seguir estão alguns exemplos de algumas das ferramentas mais populares.

- [Depurar aplicativos ASP.NET ou ASP.NET Core no Visual Studio](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Depuração Expressa](https://expressjs.com/en/guide/debugging.html)
- [Guia de depuração do Node.js](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [Depuração do Node.js no VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Depuração do Webpack](https://webpack.js.org/contribute/debugging/)

O restante deste artigo está preocupado apenas com a depuração do JavaScript do lado do cliente (que pode ser transpilado do TypeScript).

## <a name="special-cases"></a>Casos especiais

Existem alguns casos especiais em que o processo de depuração difere do normal para uma determinada combinação de plataforma, aplicativo do Office e ambiente de desenvolvimento. Se você estiver depurando qualquer um desses casos especiais, use os links nesta seção para encontrar a orientação adequada. Caso contrário, vá para [Orientação geral](#general-guidance).

- **Depurando a função `Office.initialize` ou `Office.onReady`**:[Depure as funções initialize e onReady](debug-initialize-onready.md).
- **Depuração de uma função personalizada do Excel em um ambiente de execução _não compartilhado_**: [Depuração de funções personalizadas em um ambiente de execução não compartilhado](../excel/custom-functions-debugging.md).
- **Depurando um [comando de função](../design/add-in-commands.md#types-of-add-in-commands) em um ambiente de execução _não compartilhado_**: 
    - Suplementos do Outlook em um computador de desenvolvimento Windows: [Comandos de função de depuração em suplementos do Outlook](../outlook/debug-ui-less.md) 
    - Outros suplementos de aplicativos do Office ou Outlook em um computador de desenvolvimento Mac: [Depure um comando de função com um tempo de execução não compartilhado](debug-function-command.md).
- **Depurando um suplemento do Outlook baseado em eventos**: [Depure seu suplemento do Outlook baseado em eventos](../outlook/debug-autolaunch.md). 
 
## <a name="general-guidance"></a>Diretrizes gerais

Para encontrar as diretrizes para depurar o código do lado do cliente, a primeira variável é o sistema operacional do seu computador de desenvolvimento.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux ou outra variante Unix](#debug-on-linux)

### <a name="debug-on-windows"></a>Depurar no Windows

A seguir, as diretrizes gerais para a depuração no Windows. A depuração no Windows depende do seu IDE.

- **Visual Studio**: depurar usando as ferramentas F12 do navegador. Consulte [Depurar Suplementos do Office no Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md).
- **Visual Studio Code**: Depure usando a [Extensão do Depurador de Suplemento para Visual Studio Code](debug-with-vs-extension.md).
- **Qualquer outro IDE** (ou você não quer depurar dentro do seu IDE): use as ferramentas de desenvolvedor associadas ao runtime do navegador que os suplementos usam no seu computador de desenvolvimento. Consulte uma das seguintes opções:

    - [Depurar os suplementos usando as ferramentas de desenvolvedor para o Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor para Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
    - [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](debug-add-ins-using-devtools-edge-chromium.md)

Para obter informações sobre qual runtime está sendo usado, consulte [Navegadores](../concepts/browsers-used-by-office-web-add-ins.md) usados por [Suplementos e Runtimes do Office em Suplementos do Office](runtimes.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="debug-on-mac"></a>Depurar no Mac

Veja a seguir diretrizes gerais para depuração no Mac.

- Se você estiver usando o Visual Studio Code, depure usando a [Extensão do Depurador de Suplemento para Visual Studio Code ](debug-with-vs-extension.md).
- Para qualquer outro IDE, use o Safari Web Inspector. As instruções estão em [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md).


### <a name="debug-on-linux"></a>Depurar no Linux

Não há versões da área de trabalho do Office para Linux, portanto, você precisará [realizar o sideload do suplemento do Office na Web](sideload-office-add-ins-for-testing.md) para testá-lo e depurá-lo. As diretrizes de depuração estão nos [Suplementos de depuração no Office na Web](debug-add-ins-in-office-online.md).

> [!NOTE]
> Não recomendamos que você desenvolva Suplementos do Office em um computador Linux, exceto no caso incomum em que você pode ter certeza de que todos os usuários do suplemento acessarão o suplemento por meio do Office na Web a partir de um computador Linux.

## <a name="debug-add-ins-in-staging-or-production"></a>Depurar suplementos em preparo ou produção

Para depurar um suplemento que já está em preparo ou produção, anexe um depurador da interface do usuário do suplemento. Para obter instruções, [Anexe um depurador no painel de tarefas](attach-debugger-from-task-pane.md).

## <a name="see-also"></a>Confira também

- [Runtimes em Suplementos do Office](runtimes.md)
